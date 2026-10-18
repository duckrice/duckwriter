// The editing area: a ProseMirror view. It handles the keys that are part of
// typing (Enter, Tab, Backspace…); shortcuts are in ui/shortcuts.ts.

import { baseKeymap, chainCommands } from "prosemirror-commands";
import { dropCursor } from "prosemirror-dropcursor";
import { gapCursor } from "prosemirror-gapcursor";
import { history } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import type { Node } from "prosemirror-model";
import { EditorState, Plugin, TextSelection, type Command, type Transaction } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { addRowAfter, columnResizing, goToNextCell, isInTable, tableEditing } from "prosemirror-tables";
import { insertTab, leaveListAtStart, newParagraph } from "./commands";
import { DuckTableView } from "./tableview";
import { changeListLevel, listPlugin } from "./lists";
import { fieldsPlugin } from "./fields";
import { findPlugin } from "./find";
import { spellingPlugin } from "./spelling";
import { reviewPlugin } from "./review";
import { trackPlugin } from "./track";
import { pastePlugin } from "./paste";
import { autoFormatPlugin } from "./autoformat";
import { citationsPlugin } from "./citations";
import { authorName } from "../author";
import { ImageView, type PictureHooks } from "./imageview";
import { MathView } from "./mathview";
import { paginationPlugin, sectionWidthPlugin, type PaginationHooks } from "./pagination";
import { schema } from "./schema";
import { editEquation } from "../ui/equation";
import { EndnotesView, editEndnote } from "../ui/endnotes";

export interface EditorHooks {
  /** Called after every change to the document or the selection. */
  onChange(view: EditorView, docChanged: boolean): void;
  /** Given for the main text, where pictures can be moved and resized. */
  pictures?: PictureHooks;
  /** Picture files pasted or dropped into the text. */
  onImageFiles?(files: File[], view: EditorView, pos: number | null): void;
  /** Given for the main text, where the endnotes are listed at the end. */
  context?: () => DocContext;
}

let linkOpener: ((href: string, anchor: string | null) => void) | null = null;

/** What following a link (Ctrl+click) does. */
export function setLinkOpener(open: typeof linkOpener) {
  linkOpener = open;
}

/** @param pages given for the main text, which is laid out on pages; headers, footers and notes aren't. */
/** @param extra more plugins: writing together adds its own. */
export function createState(doc: Node, ctx: () => DocContext, pages?: Omit<PaginationHooks, "context">, extra: Plugin[] = []): EditorState {
  const c = (make: (ctx: DocContext) => Command): Command => (state, dispatch, view) => make(ctx())(state, dispatch, view);
  return EditorState.create({
    doc,
    plugins: [
      autoFormatPlugin(ctx),
      history(),
      keymap({
        // Typing keys. Every other shortcut is in ui/shortcuts.ts, which also lists them (F1).
        "Shift-Enter": chainCommands((state, dispatch) => {
          if (dispatch) dispatch(state.tr.replaceSelectionWith(state.schema.nodes.hard_break.create()).scrollIntoView());
          return true;
        }),
        Enter: c((x) => newParagraph(x)),
        // In a table, Tab goes to the next cell; in the last one it adds a row, as in Word.
        Tab: chainCommands(
          goToNextCell(1),
          (state, dispatch) => {
            if (!isInTable(state)) return false;
            let added: Transaction | null = null;
            addRowAfter(state, (tr) => (added = tr));
            if (!added) return false;
            const tr: Transaction = added;
            // Into the first cell of the new row, as one step.
            const $from = state.selection.$from;
            let depth = $from.depth;
            while (depth > 0 && $from.node(depth).type.name !== "table_row") depth--;
            // Before the inserted row, not after it.
            const rowEnd = tr.mapping.map($from.after(depth), -1);
            tr.setSelection(TextSelection.near(tr.doc.resolve(rowEnd + 1)));
            if (dispatch) dispatch(tr.scrollIntoView());
            return true;
          },
          c((x) => changeListLevel(x, 1)),
          insertTab(),
        ),
        "Shift-Tab": chainCommands(goToNextCell(-1), c((x) => changeListLevel(x, -1))),
        Backspace: c((x) => leaveListAtStart(x)),
      }),
      keymap(baseKeymap),
      columnResizing({ View: DuckTableView, cellMinWidth: 30 }),
      tableEditing(),
      listPlugin(ctx),
      fieldsPlugin(),
      findPlugin(),
      spellingPlugin(),
      reviewPlugin(ctx),
      trackPlugin({ on: () => !!ctx().trackChanges, author: authorName, context: ctx }),
      pastePlugin(ctx),
      citationsPlugin(ctx),
      ...(pages ? [paginationPlugin({ context: ctx, ...pages }), sectionWidthPlugin(ctx)] : []),
      dropCursor({ color: "#0f766e", width: 2 }),
      gapCursor(),
      new Plugin({
        props: {
          handleDOMEvents: {
            // Links are followed only with Ctrl held, as in Word, and never inside the app's own window.
            click(view, event) {
              const endnote = (event.target as HTMLElement).closest<HTMLElement>(".endnote-ref");
              if (endnote?.dataset.id) {
                editEndnote(view, ctx(), endnote.dataset.id);
                return true;
              }
              const a = (event.target as HTMLElement).closest("a");
              if (!a) return false;
              event.preventDefault();
              if (!(event.ctrlKey || event.metaKey)) return false;
              const pos = view.posAtDOM(a, 0);
              const mark = schema.marks.link.isInSet(view.state.doc.resolve(Math.min(pos + 1, view.state.doc.content.size)).marks());
              if (mark) linkOpener?.(mark.attrs.href, mark.attrs.anchor);
              return true;
            },
          },
        },
      }),
      ...extra,
    ],
  });
}

/** @param className "doc" for the main text, "hf-doc" for a header or footer. */
export function createView(host: HTMLElement, state: EditorState, hooks: EditorHooks, className = "doc"): EditorView {
  const view: EditorView = new EditorView(host, {
    state,
    attributes: { spellcheck: "false", class: className },
    nodeViews: {
      math: (node, v, getPos) => new MathView(node, v, getPos, (pos) => editEquation(v, pos)),
      ...(hooks.context ? { endnotes: (node: Node, v: EditorView) => new EndnotesView(node, v, hooks.context!) } : {}),
      ...(hooks.pictures ? { image: (node: Node, v: EditorView, getPos: () => number | undefined) => new ImageView(node, v, getPos, hooks.pictures!) } : {}),
    },
    handlePaste(v, event) {
      const files = [...(event.clipboardData?.files ?? [])].filter((f) => f.type.startsWith("image/"));
      if (!files.length || !hooks.onImageFiles) return false;
      hooks.onImageFiles(files, v, null);
      return true;
    },
    handleDrop(v, event) {
      const files = [...(event.dataTransfer?.files ?? [])].filter((f) => f.type.startsWith("image/"));
      if (!files.length || !hooks.onImageFiles) return false;
      event.preventDefault();
      hooks.onImageFiles(files, v, v.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ?? null);
      return true;
    },
    dispatchTransaction(tr) {
      view.updateState(view.state.apply(tr));
      hooks.onChange(view, tr.docChanged);
    },
  });
  return view;
}
