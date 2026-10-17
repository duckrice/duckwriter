// The editing area: a ProseMirror view. It handles the keys that are part of
// typing (Enter, Tab, Backspace…); shortcuts are in ui/shortcuts.ts.

import { baseKeymap, chainCommands } from "prosemirror-commands";
import { dropCursor } from "prosemirror-dropcursor";
import { gapCursor } from "prosemirror-gapcursor";
import { history } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import type { Node } from "prosemirror-model";
import { EditorState, Plugin, type Command } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { insertTab, leaveListAtStart, newParagraph } from "./commands";
import { changeListLevel, listPlugin } from "./lists";
import { fieldsPlugin } from "./fields";
import { ImageView, type PictureHooks } from "./imageview";
import { MathView } from "./mathview";
import { paginationPlugin, type PaginationHooks } from "./pagination";
import { editEquation } from "../ui/equation";

export interface EditorHooks {
  /** Called after every change to the document or the selection. */
  onChange(view: EditorView, docChanged: boolean): void;
  /** Given for the main text, where pictures can be moved and resized. */
  pictures?: PictureHooks;
  /** Picture files pasted or dropped into the text. */
  onImageFiles?(files: File[], view: EditorView, pos: number | null): void;
}

/** @param pages given for the main text, which is laid out on pages; headers, footers and notes aren't. */
export function createState(doc: Node, ctx: () => DocContext, pages?: Omit<PaginationHooks, "context">): EditorState {
  const c = (make: (ctx: DocContext) => Command): Command => (state, dispatch, view) => make(ctx())(state, dispatch, view);
  return EditorState.create({
    doc,
    plugins: [
      history(),
      keymap({
        // Typing keys. Every other shortcut is in ui/shortcuts.ts, which also lists them (F1).
        "Shift-Enter": chainCommands((state, dispatch) => {
          if (dispatch) dispatch(state.tr.replaceSelectionWith(state.schema.nodes.hard_break.create()).scrollIntoView());
          return true;
        }),
        Enter: c((x) => newParagraph(x)),
        Tab: chainCommands(c((x) => changeListLevel(x, 1)), insertTab()),
        "Shift-Tab": c((x) => changeListLevel(x, -1)),
        Backspace: c((x) => leaveListAtStart(x)),
      }),
      keymap(baseKeymap),
      listPlugin(ctx),
      fieldsPlugin(),
      ...(pages ? [paginationPlugin({ context: ctx, ...pages })] : []),
      dropCursor({ color: "#0f766e", width: 2 }),
      gapCursor(),
      new Plugin({
        props: {
          // Links open in the browser only with Ctrl held, as in Word.
          handleClick(_view, _pos, event) {
            const link = (event.target as HTMLElement).closest("a");
            if (link && !(event.ctrlKey || event.metaKey)) event.preventDefault();
            return false;
          },
        },
      }),
    ],
  });
}

/** @param className "doc" for the main text, "hf-doc" for a header or footer. */
export function createView(host: HTMLElement, state: EditorState, hooks: EditorHooks, className = "doc"): EditorView {
  const view: EditorView = new EditorView(host, {
    state,
    attributes: { spellcheck: "true", class: className },
    nodeViews: {
      math: (node, v, getPos) => new MathView(node, v, getPos, (pos) => editEquation(v, pos)),
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
