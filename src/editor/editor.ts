// The editing area: a ProseMirror view with Word's keyboard shortcuts.

import { baseKeymap, chainCommands, selectAll } from "prosemirror-commands";
import { dropCursor } from "prosemirror-dropcursor";
import { gapCursor } from "prosemirror-gapcursor";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import type { Node } from "prosemirror-model";
import { EditorState, Plugin, type Command } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import {
  clearFormatting,
  insertPageBreak,
  insertTab,
  leaveListAtStart,
  newParagraph,
  setAlign,
  setLineSpacing,
  setStyle,
  stepSize,
  toggleFormat,
  toggleVertAlign,
} from "./commands";
import { changeListLevel, listPlugin, toggleList } from "./lists";
import { MathView } from "./mathview";
import { paginationPlugin, type PageLayout } from "./pagination";
import { editEquation, insertEquation } from "../ui/equation";

export interface EditorHooks {
  /** Called after every change to the document or the selection. */
  onChange(view: EditorView, docChanged: boolean): void;
}

/** Applies a heading style if the document has it. */
const heading = (ctx: () => DocContext, level: number): Command => (state, dispatch) => {
  const id = `Heading${level}`;
  if (!ctx().styles.byId.has(id)) return false;
  return setStyle(ctx(), id)(state, dispatch);
};

export function createState(doc: Node, ctx: () => DocContext, onLayout: (layout: PageLayout) => void = () => {}): EditorState {
  const c = (make: (ctx: DocContext) => Command): Command => (state, dispatch, view) => make(ctx())(state, dispatch, view);
  return EditorState.create({
    doc,
    plugins: [
      history(),
      keymap({
        "Mod-z": undo,
        "Mod-y": redo,
        "Mod-Shift-z": redo,
        "Mod-b": c((x) => toggleFormat(x, "bold")),
        "Mod-i": c((x) => toggleFormat(x, "italic")),
        "Mod-u": c((x) => toggleFormat(x, "underline")),
        "Mod-=": toggleVertAlign("subscript"),
        "Mod-Shift-=": toggleVertAlign("superscript"),
        "Mod-Shift-+": toggleVertAlign("superscript"),
        "Mod-l": setAlign("left"),
        "Mod-e": setAlign("center"),
        "Mod-r": setAlign("right"),
        "Mod-j": setAlign("justify"),
        "Mod-1": setLineSpacing(1),
        "Mod-2": setLineSpacing(2),
        "Mod-5": setLineSpacing(1.5),
        "Mod-Alt-1": heading(ctx, 1),
        "Mod-Alt-2": heading(ctx, 2),
        "Mod-Alt-3": heading(ctx, 3),
        "Mod-Shift-n": c((x) => setStyle(x, null)),
        "Mod-Space": c((x) => clearFormatting(x)),
        "Mod-]": c((x) => stepSize(x, 1)),
        "Mod-[": c((x) => stepSize(x, -1)),
        "Mod-Shift->": c((x) => stepSize(x, 1)),
        "Mod-Shift-<": c((x) => stepSize(x, -1)),
        "Mod-Shift-l": c((x) => toggleList(x, "bullet")),
        "Mod-Enter": insertPageBreak(),
        // Word's shortcut for a new equation.
        "Alt-=": (_state, _dispatch, view) => {
          if (view) insertEquation(view);
          return true;
        },
        "Shift-Enter": chainCommands((state, dispatch) => {
          if (dispatch) dispatch(state.tr.replaceSelectionWith(state.schema.nodes.hard_break.create()).scrollIntoView());
          return true;
        }),
        Enter: c((x) => newParagraph(x)),
        Tab: chainCommands(c((x) => changeListLevel(x, 1)), insertTab()),
        "Shift-Tab": c((x) => changeListLevel(x, -1)),
        Backspace: c((x) => leaveListAtStart(x)),
        "Mod-a": selectAll,
      }),
      keymap(baseKeymap),
      listPlugin(ctx),
      paginationPlugin({ context: ctx, onLayout }),
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

export function createView(host: HTMLElement, state: EditorState, hooks: EditorHooks): EditorView {
  const view: EditorView = new EditorView(host, {
    state,
    attributes: { spellcheck: "true", class: "doc" },
    nodeViews: {
      math: (node, v, getPos) => new MathView(node, v, getPos, (pos) => editEquation(v, pos)),
    },
    dispatchTransaction(tr) {
      view.updateState(view.state.apply(tr));
      hooks.onChange(view, tr.docChanged);
    },
  });
  return view;
}
