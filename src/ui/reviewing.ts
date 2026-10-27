// The reviewing pane: every tracked change in the document, in order, as Word's Reviewing Pane:
// who made it and when, what it is, and a button to accept or reject each. Clicking one goes there.

import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { changes, decide, type Change } from "../editor/review";
import { authorClass } from "../editor/schema";

const OPEN_KEY = "duckwriter.reviewing";

export interface ReviewingHost {
  readonly body: EditorView;
  readonly ctx: DocContext;
}

const KIND = { ins: "Inserted", del: "Deleted", fmt: "Formatted" } as const;

/** The words a change is about, shortened. */
function excerpt(view: EditorView, c: Change): string {
  if (c.row) return "a table row";
  if (c.paragraph) return "a paragraph mark (¶)";
  if (c.props) return "the paragraph's formatting";
  const text = view.state.doc.textBetween(c.from, c.to, " ", " ").replace(/\s+/g, " ").trim();
  return text.length > 80 ? `${text.slice(0, 78)}…` : text || "(a picture or a field)";
}

export class ReviewingPane {
  readonly el = h("aside", { class: "nav-pane review-pane", hidden: true, "aria-label": "Tracked changes" });
  private summary = h("div", { class: "review-summary" });
  private list = h("div", { class: "nav-list review-list" });
  private timer: number | undefined;

  constructor(private host: ReviewingHost) {
    this.el.append(h("div", { class: "nav-title" }, "Changes"), this.summary, this.list);
    try {
      this.el.hidden = localStorage.getItem(OPEN_KEY) !== "open";
    } catch {
      // Closed.
    }
  }

  get open(): boolean {
    return !this.el.hidden;
  }

  toggle() {
    this.el.hidden = !this.el.hidden;
    try {
      localStorage.setItem(OPEN_KEY, this.el.hidden ? "closed" : "open");
    } catch {
      // Not remembered.
    }
    this.render();
  }

  /** After a change: lists them again, a moment later while typing. */
  update(docChanged: boolean) {
    if (this.el.hidden) return;
    window.clearTimeout(this.timer);
    if (docChanged) this.timer = window.setTimeout(() => this.render(), 250);
    else this.render();
  }

  render() {
    if (this.el.hidden) return;
    const view = this.host.body;
    const all = changes(view.state.doc);
    const count = (k: Change["kind"]) => all.filter((c) => c.kind === k).length;
    this.summary.textContent = all.length
      ? `${all.length} ${all.length === 1 ? "change" : "changes"}: ${count("ins")} inserted, ${count("del")} deleted, ${count("fmt")} formatted`
      : "No tracked changes.";
    this.list.replaceChildren(
      ...all.map((c) => {
        const decideThis = (accept: boolean) => (event: Event) => {
          event.stopPropagation();
          const tr = decide(view.state, [c], accept, this.host.ctx);
          if (tr) view.dispatch(tr);
          this.render();
        };
        const when = c.date ? new Date(c.date) : null;
        return h(
          "div",
          {
            class: `review-item ${c.kind}`,
            role: "button",
            tabindex: 0,
            onclick: () => {
              const doc = view.state.doc;
              const from = Math.min(c.from, doc.content.size);
              view.dispatch(view.state.tr.setSelection(TextSelection.between(doc.resolve(from), doc.resolve(Math.min(Math.max(from, c.to), doc.content.size)))).scrollIntoView());
              view.focus();
            },
          },
          h(
            "div",
            { class: "review-who" },
            h("span", { class: `review-author ${authorClass(c.author)}` }, c.author || "Someone"),
            " ",
            KIND[c.kind],
            when && !Number.isNaN(when.getTime()) ? h("span", { class: "review-when" }, ` · ${when.toLocaleDateString(undefined, { day: "numeric", month: "short" })}`) : null,
          ),
          h("div", { class: "review-text" }, excerpt(view, c)),
          h(
            "div",
            { class: "review-buttons" },
            h("button", { type: "button", class: "btn small", title: "Accept this change", onclick: decideThis(true) }, "✓ Accept"),
            h("button", { type: "button", class: "btn small", title: "Reject this change", onclick: decideThis(false) }, "✗ Reject"),
          ),
        );
      }),
    );
  }
}
