// The comments, as cards in a column beside the page, each level with the
// text it is about, as in Word. A card's text is edited where it is; replies
// are listed under the comment they answer.

import type { EditorView } from "prosemirror-view";
import { authorName, setAuthorName } from "../author";
import type { CommentData, DocContext } from "../docx/model";
import { h } from "../dom";
import { activeComment, addComment, addReply, commentAnchors, removeComment, reviewKey } from "../editor/review";
import { authorClass } from "../editor/schema";
import { TextSelection } from "prosemirror-state";

export interface CommentsHost {
  readonly body: EditorView;
  readonly ctx: DocContext;
  readonly sheet: HTMLElement;
  readonly zoom: number;
  /** A comment's text, or whether it is resolved, changed. */
  commentsChanged(): void;
}

function when(date: string | null): string {
  const d = date ? new Date(date) : null;
  if (!d || Number.isNaN(d.getTime())) return "";
  const today = new Date();
  return d.toDateString() === today.toDateString()
    ? d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })
    : d.toLocaleDateString(undefined, { day: "numeric", month: "short", year: d.getFullYear() === today.getFullYear() ? undefined : "numeric" });
}

/** A textarea that grows with its text. */
function grow(area: HTMLTextAreaElement) {
  area.style.height = "auto";
  area.style.height = `${area.scrollHeight}px`;
}

export class CommentsPane {
  readonly el = h("div", { class: "comments-pane", "aria-label": "Comments" });
  private cards = new Map<string, HTMLElement>();
  /** What each card shows, to tell when it needs drawing again. */
  private drawn = new WeakMap<HTMLElement, string>();
  /** A comment just added, whose card gets the cursor. */
  private focusNext: string | null = null;

  constructor(private host: CommentsHost) {}

  get visible(): boolean {
    return this.cards.size > 0;
  }

  /** Adds a comment on the selection and puts the cursor in its card. */
  add() {
    const view = this.host.body;
    const { tr, id } = addComment(view.state, this.host.ctx, authorName());
    this.focusNext = id;
    view.dispatch(tr.setMeta(reviewKey, { active: id }));
  }

  private reply(parentId: string, text: string) {
    const view = this.host.body;
    const result = addReply(view.state, this.host.ctx, parentId, authorName());
    if (!result) return;
    this.host.ctx.comments.items.get(result.id)!.text = text;
    view.dispatch(result.tr.setMeta(reviewKey, { active: parentId }));
    this.host.commentsChanged();
  }

  remove(id: string) {
    const view = this.host.body;
    view.dispatch(removeComment(view.state, this.host.ctx, id).setMeta(reviewKey, { active: null }));
    view.focus();
  }

  /** Builds the cards again (keeping one being typed in) and puts them in place. */
  render() {
    const { ctx, body } = this.host;
    const anchors = commentAnchors(body.state.doc);
    const placed = new Set(anchors.map((a) => a.id));
    const top = anchors.filter((a) => {
      const item = ctx.comments.items.get(a.id);
      return item && !(item.parent && placed.has(item.parent));
    });
    const active = activeComment(body.state);
    // Each comment's replies, in the order of the text (once, not for every card).
    const repliesOf = new Map<string, CommentData[]>();
    for (const r of anchors) {
      const item = ctx.comments.items.get(r.id);
      if (item?.parent) (repliesOf.get(item.parent) ?? repliesOf.set(item.parent, []).get(item.parent)!).push(item);
    }
    const keep = new Set<string>();
    for (const a of top) {
      const item = ctx.comments.items.get(a.id)!;
      const replies = repliesOf.get(a.id) ?? [];
      const card = this.card(item, replies, a.id === active);
      keep.add(a.id);
      if (!card.isConnected) this.el.append(card);
    }
    for (const [id, card] of this.cards) {
      if (!keep.has(id)) {
        card.remove();
        this.cards.delete(id);
      }
    }
    // In the order of the text. (Moving a card would take the cursor out of it, so only when needed.)
    top.forEach((a, i) => {
      const card = this.cards.get(a.id);
      if (card && this.el.children[i] !== card) this.el.insertBefore(card, this.el.children[i] ?? null);
    });
    this.host.sheet.classList.toggle("has-comments", this.visible);
    this.layout();
    if (this.focusNext) {
      const area = this.cards.get(this.focusNext)?.querySelector<HTMLTextAreaElement>(".comment-name:not([hidden]) input, .comment-text");
      this.focusNext = null;
      area?.focus();
    }
  }

  private card(item: CommentData, replies: CommentData[], active: boolean): HTMLElement {
    let card = this.cards.get(item.id);
    if (!card) {
      card = h("div", { class: "comment-card", "data-id": item.id });
      card.addEventListener("mousedown", (event) => {
        if ((event.target as HTMLElement).closest("textarea, input, button")) return;
        event.preventDefault();
        this.goTo(item.id);
      });
      this.cards.set(item.id, card);
    }
    card.classList.toggle("active", active);
    card.classList.toggle("done", item.done);
    // Only what isn't being typed in is rebuilt.
    const focused = card.contains(document.activeElement);
    if (focused && card.dataset.replies === String(replies.length)) {
      for (const el of card.querySelectorAll<HTMLElement>(".comment-actions")) el.hidden = false;
      return card;
    }
    card.dataset.replies = String(replies.length);
    // Unchanged since it was drawn: nothing to rebuild (hundreds of comments, at every key).
    const drawn = JSON.stringify([item, ...replies].map((c) => [c.id, c.author, c.date, c.text, c.done]));
    if (this.drawn.get(card) === drawn) return card;
    this.drawn.set(card, drawn);
    const typing = document.activeElement instanceof HTMLTextAreaElement && card.contains(document.activeElement) ? document.activeElement : null;
    card.replaceChildren(this.entry(item, true), ...replies.map((r) => this.entry(r, false)), this.replyBox(item));
    if (typing) (card.querySelector(`[data-entry="${typing.closest<HTMLElement>("[data-entry]")?.dataset.entry}"] .comment-text`) as HTMLElement | null)?.focus();
    return card;
  }

  private entry(item: CommentData, main: boolean): HTMLElement {
    const head = h(
      "div",
      { class: "comment-head" },
      h("span", { class: `comment-avatar ${authorClass(item.author)}`, title: item.author }, item.initials || item.author.slice(0, 1).toUpperCase() || "?"),
      h("span", { class: "comment-author" }, item.author || "You"),
      h("span", { class: "comment-date" }, when(item.date)),
    );
    if (main) {
      const resolve = h("button", { type: "button", class: "comment-button", title: item.done ? "Reopen" : "Resolve" }, item.done ? "Reopen" : "✓");
      resolve.addEventListener("click", () => {
        item.done = !item.done;
        this.host.ctx.comments.version++;
        this.host.commentsChanged();
        this.render();
      });
      const remove = h("button", { type: "button", class: "comment-button", title: "Delete the comment (and its replies)" }, "🗑");
      remove.addEventListener("click", () => this.remove(item.id));
      head.append(h("span", { class: "comment-actions" }, resolve, remove));
    }
    // A new comment asks once for the name to show with comments.
    const name = h("input", { class: "comment-name-input", placeholder: "Your name", "aria-label": "Your name, shown with your comments", value: item.author });
    const nameRow = h("label", { class: "comment-name", hidden: !!item.author }, "Your name", name);
    name.addEventListener("change", () => {
      const value = name.value.trim();
      if (!value) return;
      setAuthorName(value);
      for (const c of this.host.ctx.comments.items.values()) {
        if (!c.author && !c.xml) {
          c.author = value;
          c.initials = value.split(/\s+/).map((w) => w[0]?.toUpperCase() ?? "").join("").slice(0, 3);
        }
      }
      this.host.ctx.comments.version++;
      this.host.commentsChanged();
    });
    name.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        name.dispatchEvent(new Event("change"));
        (name.closest(".comment-entry")?.querySelector(".comment-text") as HTMLElement | null)?.focus();
      }
    });
    const text = h("textarea", { class: "comment-text", rows: 1, placeholder: "Write a comment", "aria-label": `Comment by ${item.author || "you"}` }) as HTMLTextAreaElement;
    text.value = item.text;
    requestAnimationFrame(() => grow(text));
    text.addEventListener("input", () => {
      item.text = text.value;
      this.host.ctx.comments.version++;
      grow(text);
      this.layout();
      this.host.commentsChanged();
    });
    text.addEventListener("focus", () => this.select(item.parent ?? item.id));
    text.addEventListener("keydown", (event) => {
      if (event.key === "Escape" || (event.key === "Enter" && event.ctrlKey)) {
        event.preventDefault();
        event.stopPropagation();
        this.host.body.focus();
      }
    });
    return h("div", { class: `comment-entry${main ? "" : " reply"}`, "data-entry": item.id }, head, nameRow, text);
  }

  private replyBox(item: CommentData): HTMLElement {
    const box = h("textarea", { class: "comment-reply", rows: 1, placeholder: "Reply", "aria-label": "Reply" }) as HTMLTextAreaElement;
    box.addEventListener("input", () => grow(box));
    box.addEventListener("focus", () => this.select(item.id));
    box.addEventListener("keydown", (event) => {
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        const text = box.value.trim();
        if (text) this.reply(item.id, text);
        box.value = "";
      } else if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        this.host.body.focus();
      }
    });
    return h("div", { class: "comment-actions" }, box);
  }

  /** Makes a comment the one looked at, without moving the cursor. */
  private select(id: string) {
    const view = this.host.body;
    if (activeComment(view.state) !== id) view.dispatch(view.state.tr.setMeta(reviewKey, { active: id }));
  }

  /** Puts the cursor at the start of a comment's text. */
  goTo(id: string) {
    const view = this.host.body;
    const a = commentAnchors(view.state.doc).find((x) => x.id === id);
    if (!a) return;
    const pos = Math.min(a.from + 1, view.state.doc.content.size);
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(pos))).setMeta(reviewKey, { active: id }).scrollIntoView());
    view.focus();
  }

  /** Each card level with its text, without overlapping the one above. */
  layout() {
    const sheet = this.host.sheet;
    if (!this.cards.size) return;
    const zoom = this.host.zoom || 1;
    const origin = sheet.getBoundingClientRect().top;
    // Each comment's place in the text: where it starts, or else any of its marks; found in one go.
    const anchors = new Map<string, Element>();
    for (const el of sheet.querySelectorAll(".doc .comment-anchor[data-id]")) {
      const id = (el as HTMLElement).dataset.id!;
      if (!anchors.has(id) || el.classList.contains("comment-start")) anchors.set(id, el);
    }
    // Everything measured first, then everything moved: no layout in between.
    const cards = [...(this.el.children as HTMLCollectionOf<HTMLElement>)];
    const wants = cards.map((card) => {
      const anchor = anchors.get(card.dataset.id!);
      return anchor ? (anchor.getBoundingClientRect().top - origin) / zoom - 6 : null;
    });
    const heights = cards.map((card) => card.offsetHeight);
    let bottom = 0;
    cards.forEach((card, i) => {
      const top = Math.max(wants[i] ?? bottom, bottom);
      card.style.top = `${top}px`;
      bottom = top + heights[i] + 8;
    });
  }
}

/** Asks once for the name shown with comments and tracked changes. Null if cancelled. */
export function askName(why: string): Promise<string | null> {
  const known = authorName();
  if (known) return Promise.resolve(known);
  return new Promise((resolve) => {
    const input = h("input", { class: "tb-input dialog-input", placeholder: "Your name", "aria-label": "Your name" }) as HTMLInputElement;
    let answer: string | null = null;
    const ok = () => {
      const name = input.value.trim();
      if (!name) return input.focus();
      answer = name;
      setAuthorName(name);
      dialog.close();
    };
    const dialog = h(
      "dialog",
      { class: "dialog" },
      h("h2", {}, "Your name"),
      h("p", { class: "dialog-hint" }, `${why} It's remembered on this computer.`),
      h("div", { class: "dialog-form" }, h("label", {}, "Name", input)),
      h(
        "div",
        { class: "dialog-footer" },
        h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
        h("button", { type: "button", class: "btn primary", onclick: ok }, "OK"),
      ),
    );
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") ok();
    });
    dialog.addEventListener("close", () => {
      dialog.remove();
      resolve(answer);
    });
    document.body.append(dialog);
    dialog.showModal();
    input.focus();
  });
}
