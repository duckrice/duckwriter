// A box with a list, as Word's font and size boxes: the list opens with every
// choice, the current one marked and in view; typing narrows it down; Up and
// Down move through it, Enter or a click chooses, Esc closes. (The browser's
// own suggestion list only ever shows what matches the box, so with "11" in it
// only 11 is offered.)

import { h } from "../dom";

export interface ComboOptions {
  input: HTMLInputElement;
  /** The choices, in order. */
  options: () => string[];
  /** How a choice is drawn in the list (a font in its own face). */
  style?: (option: string) => string;
  /** A choice was made (the box already shows it). */
  choose: (value: string) => void;
  /** Leaving the box for the text, after a choice or Esc. */
  done: () => void;
}

/** Turns an input into a combo box; returns the element to put in its place. */
export function combo({ input, options, style, choose, done }: ComboOptions): HTMLElement {
  const arrow = h("button", { type: "button", class: "combo-arrow", tabindex: "-1", "aria-label": `Show the list: ${input.title}`, title: input.title });
  arrow.innerHTML = '<svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true"><path d="M2 3.5 5 6.5 8 3.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>';
  const list = h("div", { class: "combo-list", role: "listbox", hidden: true });
  const wrap = h("div", { class: "combo" }, input, arrow, list);
  input.setAttribute("role", "combobox");
  input.setAttribute("aria-autocomplete", "list");
  input.removeAttribute("list");

  let shown: string[] = [];
  let active = -1;
  /** What the box held when the list opened, to tell typing from looking. */
  let opened = "";

  const mark = (i: number) => {
    active = i;
    [...list.children].forEach((el, k) => el.classList.toggle("active", k === i));
    (list.children[i] as HTMLElement | undefined)?.scrollIntoView({ block: "nearest" });
  };
  const draw = (filter: string) => {
    const all = options();
    const q = filter.trim().toLowerCase();
    // Narrowed only by what was typed since the list opened; otherwise everything.
    shown = q ? all.filter((o) => o.toLowerCase().startsWith(q)).concat(all.filter((o) => !o.toLowerCase().startsWith(q) && o.toLowerCase().includes(q))) : all;
    list.replaceChildren(
      ...shown.map((o, i) => {
        const item = h("div", { class: "combo-item", role: "option" }, o);
        if (style) item.setAttribute("style", style(o));
        item.addEventListener("mousedown", (event) => {
          event.preventDefault();
          pick(i);
        });
        return item;
      }),
    );
  };
  const open = () => {
    if (!list.hidden) return;
    opened = input.value;
    draw("");
    list.hidden = false;
    // The current choice marked, and in view.
    const current = shown.findIndex((o) => o.toLowerCase() === input.value.trim().toLowerCase());
    mark(current);
    if (current < 0) list.scrollTop = 0;
  };
  const close = () => {
    list.hidden = true;
    active = -1;
  };
  const pick = (i: number) => {
    const value = shown[i];
    if (value === undefined) return;
    input.value = value;
    close();
    choose(value);
    done();
  };

  arrow.addEventListener("mousedown", (event) => {
    event.preventDefault();
    if (list.hidden) {
      input.focus();
      open();
    } else close();
  });
  input.addEventListener("mousedown", () => {
    // A click in the box opens the list; the text is selected so typing replaces it.
    if (list.hidden) requestAnimationFrame(open);
  });
  input.addEventListener("input", () => {
    if (list.hidden) {
      opened = "";
      list.hidden = false;
    }
    draw(input.value === opened ? "" : input.value);
    mark(input.value.trim() ? 0 : -1);
  });
  input.addEventListener("keydown", (event) => {
    if ((event.key === "ArrowDown" || event.key === "ArrowUp") && !event.ctrlKey) {
      event.preventDefault();
      if (list.hidden) return open();
      const next = active + (event.key === "ArrowDown" ? 1 : -1);
      mark(Math.max(0, Math.min(shown.length - 1, next)));
    } else if (event.key === "Enter") {
      event.preventDefault();
      if (!list.hidden && active >= 0) return pick(active);
      close();
      choose(input.value.trim());
      done();
    } else if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      if (!list.hidden) close();
      done();
    } else if (event.key === "Tab") close();
  });
  input.addEventListener("blur", () => close());
  return wrap;
}
