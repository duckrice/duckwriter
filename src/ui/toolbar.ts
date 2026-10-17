// The formatting toolbar. It shows the formatting at the cursor and applies
// changes through the editor's commands.

import { redo, undo, undoDepth, redoDepth } from "prosemirror-history";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import { HIGHLIGHTS, type DocContext } from "../docx/model";
import {
  clearFormatting,
  formattingAt,
  insertPageBreak,
  setAlign,
  setLineSpacing,
  setMark,
  setParagraphs,
  setStyle,
  stepIndent,
  stepSize,
  toggleFormat,
  toggleVertAlign,
} from "../editor/commands";
import { changeListLevel, kindOf, listOf, toggleList } from "../editor/lists";
import { schema } from "../editor/schema";
import { insertEquation } from "./equation";
import { icon, type IconName } from "./icons";

export interface ToolbarHost {
  readonly view: EditorView;
  readonly ctx: DocContext;
  newDocument(): void;
  open(): void;
  save(): void;
  editHeader(): void;
}

const FONTS = [
  "Aptos", "Arial", "Arial Black", "Book Antiqua", "Calibri", "Calibri Light", "Cambria", "Candara", "Century Gothic",
  "Comic Sans MS", "Consolas", "Constantia", "Corbel", "Courier New", "Franklin Gothic Medium", "Garamond", "Georgia",
  "Impact", "Lucida Console", "Palatino Linotype", "Segoe UI", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana",
];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** Word's standard colours. */
const COLORS = [
  "#000000", "#404040", "#7f7f7f", "#bfbfbf", "#ffffff",
  "#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050",
  "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0",
];
const HIGHLIGHT_CHOICES = ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black"];

export class Toolbar {
  readonly el: HTMLElement;
  private buttons = new Map<string, HTMLButtonElement>();
  private styleSelect = h("select", { class: "tb-select style-select", title: "Style" });
  private fontInput = h("input", { class: "tb-input font-input", title: "Font", list: "dw-fonts", spellcheck: "false" });
  private sizeInput = h("input", { class: "tb-input size-input", title: "Font size", list: "dw-sizes", inputmode: "decimal" });
  private textColorBar = h("span", { class: "color-bar" });
  private highlightBar = h("span", { class: "color-bar" });
  private lastTextColor = "#c00000";
  private lastHighlight = "yellow";
  private openMenu: HTMLElement | null = null;

  constructor(private host: ToolbarHost) {
    const fonts = h("datalist", { id: "dw-fonts" }, ...FONTS.map((f) => h("option", { value: f })));
    const sizes = h("datalist", { id: "dw-sizes" }, ...SIZES.map((s) => h("option", { value: String(s) })));

    this.styleSelect.addEventListener("change", () => {
      this.run(setStyle(this.host.ctx, this.styleSelect.value || null));
    });
    this.fontInput.addEventListener("change", () => {
      const family = this.fontInput.value.trim();
      if (family) this.run(setMark(schema.marks.font, { family }));
    });
    this.sizeInput.addEventListener("change", () => {
      const size = Number(this.sizeInput.value.replace(",", "."));
      if (size >= 1 && size <= 1638) this.run(setMark(schema.marks.size, { pt: Math.round(size * 2) / 2 }));
    });
    // Choosing from the list should apply straight away and give the text back the focus.
    for (const input of [this.fontInput, this.sizeInput]) {
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          input.dispatchEvent(new Event("change"));
          this.host.view.focus();
        } else if (event.key === "Escape") this.host.view.focus();
      });
      input.addEventListener("focus", () => input.select());
    }

    const ctx = () => this.host.ctx;
    const group = (...items: HTMLElement[]) => h("div", { class: "tb-group" }, ...items);
    this.el = h(
      "div",
      { class: "toolbar", role: "toolbar" },
      group(
        this.button("newDoc", "New document (Ctrl+N)", () => this.host.newDocument()),
        this.button("open", "Open (Ctrl+O)", () => this.host.open()),
        this.button("save", "Save (Ctrl+S)", () => this.host.save()),
      ),
      group(this.button("undo", "Undo (Ctrl+Z)", () => this.run(undo)), this.button("redo", "Redo (Ctrl+Y)", () => this.run(redo))),
      group(this.styleSelect),
      group(
        this.fontInput,
        this.sizeInput,
        this.button("zoomIn", "Bigger text (Ctrl+])", () => this.run(stepSize(ctx(), 1)), "sizeUp"),
        this.button("zoomOut", "Smaller text (Ctrl+[)", () => this.run(stepSize(ctx(), -1)), "sizeDown"),
      ),
      group(
        this.button("bold", "Bold (Ctrl+B)", () => this.run(toggleFormat(ctx(), "bold"))),
        this.button("italic", "Italic (Ctrl+I)", () => this.run(toggleFormat(ctx(), "italic"))),
        this.button("underline", "Underline (Ctrl+U)", () => this.run(toggleFormat(ctx(), "underline"))),
        this.button("strike", "Strikethrough", () => this.run(toggleFormat(ctx(), "strike"))),
        this.button("superscript", "Superscript (Ctrl+Shift++)", () => this.run(toggleVertAlign("superscript"))),
        this.button("subscript", "Subscript (Ctrl+=)", () => this.run(toggleVertAlign("subscript"))),
      ),
      group(
        this.splitButton("textColor", "Text colour", this.textColorBar, () => this.run(setMark(schema.marks.color, { color: this.lastTextColor })), () => this.colorMenu()),
        this.splitButton("highlight", "Highlight", this.highlightBar, () => this.run(setMark(schema.marks.highlight, { color: this.lastHighlight })), () => this.highlightMenu()),
        this.button("clearFormat", "Clear formatting (Ctrl+Space)", () => this.run(clearFormatting(ctx()))),
      ),
      group(
        this.button("alignLeft", "Align left (Ctrl+L)", () => this.run(setAlign("left"))),
        this.button("alignCenter", "Centre (Ctrl+E)", () => this.run(setAlign("center"))),
        this.button("alignRight", "Align right (Ctrl+R)", () => this.run(setAlign("right"))),
        this.button("alignJustify", "Justify (Ctrl+J)", () => this.run(setAlign("justify"))),
      ),
      group(
        this.button("bullets", "Bullets", () => this.run(toggleList(ctx(), "bullet"))),
        this.button("numbers", "Numbering", () => this.run(toggleList(ctx(), "number"))),
        this.button("indentLess", "Decrease indent (Shift+Tab in a list)", () => this.indent(-1)),
        this.button("indentMore", "Increase indent (Tab in a list)", () => this.indent(1)),
        this.splitButton("lineSpacing", "Line and paragraph spacing", null, () => this.spacingMenu(), () => this.spacingMenu()),
      ),
      group(
        this.button("equation", "Equation (Alt+=)", () => insertEquation(this.host.view)),
        this.splitButton("symbol", "Symbol", null, () => this.symbolMenu(), () => this.symbolMenu()),
        this.button("pageBreak", "Page break (Ctrl+Enter)", () => this.run(insertPageBreak())),
        this.button("header", "Header and footer", () => this.host.editHeader()),
      ),
      fonts,
      sizes,
    );
    this.textColorBar.style.background = this.lastTextColor;
    this.highlightBar.style.background = HIGHLIGHTS[this.lastHighlight];
    document.addEventListener("pointerdown", (event) => {
      if (this.openMenu && !this.openMenu.contains(event.target as Node) && !(event.target as HTMLElement).closest(".split-arrow")) this.closeMenu();
    });
  }

  /** Runs a command on the editor and gives it the focus back. */
  run(command: Command) {
    const { view } = this.host;
    command(view.state, view.dispatch, view);
    view.focus();
  }

  private indent(direction: 1 | -1) {
    const ctx = this.host.ctx;
    const { view } = this.host;
    if (!changeListLevel(ctx, direction)(view.state, view.dispatch)) this.run(stepIndent(ctx, direction));
    view.focus();
  }

  private button(name: IconName, title: string, onClick: () => void, key: string = name): HTMLButtonElement {
    const button = h("button", { type: "button", class: "tb", title, "aria-label": title, html: icon(name) });
    // Keep the editor's selection: the button must not take the focus.
    button.addEventListener("mousedown", (event) => event.preventDefault());
    button.addEventListener("click", onClick);
    this.buttons.set(key, button);
    return button;
  }

  private splitButton(name: IconName, title: string, bar: HTMLElement | null, onMain: () => void, onMenu: () => HTMLElement): HTMLElement {
    const main = this.button(name, title, onMain);
    if (bar) main.append(bar);
    const arrow = h("button", { type: "button", class: "tb split-arrow", title: `${title} options`, "aria-label": `${title} options`, html: icon("caret", 14) });
    arrow.addEventListener("mousedown", (event) => event.preventDefault());
    const wrap = h("div", { class: "split" }, main, arrow);
    arrow.addEventListener("click", () => {
      if (this.openMenu?.dataset.owner === name) return this.closeMenu();
      this.closeMenu();
      const menu = onMenu();
      menu.dataset.owner = name;
      menu.addEventListener("mousedown", (event) => {
        if (!(event.target instanceof HTMLInputElement)) event.preventDefault();
      });
      wrap.append(menu);
      this.openMenu = menu;
    });
    return wrap;
  }

  private closeMenu() {
    this.openMenu?.remove();
    this.openMenu = null;
  }

  private swatches(colors: string[], css: (c: string) => string, pick: (c: string) => void, labelFor: (c: string) => string) {
    return h(
      "div",
      { class: "swatches" },
      ...colors.map((color) => {
        const swatch = h("button", { type: "button", class: "swatch", title: labelFor(color), "aria-label": labelFor(color) });
        swatch.style.background = css(color);
        swatch.addEventListener("click", () => {
          pick(color);
          this.closeMenu();
        });
        return swatch;
      }),
    );
  }

  private colorMenu(): HTMLElement {
    const pick = (color: string) => {
      this.lastTextColor = color;
      this.textColorBar.style.background = color;
      this.run(setMark(schema.marks.color, { color }));
    };
    const custom = h("input", { type: "color", class: "custom-color", title: "More colours", value: this.lastTextColor });
    custom.addEventListener("change", () => {
      pick(custom.value);
      this.closeMenu();
    });
    const automatic = h("button", { type: "button", class: "menu-item" }, "Automatic");
    automatic.addEventListener("click", () => {
      this.run(setMark(schema.marks.color, null));
      this.closeMenu();
    });
    return h("div", { class: "menu" }, automatic, this.swatches(COLORS, (c) => c, pick, (c) => c), h("label", { class: "menu-item" }, custom, "More colours…"));
  }

  private highlightMenu(): HTMLElement {
    const none = h("button", { type: "button", class: "menu-item" }, "No highlight");
    none.addEventListener("click", () => {
      this.run(setMark(schema.marks.highlight, null));
      this.closeMenu();
    });
    const pick = (name: string) => {
      this.lastHighlight = name;
      this.highlightBar.style.background = HIGHLIGHTS[name];
      this.run(setMark(schema.marks.highlight, { color: name }));
    };
    return h("div", { class: "menu" }, this.swatches(HIGHLIGHT_CHOICES, (c) => HIGHLIGHTS[c], pick, (c) => c.replace(/([A-Z])/g, " $1").toLowerCase()), none);
  }

  private symbolMenu(): HTMLElement {
    const groups: [string, string][] = [
      ["Greek", "α β γ δ ε ζ η θ ι κ λ μ ν ξ π ρ σ τ υ φ χ ψ ω Γ Δ Θ Λ Ξ Π Σ Φ Ψ Ω"],
      ["Maths", "± ∓ × ÷ ⋅ ≤ ≥ ≠ ≈ ≡ ∼ ∝ ∞ ∂ ∇ √ ∑ ∏ ∫ ∮ ∈ ∉ ⊂ ⊆ ∪ ∩ ∅ ∀ ∃ ¬ ∧ ∨ ∴ ′ ″"],
      ["Arrows", "→ ← ↔ ⇒ ⇐ ⇔ ↑ ↓ ⇌ ↦"],
      ["Units and more", "° ‰ µ Å Ω ℃ ℉ ½ ⅓ ¼ ¾ ² ³ ¹ ⁰ ⁻ ₀ ₁ ₂ ₃ € £ © ® ™ § ¶ † ‡ • … – —"],
    ];
    const insert = (text: string) => {
      const { view } = this.host;
      view.dispatch(view.state.tr.insertText(text).scrollIntoView());
      view.focus();
    };
    return h(
      "div",
      { class: "menu symbol-menu" },
      ...groups.flatMap(([title, chars]) => [
        h("div", { class: "menu-heading" }, title),
        h(
          "div",
          { class: "symbol-grid" },
          ...chars.split(" ").map((ch) => {
            const b = h("button", { type: "button", class: "symbol-button", title: ch }, ch);
            b.addEventListener("click", () => insert(ch));
            return b;
          }),
        ),
      ]),
    );
  }

  private spacingMenu(): HTMLElement {
    const current = formattingAt(this.host.ctx, this.host.view.state).para;
    const item = (label: string, active: boolean, onClick: () => void) => {
      const button = h("button", { type: "button", class: `menu-item${active ? " checked" : ""}` }, label);
      button.addEventListener("click", () => {
        onClick();
        this.closeMenu();
      });
      return button;
    };
    const lines = [1, 1.15, 1.5, 2, 2.5, 3].map((m) =>
      item(m.toFixed(m % 1 ? 2 : 1).replace(/0$/, ""), current.lineSpacing !== undefined && Math.abs(current.lineSpacing - m) < 0.01, () => this.run(setLineSpacing(m))),
    );
    const before = (current.spaceBefore ?? 0) > 0;
    const after = (current.spaceAfter ?? 0) > 0;
    return h(
      "div",
      { class: "menu" },
      ...lines,
      h("div", { class: "menu-separator" }),
      item(before ? "Remove space before paragraph" : "Add space before paragraph", false, () => this.run(setParagraphs(() => ({ spaceBefore: before ? 0 : 12 })))),
      item(after ? "Remove space after paragraph" : "Add space after paragraph", false, () => this.run(setParagraphs(() => ({ spaceAfter: after ? 0 : 8 })))),
    );
  }

  /** Shows the formatting at the cursor. */
  update() {
    const { view, ctx } = this.host;
    const { run, para, paragraph } = formattingAt(ctx, view.state);
    const active = (key: string, on: boolean | undefined) => this.buttons.get(key)?.classList.toggle("active", !!on);
    active("bold", run.bold);
    active("italic", run.italic);
    active("underline", run.underline);
    active("strike", run.strike);
    active("superscript", run.vertAlign === "superscript");
    active("subscript", run.vertAlign === "subscript");
    const align = para.align ?? "left";
    active("alignLeft", align === "left");
    active("alignCenter", align === "center");
    active("alignRight", align === "right");
    active("alignJustify", align === "justify");
    const list = paragraph ? listOf(ctx, paragraph) : null;
    const kind = list ? kindOf(ctx, list.numId, list.ilvl) : null;
    active("bullets", kind === "bullet");
    active("numbers", kind === "number");
    this.buttons.get("undo")!.disabled = undoDepth(view.state) === 0;
    this.buttons.get("redo")!.disabled = redoDepth(view.state) === 0;

    if (document.activeElement !== this.fontInput) this.fontInput.value = run.font ?? "";
    if (document.activeElement !== this.sizeInput) this.sizeInput.value = run.size !== undefined ? String(run.size) : "";
    this.syncStyles(paragraph?.attrs.style ?? ctx.styles.defaultParagraph);
  }

  private syncStyles(current: string | null) {
    const { styles } = this.host.ctx;
    const quick = [...styles.byId.values()]
      .filter((s) => s.type === "paragraph" && (s.quick || s.id === current) && (!s.hidden || s.id === current))
      .sort((a, b) => a.priority - b.priority || a.name.localeCompare(b.name));
    const key = quick.map((s) => s.id).join("|");
    if (this.styleSelect.dataset.key !== key) {
      this.styleSelect.replaceChildren(...quick.map((s) => h("option", { value: s.id }, displayName(s.name))));
      this.styleSelect.dataset.key = key;
    }
    if (current) this.styleSelect.value = current;
  }
}

/** Word stores built-in style names in lower case ("heading 1"); show them as Word does. */
function displayName(name: string): string {
  return name.charAt(0).toUpperCase() + name.slice(1);
}
