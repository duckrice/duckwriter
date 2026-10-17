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
import { choosePicture, setSnap, snapOn } from "./pictures";
import { TABLE_COMMANDS, insertTable, isInTable, setTableLines, tableGrid } from "./tables";
import { captionDialog, headingsNumbered, insertReference, insertToc, referenceTargets, toggleHeadingNumbers } from "./references";
import { icon, type IconName } from "./icons";

export interface ToolbarHost {
  readonly view: EditorView;
  readonly ctx: DocContext;
  /** Switches the line grid on or off. */
  setLineGrid(on: boolean): void;
  newDocument(): void;
  open(): void;
  save(): void;
  editHeader(): void;
  insertFootnote(): void;
  showShortcuts(): void;
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
  private menuArrows = new Map<string, HTMLButtonElement>();
  /** The table buttons, shown while the cursor is in a table. */
  private tableGroup: HTMLElement;
  private styleByKeyboard = false;

  constructor(private host: ToolbarHost) {
    const fonts = h("datalist", { id: "dw-fonts" }, ...FONTS.map((f) => h("option", { value: f })));
    const sizes = h("datalist", { id: "dw-sizes" }, ...SIZES.map((s) => h("option", { value: String(s) })));

    this.styleSelect.addEventListener("change", () => {
      // Chosen with the keyboard (Ctrl+Shift+S), the box keeps the focus until Enter.
      this.run(setStyle(this.host.ctx, this.styleSelect.value || null), !this.styleByKeyboard);
    });
    this.styleSelect.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === "Escape") {
        event.preventDefault();
        this.styleByKeyboard = false;
        this.host.view.focus();
      }
    });
    this.styleSelect.addEventListener("blur", () => (this.styleByKeyboard = false));
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
    const textButton = (label: string, title: string, command: () => void) => {
      const b = h("button", { type: "button", class: "tb tb-text", title, "aria-label": title }, label);
      b.addEventListener("mousedown", (event) => event.preventDefault());
      b.addEventListener("click", command);
      return b;
    };
    const t = TABLE_COMMANDS;
    this.tableGroup = h(
      "div",
      { class: "tb-group table-tools", hidden: true },
      h("span", { class: "tb-caption" }, "Table"),
      textButton("+ Row above", "Insert a row above (Ctrl+Alt+Shift+Up)", () => this.run(t.rowAbove)),
      textButton("+ Row below", "Insert a row below (Ctrl+Alt+Shift+Down)", () => this.run(t.rowBelow)),
      textButton("+ Column left", "Insert a column to the left (Ctrl+Alt+Shift+Left)", () => this.run(t.columnLeft)),
      textButton("+ Column right", "Insert a column to the right (Ctrl+Alt+Shift+Right)", () => this.run(t.columnRight)),
      textButton("− Row", "Delete the row (Ctrl+Alt+Shift+Backspace)", () => this.run(t.deleteRow)),
      textButton("− Column", "Delete the column (Ctrl+Alt+Shift+Delete)", () => this.run(t.deleteColumn)),
      textButton("Merge", "Merge the selected cells (Ctrl+Alt+Shift+M)", () => this.run(t.merge)),
      textButton("Split", "Split the cell (Ctrl+Alt+Shift+N)", () => this.run(t.split)),
      textButton("Header row", "First row as a header, repeated on each page in Word (Ctrl+Alt+Shift+R)", () => this.run(t.headerRow)),
      textButton("Lines", "Lines around the cells, on or off (Ctrl+Alt+Shift+L)", () => this.toggleTableLines()),
      this.splitButton("shading", "Cell shading", null, () => this.shadingMenu(), () => this.shadingMenu()),
      textButton("Delete table", "Delete the whole table", () => this.run(t.deleteTable)),
    );
    this.el = h(
      "div",
      { class: "toolbar", role: "toolbar" },
      group(
        this.button("newDoc", "New document (Ctrl+N)", () => this.host.newDocument()),
        this.button("open", "Open (Ctrl+O)", () => this.host.open()),
        this.button("save", "Save (Ctrl+S)", () => this.host.save()),
        this.button("keyboard", "Keyboard shortcuts (F1)", () => this.host.showShortcuts()),
      ),
      group(this.button("undo", "Undo (Ctrl+Z)", () => this.run(undo)), this.button("redo", "Redo (Ctrl+Y)", () => this.run(redo))),
      group(this.styleSelect),
      group(
        this.fontInput,
        this.sizeInput,
        this.button("zoomIn", "Bigger text (Ctrl+> or Ctrl+])", () => this.run(stepSize(ctx(), 1)), "sizeUp"),
        this.button("zoomOut", "Smaller text (Ctrl+< or Ctrl+[)", () => this.run(stepSize(ctx(), -1)), "sizeDown"),
      ),
      group(
        this.button("bold", "Bold (Ctrl+B)", () => this.run(toggleFormat(ctx(), "bold"))),
        this.button("italic", "Italic (Ctrl+I)", () => this.run(toggleFormat(ctx(), "italic"))),
        this.button("underline", "Underline (Ctrl+U)", () => this.run(toggleFormat(ctx(), "underline"))),
        this.button("strike", "Strikethrough (Ctrl+Shift+X)", () => this.run(toggleFormat(ctx(), "strike"))),
        this.button("superscript", "Superscript (Ctrl++)", () => this.run(toggleVertAlign("superscript"))),
        this.button("subscript", "Subscript (Ctrl+=)", () => this.run(toggleVertAlign("subscript"))),
      ),
      group(
        this.splitButton("textColor", "Text colour (Ctrl+Shift+O)", this.textColorBar, () => this.run(setMark(schema.marks.color, { color: this.lastTextColor })), () => this.colorMenu()),
        this.splitButton("highlight", "Highlight (Ctrl+Alt+H)", this.highlightBar, () => this.run(setMark(schema.marks.highlight, { color: this.lastHighlight })), () => this.highlightMenu()),
        this.button("clearFormat", "Clear formatting (Ctrl+Space)", () => this.run(clearFormatting(ctx()))),
      ),
      group(
        this.button("alignLeft", "Align left (Ctrl+L)", () => this.run(setAlign("left"))),
        this.button("alignCenter", "Centre (Ctrl+E)", () => this.run(setAlign("center"))),
        this.button("alignRight", "Align right (Ctrl+R)", () => this.run(setAlign("right"))),
        this.button("alignJustify", "Justify (Ctrl+J)", () => this.run(setAlign("justify"))),
      ),
      group(
        this.button("bullets", "Bullets (Ctrl+Shift+L)", () => this.run(toggleList(ctx(), "bullet"))),
        this.button("numbers", "Numbering (Ctrl+Shift+7)", () => this.run(toggleList(ctx(), "number"))),
        this.button("indentLess", "Decrease indent (Ctrl+Shift+M)", () => this.indent(-1)),
        this.button("indentMore", "Increase indent (Ctrl+M)", () => this.indent(1)),
        this.splitButton("lineSpacing", "Line spacing (Ctrl+1, Ctrl+5, Ctrl+2)", null, () => this.spacingMenu(), () => this.spacingMenu()),
      ),
      group(
        this.button("picture", "Picture (Ctrl+Alt+P)", () => void choosePicture(this.host.view, this.host.ctx)),
        this.splitButton("table", "Table (Ctrl+Alt+Shift+T)", null, () => this.tableMenu(), () => this.tableMenu()),
        this.button("equation", "Equation (Alt+=)", () => insertEquation(this.host.view)),
        this.splitButton("symbol", "Symbol (Ctrl+Shift+Y)", null, () => this.symbolMenu(), () => this.symbolMenu()),
        this.button("pageBreak", "Page break (Ctrl+Enter)", () => this.run(insertPageBreak())),
        this.button("header", "Header and footer (Ctrl+Alt+O)", () => this.host.editHeader()),
        this.button("footnote", "Footnote (Ctrl+Alt+F)", () => this.host.insertFootnote()),
        this.splitButton("references", "References (Ctrl+Alt+R)", null, () => this.referencesMenu(), () => this.referencesMenu()),
        this.splitButton("layout", "Layout", null, () => this.layoutMenu(), () => this.layoutMenu()),
      ),
      this.tableGroup,
      fonts,
      sizes,
    );
    this.textColorBar.style.background = this.lastTextColor;
    this.highlightBar.style.background = HIGHLIGHTS[this.lastHighlight];
    // Escape closes an open menu, wherever the cursor is.
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && this.openMenu) {
        this.closeMenu();
        this.host.view.focus();
      }
    });
    document.addEventListener("pointerdown", (event) => {
      if (this.openMenu && !this.openMenu.contains(event.target as Node) && !(event.target as HTMLElement).closest(".split-arrow")) this.closeMenu();
    });
  }

  /** Runs a command on the editor and gives it the focus back. */
  run(command: Command, focus = true) {
    const { view } = this.host;
    command(view.state, view.dispatch, view);
    if (focus) view.focus();
  }

  /** Puts the cursor in the font, size or style box, for choosing from the keyboard. */
  focusBox(box: "font" | "size" | "style") {
    if (box === "style") {
      this.styleByKeyboard = true;
      this.styleSelect.focus();
      return;
    }
    const input = box === "font" ? this.fontInput : this.sizeInput;
    input.focus();
    input.select();
  }

  /** Opens one of the toolbar's menus, ready to use with the arrow keys. */
  openMenuByKeyboard(name: string) {
    this.menuArrows.get(name)?.click();
    const first = this.openMenu?.querySelector<HTMLElement>(".menu-item, .symbol-button, .swatch, .table-grid-cell");
    first?.focus();
  }

  applyTextColor() {
    this.run(setMark(schema.marks.color, { color: this.lastTextColor }));
  }

  applyHighlight() {
    this.run(setMark(schema.marks.highlight, { color: this.lastHighlight }));
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
    this.menuArrows.set(name, arrow);
    const wrap = h("div", { class: "split" }, main, arrow);
    arrow.addEventListener("click", () => {
      if (this.openMenu?.dataset.owner === name) return this.closeMenu();
      this.closeMenu();
      const menu = onMenu();
      menu.dataset.owner = name;
      menu.addEventListener("mousedown", (event) => {
        if (!(event.target instanceof HTMLInputElement)) event.preventDefault();
      });
      menu.addEventListener("keydown", (event) => this.menuKeys(event, menu));
      wrap.append(menu);
      this.openMenu = menu;
    });
    return wrap;
  }

  /** Arrow keys move through a menu's choices; Escape closes it. */
  private menuKeys(event: KeyboardEvent, menu: HTMLElement) {
    const items = [...menu.querySelectorAll<HTMLElement>(".menu-item, .symbol-button, .swatch, .table-grid-cell")];
    const at = items.indexOf(document.activeElement as HTMLElement);
    // How many choices make a row: those level with the one that has the focus.
    const current = items[at] ?? items[0];
    const across = current ? Math.max(1, items.filter((i) => i.parentElement === current.parentElement && i.offsetTop === current.offsetTop).length) : 1;
    const go = (to: number) => {
      event.preventDefault();
      items[Math.max(0, Math.min(items.length - 1, to))]?.focus();
    };
    if (event.key === "ArrowDown") go(at + (items[at]?.matches(".menu-item") ? 1 : across));
    else if (event.key === "ArrowUp") go(at - (items[at]?.matches(".menu-item") ? 1 : across));
    else if (event.key === "ArrowRight") go(at + 1);
    else if (event.key === "ArrowLeft") go(at - 1);
    else if ((event.key === "Enter" || event.key === " ") && at >= 0) {
      event.preventDefault();
      items[at].click();
    } else if (event.key === "Escape") {
      event.preventDefault();
      this.closeMenu();
      this.host.view.focus();
    }
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

  private tableMenu(): HTMLElement {
    return tableGrid((rows, cols) => {
      this.closeMenu();
      insertTable(this.host.view, this.host.ctx, rows, cols);
    });
  }

  private shadingMenu(): HTMLElement {
    const none = h("button", { type: "button", class: "menu-item" }, "No shading");
    none.addEventListener("click", () => {
      this.run(TABLE_COMMANDS.shade(null));
      this.closeMenu();
    });
    const shades = ["#f2f2f2", "#d9d9d9", "#fff2cc", "#e2efda", "#ddebf7", "#fce4d6", "#ededed", "#d9e1f2", "#ffd966", "#a9d08e", "#9bc2e6", "#f4b084"];
    return h("div", { class: "menu" }, this.swatches(shades, (c) => c, (c) => this.run(TABLE_COMMANDS.shade(c)), (c) => c), none);
  }

  /** Lines around the table's cells, on or off: the opposite of what shows now. */
  toggleTableLines() {
    const { view } = this.host;
    const $from = view.state.selection.$from;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type.name !== "table") continue;
      const table = view.nodeDOM($from.before(d)) as HTMLElement | null;
      const cell = table?.querySelector("td, th");
      const shown = !!cell && getComputedStyle(cell).borderTopStyle === "solid" && getComputedStyle(cell).borderTopColor !== "rgba(0, 0, 0, 0)";
      this.run(setTableLines(!shown));
      return;
    }
  }

  /** Snapping pictures to the grid. */
  private layoutMenu(): HTMLElement {
    const snap = h("button", { type: "button", class: `menu-item${snapOn() ? " checked" : ""}` }, `${snapOn() ? "✓ " : ""}Snap pictures to grid`);
    snap.addEventListener("click", () => {
      setSnap(!snapOn());
      this.closeMenu();
    });
    const gridOn = this.host.ctx.page.lineGrid !== null;
    const lines = h("button", { type: "button", class: `menu-item${gridOn ? " checked" : ""}` }, `${gridOn ? "✓ " : ""}Line grid`);
    lines.addEventListener("click", () => {
      this.host.setLineGrid(!gridOn);
      this.closeMenu();
    });
    return h(
      "div",
      { class: "menu layout-menu" },
      snap,
      h("div", { class: "menu-hint" }, "Pictures line up on a half-centimetre grid and with the margins when moved or resized. Hold Alt to place freely."),
      h("div", { class: "menu-separator" }),
      lines,
      h("div", { class: "menu-hint" }, "Every line of text sits on the same evenly spaced lines, on every page, as in Word's document grid."),
    );
  }

  /** Captions, cross-references, numbered headings and the table of contents. */
  private referencesMenu(): HTMLElement {
    const { ctx } = this.host;
    const body = this.host.view;
    const item = (label: string, onClick: () => void, checked = false) => {
      const b = h("button", { type: "button", class: `menu-item${checked ? " checked" : ""}` }, label);
      b.addEventListener("click", () => {
        this.closeMenu();
        onClick();
      });
      return b;
    };
    const targets = referenceTargets(ctx, body.state.doc);
    return h(
      "div",
      { class: "menu references-menu" },
      item("Insert caption…", () => captionDialog(body, ctx)),
      item("Table of contents", () => insertToc(body, ctx)),
      item(headingsNumbered(ctx) ? "✓ Numbered headings" : "Numbered headings (1, 1.1, 1.1.1)", () => toggleHeadingNumbers(body, ctx)),
      item("Footnote (Ctrl+Alt+F)", () => this.host.insertFootnote()),
      h("div", { class: "menu-separator" }),
      h("div", { class: "menu-heading" }, targets.length ? "Refer to…" : "Refer to… (add a caption or heading first)"),
      ...targets.map((t) => item(t.label.length > 60 ? `${t.label.slice(0, 57)}…` : t.label, () => insertReference(body, t))),
    );
  }

  private symbolMenu(): HTMLElement {
    const groups: [string, string][] = [
      ["Greek", "α β γ δ ε ζ η θ ι κ λ μ ν ξ π ρ σ τ υ φ χ ψ ω Γ Δ Θ Λ Ξ Π Σ Φ Ψ Ω"],
      ["Maths", "± ∓ × ÷ ⋅ ≤ ≥ ≠ ≈ ≡ ∼ ∝ ∞ ∂ ∇ √ ∑ ∏ ∫ ∮ ∈ ∉ ⊂ ⊆ ∪ ∩ ∅ ∀ ∃ ¬ ∧ ∨ ∴ ′ ″"],
      ["Arrows", "→ ← ↔ ⇒ ⇐ ⇔ ↑ ↓ ⇌ ↦"],
      ["Units and more", "° ‰ µ Å Ω ℃ ℉ ½ ⅓ ¼ ¾ ² ³ ¹ ⁰ ⁻ ₀ ₁ ₂ ₃ € £ © ® ™ § ¶ † ‡ • … – —"],
    ];
    const insert = (text: string, fromKeyboard: boolean) => {
      const { view } = this.host;
      view.dispatch(view.state.tr.insertText(text).scrollIntoView());
      // Chosen with the keyboard, the menu keeps the focus for the next one (Esc to finish).
      if (!fromKeyboard) view.focus();
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
            b.addEventListener("click", (event) => insert(ch, event.detail === 0));
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
    this.tableGroup.hidden = !isInTable(view.state);
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
