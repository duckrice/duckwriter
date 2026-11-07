import { describe, expect, it } from "vitest";
import { MathEdit, fromLatex, rowAt } from "../src/math/edit";
import { parseMath } from "../src/math/ast";

const typed = (text: string) => {
  const e = new MathEdit();
  e.type(text);
  return e;
};

describe("typing an equation straight into the text", () => {
  it("builds up subscripts on space, with every letter after _ in it", () => {
    expect(typed("I_inn ").latex()).toBe("I_{inn}");
    expect(typed("x^2+y^2=r^2 ").latex()).toBe("x^2+y^2=r^2");
    expect(typed("x_i^2 ").latex()).toBe("x_i^2");
  });

  it("x_ and a space makes an empty box with the cursor in it", () => {
    const e = typed("i_ ");
    expect(e.caret.path).toEqual([0, 1]);
    expect(e.row).toEqual([]);
    e.type("max");
    e.right();
    e.type("+1");
    expect(e.latex()).toBe("i_{max}+1");
  });

  it("fractions take whole operands, and brackets around one go away", () => {
    expect(typed("a/b ").latex()).toBe("\\frac{a}{b}");
    expect(typed("ab/c ").latex()).toBe("\\frac{ab}{c}");
    expect(typed("(a+b)/c ").latex()).toBe("\\frac{a+b}{c}");
    expect(typed("1+x^2/3 ").latex()).toBe("1+\\frac{x^2}{3}");
    expect(typed("x^(n+1) ").latex()).toBe("x^{n+1}");
    // A second ^ goes on top of the first; sin and lim are not part of a fraction after them.
    expect(typed("e^-x^2 ").latex()).toBe("e^{-x^2}");
    expect(typed("sin x/2 ").latex()).toBe("\\sin\\frac{x}{2}");
  });

  it("signs belong to the script after ^ and _", () => {
    expect(typed("x^-1 ").latex()).toBe("x^{-1}");
    expect(typed("Na^+ +Cl^- ").latex()).toBe("Na^++Cl^-");
  });

  it("brackets grow around what's inside", () => {
    const e = typed("(a/b )");
    expect(e.root[0]).toMatchObject({ kind: "delim", open: "(", close: ")" });
    expect(e.latex()).toBe("\\left( \\frac{a}{b} \\right)");
  });

  it("commands and function names", () => {
    expect(typed("\\alpha +\\beta ").latex()).toBe("\\alpha+\\beta");
    expect(typed("\\alpha_1 ").latex()).toBe("\\alpha_1");
    expect(typed("sin x").latex()).toBe("\\sin x");
    const f = typed("\\frac ");
    expect(f.caret.path).toEqual([0, 0]);
    f.type("1");
    f.vertical(1);
    f.type("2");
    expect(f.latex()).toBe("\\frac{1}{2}");
    const s = typed("\\sqrt ");
    s.type("x");
    expect(s.latex()).toBe("\\sqrt{x}");
    const sum = typed("\\sum ");
    sum.type("i=1");
    sum.right();
    sum.type("n");
    sum.right();
    sum.type("i");
    expect(parseMath(sum.latex())).toEqual(parseMath("\\sum_{i=1}^{n} i"));
    const lim = typed("lim ");
    lim.type("x\\to 0");
    lim.right();
    expect(lim.latex()).toBe("\\lim_{x\\to0}");
  });

  it("quoted words stay ordinary text", () => {
    expect(typed('"rate"=k').latex()).toBe("\\text{rate}=k");
  });

  it("the arrows go into and out of every box", () => {
    const e = typed("a/b ");
    e.home();
    expect(e.right()).toBe(true); // into the top
    expect(e.caret.path).toEqual([0, 0]);
    e.right();
    e.right(); // to the bottom
    expect(e.caret.path).toEqual([0, 1]);
    e.right();
    e.right(); // out
    expect(e.caret).toEqual({ path: [], index: 1 });
    expect(e.right()).toBe(false); // leaves the equation
    e.left();
    expect(e.caret.path).toEqual([0, 1]);
  });

  it("backspace marks a box first, then removes it; at a box's start it comes apart", () => {
    const e = typed("a/b ");
    e.backspace();
    expect(e.marked).toBe(true);
    e.backspace();
    expect(e.isEmpty()).toBe(true);
    const f = typed("x_ ");
    f.type("i");
    f.left();
    f.backspace();
    expect(f.latex()).toBe("xi");
  });

  it("a space with nothing to build stays a space, and keeps things apart", () => {
    expect(typed("a b").latex()).toBe("a\\;b");
    expect(typed("a b/c ").latex()).toBe("a\\;\\frac{b}{c}");
    expect(new MathEdit(fromLatex("a\\;b")).latex()).toBe("a\\;b");
  });

  it("Enter in a matrix adds a row", () => {
    const e = typed("\\pmatrix ");
    e.type("a");
    e.nextCell();
    e.type("b");
    e.addMatrixRow();
    e.type("c");
    expect(e.latex()).toBe("\\begin{pmatrix} a & b \\\\ c &  \\\\  &  \\end{pmatrix}");
  });

  it("leaving builds up whatever is still typed out", () => {
    const e = typed("E=mc^2");
    e.finish();
    expect(e.latex()).toBe("E=mc^2");
    expect(rowAt(e.root, [])).toHaveLength(4);
  });

  it("half-typed shorthand is saved so it reads back the same", () => {
    for (const text of ["x_", "\\alph", "a^", "50%", "a&b"]) {
      const latex = typed(text).latex();
      expect(() => parseMath(latex)).not.toThrow();
      expect(new MathEdit(fromLatex(latex)).latex()).toBe(latex);
    }
  });

  it("equations from files open for editing unchanged", () => {
    for (const latex of [
      "x^2+y^2=r^2",
      "\\frac{a+b}{c}",
      "\\sqrt{x}+\\sqrt[3]{y}",
      "\\sum_{i=1}^{n} x_i",
      "\\lim_{x\\to 0}\\frac{\\sin x}{x}=1",
      "\\vec{F}=m\\vec{a}",
      "\\text{rate}=k[A]^2",
      "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}",
    ]) {
      const again = new MathEdit(fromLatex(latex)).latex();
      expect(parseMath(again)).toEqual(parseMath(latex));
    }
  });
});
