import { describe, expect, it } from "vitest";
import { cssFontFamily } from "../src/editor/schema";

describe("Linux", () => {
  it("falls back to the free fonts with Office's letter widths", () => {
    expect(cssFontFamily("Calibri")).toBe('"Calibri", Carlito, sans-serif');
    expect(cssFontFamily("Times New Roman")).toBe('"Times New Roman", "Liberation Serif", Tinos, serif');
    expect(cssFontFamily("Aptos")).toBe('"Aptos", sans-serif');
  });
});
