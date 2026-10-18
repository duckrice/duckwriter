// Odd files: the samples with elements left out and attributes garbled at
// random (always the same randomness). DuckWriter must open each or say
// plainly it can't, never fail inside; and what it saves must be sound.

import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { checkPackage } from "../src/docx/check";
import { readDocx } from "../src/docx/read";
import { writeDocx } from "../src/docx/write";
import { toHtml, toMarkdown, toText } from "../src/convert/export";
import { toLatex } from "../src/convert/latex";
import { createState } from "../src/editor/editor";
import { listLabels } from "../src/editor/lists";

const samples = readdirSync(join(__dirname, "..", "samples")).filter((n) => n.endsWith(".docx") && n !== "not-a-docx.docx");

/** A small random number generator, the same every run. */
function random(seed: number) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 2 ** 32;
  };
}

const GARBLE = ["", "-1", "999999999", "abc", "auto", "0", "true", "%", "1.5"];

/** Removes some elements and changes some attribute values in a part. */
function mutate(xml: string, rnd: () => number): string {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const all = [...doc.getElementsByTagName("*")].slice(1);
  for (const el of all) {
    const r = rnd();
    if (r < 0.06 && el.parentNode) el.parentNode.removeChild(el);
    else if (r < 0.25 && el.attributes.length) {
      const a = el.attributes[Math.floor(rnd() * el.attributes.length)];
      if (!a.name.startsWith("xmlns")) a.value = GARBLE[Math.floor(rnd() * GARBLE.length)];
    }
  }
  return new XMLSerializer().serializeToString(doc);
}

describe("odd files", () => {
  for (const name of samples) {
    it(`${name}, garbled 30 ways, opens or says why not, and saves soundly`, async () => {
      const bytes = new Uint8Array(readFileSync(join(__dirname, "..", "samples", name)));
      const rnd = random(name.length * 7919);
      for (let round = 0; round < 30; round++) {
        const zip = await JSZip.loadAsync(bytes);
        for (const part of Object.keys(zip.files).filter((p) => /^word\/[^/]+\.xml$/.test(p))) {
          zip.file(part, mutate(await zip.file(part)!.async("string"), rnd));
        }
        const odd = await zip.generateAsync({ type: "uint8array" });
        let opened;
        try {
          opened = await readDocx(odd);
        } catch (error) {
          // Refusing is fine, with a message meant for people.
          expect(error).toBeInstanceOf(Error);
          expect((error as Error).message).not.toMatch(/undefined|null|Cannot read|is not a function/);
          continue;
        }
        // What opening it in the app does with it: the editor, list numbers, the copies.
        createState(opened.doc, () => opened.context);
        listLabels(opened.context, opened.doc);
        toText(opened.doc, opened.context);
        await toMarkdown(opened.doc, opened.context);
        await toHtml(opened.doc, opened.context, "odd");
        await toLatex(opened.doc, opened.context, "odd");
        const saved = await writeDocx(opened.doc, opened.context);
        // What we add must be sound; damage that came with the file may stay.
        const before = new Set(await checkPackage(odd));
        const after = (await checkPackage(saved)).filter((p) => !before.has(p));
        expect(after, `round ${round}`).toEqual([]);
      }
    });
  }
});
