// AutoCorrect: common typing slips corrected as a word is finished (teh →
// the), a few symbols ((c) → ©), and corrections of one's own, added from the
// spelling menu ("Always correct to …"), remembered on this computer.

import type { EditorState, Transaction } from "prosemirror-state";

const KEY = "duckwriter.autocorrect";

/** Slips common enough to correct without asking. */
const BUILT_IN: Record<string, string> = {
  teh: "the", adn: "and", hte: "the", taht: "that", wiht: "with", waht: "what", whcih: "which", wich: "which", thier: "their",
  recieve: "receive", recieved: "received", seperate: "separate", seperately: "separately", occured: "occurred", occurence: "occurrence",
  occuring: "occurring", definately: "definitely", untill: "until", accomodate: "accommodate", acheive: "achieve", acheived: "achieved",
  beleive: "believe", enviroment: "environment", existance: "existence", goverment: "government", independant: "independent",
  refered: "referred", succesful: "successful", successfull: "successful", tommorow: "tomorrow", truely: "truly", wierd: "weird",
  arguement: "argument", calender: "calendar", concious: "conscious", embarass: "embarrass", experiance: "experience",
  foriegn: "foreign", freind: "friend", gaurd: "guard", happend: "happened", immediatly: "immediately", knowlege: "knowledge",
  neccessary: "necessary", noticable: "noticeable", posession: "possession", publically: "publicly", recomend: "recommend",
  relevent: "relevant", resistence: "resistance", sucess: "success", supercede: "supersede", tendancy: "tendency",
  threshhold: "threshold", wether: "whether", writting: "writing",
};

/** Symbols typed with ordinary keys. */
const SYMBOLS: [string, string][] = [
  ["(c)", "©"],
  ["(r)", "®"],
  ["(tm)", "™"],
  ["...", "…"],
];

function own(): Record<string, string> {
  try {
    const value = JSON.parse(localStorage.getItem(KEY) ?? "{}");
    return value && typeof value === "object" ? value : {};
  } catch {
    return {};
  }
}

/** From now on, `word` is corrected to `correction` as it is typed. */
export function addCorrection(word: string, correction: string) {
  const list = own();
  list[word.toLowerCase()] = correction;
  try {
    localStorage.setItem(KEY, JSON.stringify(list));
  } catch {
    // For now only.
  }
}

/** The correction for a word, keeping its capitals ("Teh" → "The", "TEH" → "THE"). */
export function correctionFor(word: string): string | null {
  const lower = word.toLowerCase();
  const found = own()[lower] ?? BUILT_IN[lower];
  if (!found || found === word) return null;
  if (word.length > 1 && word === word.toUpperCase()) return found.toUpperCase();
  if (word[0] !== word[0].toLowerCase()) return found[0].toUpperCase() + found.slice(1);
  return found;
}

/** Corrects the word or symbol just finished (by `typed`: a space or punctuation). */
export function autoCorrect(state: EditorState, typed: string): Transaction | null {
  const { $from, empty } = state.selection;
  if (!empty || !$from.parent.isTextblock || !/^[\s.,;:!?)\]"'’]$/.test(typed)) return null;
  const before = $from.parent.textBetween(0, $from.parentOffset, undefined, "￼");
  for (const [from, to] of SYMBOLS) {
    if (before.endsWith(from)) return state.tr.insertText(to, $from.pos - from.length, $from.pos);
  }
  // "word -- word": an en dash.
  if (typed === " " && / --$/.test(before)) return state.tr.insertText("–", $from.pos - 2, $from.pos);
  const word = /[\p{L}'’]+$/u.exec(before)?.[0];
  if (!word) return null;
  const correction = correctionFor(word);
  return correction ? state.tr.insertText(correction, $from.pos - word.length, $from.pos) : null;
}
