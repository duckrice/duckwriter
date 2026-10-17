// The name shown with your comments and tracked changes, asked for once and
// remembered on this computer.

const AUTHOR_KEY = "duckwriter.author";

export function authorName(): string {
  try {
    return localStorage.getItem(AUTHOR_KEY) ?? "";
  } catch {
    return "";
  }
}

export function setAuthorName(name: string) {
  try {
    localStorage.setItem(AUTHOR_KEY, name);
  } catch {
    // Asked again next time.
  }
}
