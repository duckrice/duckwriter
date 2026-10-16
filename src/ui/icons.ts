// 24x24 outline icons, drawn with the current text colour.

const PATHS = {
  newDoc: '<path d="M7 3h7l5 5v11a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M14 3v5h5M12 11v6M9 14h6"/>',
  open: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
  save: '<path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M8 3v5h7V3M7 21v-7h10v7"/>',
  undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/>',
  redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H10a6 6 0 0 0 0 12h3"/>',
  bold: '<path d="M7 5h6a3.500 3.500 0 0 1 0 7H7zM7 12h7a3.500 3.500 0 0 1 0 7H7z" stroke-width="2.400"/>',
  italic: '<path d="M10 5h8M6 19h8M14 5l-4 14"/>',
  underline: '<path d="M7 4v7a5 5 0 0 0 10 0V4M5 21h14"/>',
  strike: '<path d="M4 12h16M16.500 7.500C16 5.500 14 4.500 12 4.500c-2.500 0-4.500 1.300-4.500 3.300 0 1.500 1 2.500 3 3.200M8 16.500c.5 2 2.300 3 4.300 3 2.700 0 4.500-1.400 4.500-3.400 0-.8-.2-1.500-.7-2.100"/>',
  superscript: '<path d="m4 8 8 10M12 8l-8 10M16 9c0-1.500 1-2.500 2.300-2.500S20.500 7.300 20.500 8.500c0 2-4.500 3-4.500 4h4.500"/>',
  subscript: '<path d="m4 5 8 10M12 5 4 15M16 16c0-1.500 1-2.500 2.300-2.500s2.200.8 2.200 2c0 2-4.500 3-4.500 4h4.500"/>',
  textColor: '<path d="m6 16 6-13 6 13M8.300 11h7.400"/>',
  highlight: '<path d="m9 11-6 6v3h9l3-3"/><path d="m22 12-4.600 4.600a2 2 0 0 1-2.800 0l-5.200-5.200a2 2 0 0 1 0-2.800L14 4"/>',
  alignLeft: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
  alignCenter: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
  alignRight: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
  alignJustify: '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
  bullets: '<circle cx="5" cy="7" r="1.200" fill="currentColor"/><circle cx="5" cy="12" r="1.200" fill="currentColor"/><circle cx="5" cy="17" r="1.200" fill="currentColor"/><path d="M9 7h11M9 12h11M9 17h11"/>',
  numbers: '<path d="M10 7h10M10 12h10M10 17h10M4 5h1.500v4M4 9h3M4 14.500c0-.8.700-1.500 1.500-1.500S7 13.600 7 14.300c0 1.200-3 1.500-3 3.200h3" stroke-width="1.500"/>',
  indentMore: '<path d="M4 5h16M11 10h9M11 14h9M4 19h16M4 9.500 7 12l-3 2.500"/>',
  indentLess: '<path d="M4 5h16M11 10h9M11 14h9M4 19h16M7 9.500 4 12l3 2.500"/>',
  lineSpacing: '<path d="M11 6h9M11 12h9M11 18h9M5 4v16M3 6l2-2 2 2M3 18l2 2 2-2"/>',
  clearFormat: '<path d="M6 5h12M12 5 9 19M4 20 20 4"/>',
  caret: '<path d="m7 10 5 5 5-5"/>',
  pageBreak: '<path d="M6 3v5h12V3M6 21v-5h12v5M3 12h3M9 12h2M13 12h2M18 12h3"/>',
  zoomIn: '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4M8 11h6M11 8v6"/>',
  zoomOut: '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4M8 11h6"/>',
} as const;

export type IconName = keyof typeof PATHS;

export function icon(name: IconName, size = 18): string {
  return `<svg class="icon" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.800" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${PATHS[name]}</svg>`;
}
