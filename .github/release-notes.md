DuckWriter is a free word processor for Windows and Linux that opens and saves Word documents (.docx), with what scientific writing needs: equations, footnotes and endnotes, captions and cross-references, citations and a reference list (APA, Harvard, Chicago, IEEE, Vancouver; BibTeX, RIS, CSL-JSON, DOI lookup, Zotero), tracked changes and comments, and sections with their own page numbers. It also opens older Word documents (.doc), RTF, OpenDocument and PDFs (worked out into an editable document), and saves copies as PDF, LaTeX, Markdown, web pages and plain text.

## Install on Windows

1. Download **DuckWriter-Setup.exe** below. That one file is all you need.
2. Run it. Windows may say "Windows protected your PC", because the installer isn't code-signed. Click **More info**, then **Run anyway**.
3. Open DuckWriter from the Start menu, or right-click a Word document and choose **Open with**.

It needs Windows 10 or 11. Your documents stay on your computer. DuckWriter goes online only to check for updates, and when you ask it to: to look up a DOI, or to write together with someone through a relay server.

## Install on Linux

1. Download **DuckWriter-x86_64.AppImage** below.
2. Make it runnable and start it: `chmod +x DuckWriter-x86_64.AppImage && ./DuckWriter-x86_64.AppImage`. On Arch, AppImages need `sudo pacman -S fuse2` once.
3. It updates itself, like the Windows version. On Debian or Ubuntu you can install **DuckWriter-amd64.deb** instead (it doesn't update itself).

Optional extras:

- Spelling: `sudo pacman -S hunspell-en_us` (and your language's Hunspell dictionary).
- Fonts that match Word's (Calibri, Cambria, Arial, Times New Roman): `sudo pacman -S ttf-carlito ttf-caladea ttf-liberation`.

The "Source code" files below are added by GitHub automatically; you don't need them.
