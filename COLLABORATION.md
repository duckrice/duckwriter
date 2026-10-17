# Writing together: three ways, to choose from

DuckWriter can now let several people write in the same document at the same
time: each sees the others' typing as it happens, with a coloured cursor and
their name. All three ways below are built and tested; they differ only in
**how the messages travel**. You choose one (or keep several) as the way
DuckWriter should offer it.

How to try: **Write together** (the people button in the toolbar, or
Alt+Shift+W) → choose a way → **Start sharing** → send the invitation
(`DW1.…`) to the other person → they choose **Write together → Join** and
paste it.

## What is the same in all three

- **The person sharing keeps the document.** Their DuckWriter puts everyone's
  changes in order (ProseMirror's collaboration protocol), so two people typing
  in the same sentence never lose anything, and everyone ends up with exactly
  the same text. Only they save it to the file; the others can save a copy.
- **Shared live:** the text, change by change, with formatting, tables,
  equations, pictures, citations, tracked changes (each under their own name)
  and comments. Footnotes, headers and footers are shared too, as whole notes
  (if two people edit the *same* footnote at once, the last one wins).
- **Encrypted end to end** (AES-256-GCM). The key is in the invitation and never
  travels with the messages, so the network, the relay server or the cloud
  folder in between can't read the document. A wrong or old invitation can't
  join.
- **No accounts.** You type your name once; it is shown by your cursor.
- The person sharing has to keep DuckWriter open while the others write. When
  they stop, everyone keeps their copy.
- Undo undoes only your own changes.

## The three ways

| | A. On this network | B. Relay server | C. Shared folder |
|---|---|---|---|
| **Where it works** | Same Wi-Fi / office / university network | Anywhere with internet | Anywhere, through OneDrive, Dropbox, Google Drive or a network drive everyone has |
| **What you need** | Nothing | A small server running `server/relay.mjs` (see below) | A folder everyone has synced |
| **Speed** | Instant | Instant (one hop over the internet) | As fast as the folder syncs: about 1–10 s with OneDrive/Dropbox, instant on a network drive |
| **Cost** | Free | Free tier or ~5 USD/month for a tiny server; or a Raspberry Pi at home | Free (uses the cloud storage people already have) |
| **Privacy** | Never leaves the network | Encrypted; the server sees only scrambled text and when people are online | Encrypted log files sit in the shared folder while sharing |
| **Set-up for you** | None | Run the server once (I can set it up when you have an account somewhere) | None |
| **Snags** | Windows asks once whether DuckWriter may use the network (allow private networks). Some guest/eduroam networks block devices from seeing each other. | Someone has to run the server | Slower; leaves a `.duckwriter-live` folder (small, can be deleted afterwards) |

### My recommendation

Offer **A** and **C** straight away: they need no server and cost nothing, and
between them they cover "we're in the same place" and "we're not, but we share
a OneDrive/Dropbox folder". Add **B** if you want real-time writing over the
internet without a shared folder; it needs one small server, which I can
deploy once you pick a host (Fly.io, Render, a cheap VPS, or a Raspberry Pi).

### Not built (ideas for later)

- **D. Always-on cloud documents** (like Google Docs): the server keeps the
  document, so people can write even when the sharer is offline. That means a
  real cloud service with accounts and storage, and the server would have to
  be able to read documents (or a much more involved encrypted design). Weeks of
  work and a running cost; worth it only if you want DuckWriter to become a
  hosted service.
- **E. Peer-to-peer over the internet (WebRTC)**: direct connections between
  computers without a relay, but it still needs a small "introduction" server,
  and fails on some networks; B is simpler and more reliable.

## Where the code is

- `src/collab/session.ts`: the shared core (who is the authority, sending and
  receiving changes, pictures, comments, footnotes, the starting copy).
- `src/collab/hub.ts`: the ways of connecting (`SocketHub` for A and B,
  `FolderHub` for C). `src/collab/crypto.ts`: the encryption.
  `src/collab/presence.ts`: the others' cursors.
- `src-tauri/src/collab.rs`: A's meeting point inside DuckWriter, and C's file
  access (limited to `.duckwriter-live` folders).
- `server/relay.mjs`: B's relay server.
- `src/ui/share.ts`: the Write together dialog.
- Tests: `tests/collab.test.ts` (two people typing at once over a slow
  connection), and the app test runs two real DuckWriters against each other
  over all three.
