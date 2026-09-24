# Email-Variations-Generator
Generate variation of duplicate valid Email for your email address

**Live demo:** https://fea-email-variations-generator.vercel.app/

## How it works

Gmail ignores dots (`.`) in the username part of an address, so
`johndoe@gmail.com`, `john.doe@gmail.com` and `j.o.h.n.d.o.e@gmail.com` all
deliver to the same inbox. This tool lists every way of placing dots in your
username, which is handy for sign-up testing, filtering, or spotting where an
address was shared.

- Existing dots are ignored, so `john.doe@gmail.com` and `johndoe@gmail.com`
  give the same 64 variations (never `john..doe`).
- A `+tag` suffix (e.g. `john+news@gmail.com`) is kept unchanged; only the
  part before `+` is varied.
- A username of *n* characters has 2^(n-1) variations. The output is capped at
  8,192 (every variation for usernames up to 14 characters) to keep the page
  responsive; the page tells you when results are truncated.
- Results can be exported to an `.xlsx` file (via [SheetJS](https://sheetjs.com/)).

> Dot-insensitivity is a Gmail / Google Workspace behavior. Most other
> providers treat dotted addresses as different mailboxes.

## Run locally

It's a static page with no build step: open `index.html` in a browser, or
serve the folder (e.g. `npx serve .`).

## License

[MIT](LICENSE)
