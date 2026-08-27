# Why `image-size` is stubbed

Two HIGH advisories — [GHSA-w3rx-r6r6-pgpr][icns] (ICNS parser) and
[GHSA-5p2g-fcmc-qvqq][jxl] (JXL/HEIF parsers) — cover **every published version**
of `image-size`. Both give their vulnerable range as `<=2.0.2`, and 2.0.2 is the
latest release. There is nothing to upgrade to, so the `overrides` pin that
cleared `qs` has no target here.

npm's own remedy is worse than the disease. `npm audit fix --force` names
`pptxgenjs@1.1.5` against the **4.0.1** this project runs: a three-major
**downgrade** of the deck writer, to patch a parser that never executes.

## The dependency is not used

`image-size` reaches us as `pptxgenjs -> image-size`. Checked 2026-08-27:

- It appears in `node_modules/pptxgenjs/package.json` and **nowhere in
  pptxgenjs's code** — no `require("image-size")`, no `from "image-size"`, in
  `dist/` or anywhere else in the package.
- Nothing else in the whole `node_modules` tree imports it either.
- This project never calls it, and never calls `addImage`, which is the
  pptxgenjs API the real package exists to serve.

So it is declared-but-unimported metadata: npm installs it, `npm audit` flags it,
and no code path can reach it. That is what makes replacing it safe — there is no
call site to break.

## What the stub does

`overrides` in the root `package.json` points `image-size` at this directory.
The vulnerable code is never installed, so the advisories have nothing to match.

The stub **throws** if anything imports it. It deliberately does not return a
plausible size: the entire justification here is "no call site exists", and if
that stops being true the correct outcome is a loud failure naming this
directory, not a wrong dimension flowing quietly into a generated deck.

## When to remove this

Delete the `overrides` entry and this directory as soon as `image-size` ships a
release outside `<=2.0.2`, or pptxgenjs drops the dependency it does not use.
Then `npm audit` should be clean on its own.

**And remove it immediately if this project ever calls `addImage`** — putting a
bitmap into a generated deck makes the real package live, and the stub would
turn a working feature into a thrown error. That is the intended behaviour, but
the fix is the real package, not a softer stub.

[icns]: https://github.com/advisories/GHSA-w3rx-r6r6-pgpr
[jxl]: https://github.com/advisories/GHSA-5p2g-fcmc-qvqq
