/*
 * A stub standing in for `image-size`, which nothing imports.
 *
 * IT THROWS RATHER THAN RETURNING SOMETHING PLAUSIBLE. The whole basis for this
 * override is that no call site exists — verified across the entire dependency
 * tree, not just our own code. If that ever stops being true, the right outcome
 * is a loud failure naming this file, not a silently wrong image dimension
 * flowing into a generated deck.
 *
 * See README.md beside this file for why the real package is not installed.
 */
const WHY =
  "image-size is stubbed in this project: every published version is covered by " +
  "GHSA-w3rx-r6r6-pgpr and GHSA-5p2g-fcmc-qvqq, and nothing imported it. " +
  "Something now does. See vendor/image-size-stub/README.md — the fix is to " +
  "install the real package once a patched release exists, not to soften this stub.";

function imageSize() {
  throw new Error(WHY);
}

module.exports = imageSize;
module.exports.imageSize = imageSize;
module.exports.default = imageSize;
