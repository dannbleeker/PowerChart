import js from "@eslint/js";
import tseslint from "typescript-eslint";
import prettier from "eslint-config-prettier";
import globals from "globals";

/**
 * Flat ESLint config. Correctness-focused: the TypeScript recommended set for
 * real bugs, Prettier owns formatting (eslint-config-prettier turns the
 * stylistic rules off so the two never fight). A handful of rules are tuned to
 * this codebase's established idioms (underscore-ignored args, deliberate
 * empty catch blocks that document a swallowed host error, the Office.js casts).
 */
export default tseslint.config(
  {
    ignores: [
      "dist/**",
      "dist-lib/**",
      "skill-dist/**",
      "coverage/**",
      "examples/**",
      "public/**",
      "node_modules/**",
      "**/*.snap",
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  prettier,
  {
    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.node,
        Office: "readonly",
        PowerPoint: "readonly",
        Excel: "readonly",
      },
    },
    rules: {
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrors: "none" },
      ],
      // The renderers reach preview-only Office.js APIs through `as unknown as`
      // casts; that's a deliberate, documented boundary, not sloppy typing.
      "@typescript-eslint/no-explicit-any": "off",
      // Swallowed host errors are documented with a comment in the catch body.
      "no-empty": ["error", { allowEmptyCatch: true }],
    },
  },
);
