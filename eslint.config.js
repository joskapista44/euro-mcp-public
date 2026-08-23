// Three named rules, deliberately NOT extending a preset -- these are the exact LINT-lane
// defects a quality gate targets (eval(), == instead of ===, unused vars), nothing broader.
//
// ONE block, not two: a TS-aware setup would split TS (parser + @typescript-eslint/no-unused-vars)
// from plain JS/MJS, but only makes sense where both exist. This repo has ZERO .ts files
// (measured: 8 .cjs | 19 .js | 11 .mjs | 0 .ts, node_modules/.git excluded) -- a TS parser here
// would be unexercised config, not coverage.
//
// `eqeqeq` carries a `null: 'ignore'` exception. A corpus run found exactly one eqeqeq finding,
// and it is the `!= null` idiom in mcp.js (checks a value against both `null` and `undefined` in
// one comparison). Fixing it to `!==` would be a BEHAVIOR CHANGE, not a lint fix -- `undefined`
// would stop matching. This is eslint's own documented exception, not the rule turned off: every
// OTHER `==`/`!=` (including a loose comparison where neither side is `null`) is still an error.
export default [
  {
    files: ['**/*.js', '**/*.cjs', '**/*.mjs'],
    rules: {
      'no-eval': 'error',
      eqeqeq: ['error', 'always', { null: 'ignore' }],
      'no-unused-vars': 'error',
    },
  },
];
