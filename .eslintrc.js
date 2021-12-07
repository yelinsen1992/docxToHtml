module.exports = {
  root: true,
  env: {
    node: true
  },
  extends: 'standard',
  rules: {
    'no-new': 'off',
    'space-before-function-paren': 0,
    'comma-dangle': 0,
    semi: [2, 'never'],
    indent: ['error', 2, { SwitchCase: 1 }],
    'no-var': 'error',
    'arrow-parens': 0,
    'generator-star-spacing': 0,
    'prefer-const': ['error', {
      destructuring: 'any',
      ignoreReadBeforeAssign: false
    }],
    'no-debugger': process.env.NODE_ENV === 'production' ? 2 : 0
  }
}
