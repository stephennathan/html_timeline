/**
 * lint-staged configuration for html-timeline module.
 */

export default {
  '*.ts': () => 'npm run typecheck',
  '*.{js,ts}': (files) => `npx eslint --fix ${files.join(' ')}`,
  '*.{ts,js,json,md}': (files) => `npx prettier --write ${files.join(' ')}`,
};
