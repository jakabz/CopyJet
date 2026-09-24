// Unit and contract tests for src/core (run with `npx jest`). Separate from the Heft/SPFx build.
/** @type {import('jest').Config} */
module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  testMatch: ['**/*.test.ts'],
  transform: {
    // PnPjs 4 ships ES modules only, so its .js files are transpiled as well.
    '^.+\\.(ts|tsx|js)$': ['ts-jest', { tsconfig: '<rootDir>/tsconfig.jest.json' }]
  },
  transformIgnorePatterns: ['/node_modules/(?!@pnp/)'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'json']
};
