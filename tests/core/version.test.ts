import { readFileSync } from 'fs';
import { join } from 'path';
import { COPYJET_VERSION, SOLUTION_VERSION } from '../../src/core/version';

const root = join(__dirname, '..', '..');
const json = (file: string): { [k: string]: unknown } => JSON.parse(readFileSync(join(root, file), 'utf8').replace(/^﻿/, ''));

describe('versions', () => {
  it('show the uploaded package version in the footer and write package.json\'s into templates', () => {
    const solution = json('config/package-solution.json').solution as { version: string; features?: Array<{ version: string }> };
    expect(SOLUTION_VERSION).toBe(solution.version);
    (solution.features || []).forEach((f) => expect(f.version).toBe(solution.version));
    expect(COPYJET_VERSION).toBe(json('package.json').version);
  });
});
