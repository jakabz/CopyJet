import { readFileSync } from 'fs';
import { join } from 'path';
import { openTemplate } from '../../src/core/packager';
import { buildPlan } from '../../src/core/planner';

// Structure of the first template exported from the Forrás test site (tenant anonymised).
const text = readFileSync(join(__dirname, '..', 'fixtures', 'templates', 'forras-sablon.json'), 'utf8');

describe('template exported by the Setup (Forrás test site)', () => {
  it('opens, validates and plans: lists and site columns first, then columns, then views', async () => {
    const reader = await openTemplate(new Blob([text]));
    const plan = buildPlan(reader.manifest);
    expect(plan.excluded).toEqual([]);
    expect(plan.levels.map((l) => l.map((s) => s.ref.key))).toEqual([
      ['field:SiteColumn1', 'field:SiteColumn2', 'list:Teszt_lista', 'list:Teszt_lookup_forrs'],
      [
        'listField:Teszt_lista/ListaLookup',
        'listField:Teszt_lista/ListaSzoveg',
        'listField:Teszt_lista/Lookup',
        'listField:Teszt_lista/SiteColumn1',
        'listField:Teszt_lista/SiteColumn2',
        'listField:Teszt_lista/V_x00e1_lassz',
        'listField:Teszt_lista/Valaki'
      ],
      ['view:Teszt_lista/Minden elem', 'view:Teszt_lista/Nyitott elemek', 'view:Teszt_lista/Ügyfél nézet']
    ]);
  });
});
