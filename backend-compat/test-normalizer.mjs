// Node 24 can strip the TypeScript types directly; no cloud access or credentials.
import assert from 'node:assert/strict';
import { normalizePrescription } from './rx-normalizer.ts';
const normalized = normalizePrescription({v: 4, date: '2026-09-18',
  doctor: {name: 'د. أحمد', specialty: 'GP'}, patient: {name: 'أحمد علي'},
  drugs: [{brand_name: 'TEST BRAND', generic_name: 'TEST SCIENTIFIC', dosage: 'TEST',
    frequency: '1x1', duration: 'TEST ONLY', notes: 'مع الطعام', quantity: '0'},
    {brand_name: 'BRAND ONLY'}, {generic_name: 'SCIENTIFIC ONLY'}]});
assert.equal(normalized.patient, 'أحمد علي');
assert.equal(normalized.doctor, 'د. أحمد · GP');
assert.equal(normalized.medications[0].name, 'TEST BRAND — TEST SCIENTIFIC');
assert.equal(normalized.medications[0].instructions, '1x1 · TEST ONLY · مع الطعام · Quantity: 0');
assert.equal(normalized.medications.length, 3);
assert.equal(normalized.medications[1].name, 'BRAND ONLY');
assert.equal(normalized.medications[2].name, 'SCIENTIFIC ONLY');
const legacy = {patient: 'TEST PATIENT', doctor: 'TEST DOCTOR', medications: [
  {name: 'TEST MEDICINE', dosage: 'TEST', instructions: 'DO NOT USE'}]};
assert.deepEqual(normalizePrescription(legacy).medications, legacy.medications);
for (const invalid of [null, [], 'text', {}, {drugs: 'text'}]) {
  assert.equal(normalizePrescription(invalid), undefined);
}
assert.deepEqual(normalizePrescription({drugs: [null, {name: {bad: true}}, {brand_name: 'Brand'}]}).medications,
  [{name: 'Brand', dosage: undefined, instructions: ''}]);
console.log('Backend normalizer: v4, Arabic, brand/scientific-only, legacy and malformed-data checks passed');
