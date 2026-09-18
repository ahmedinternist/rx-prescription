/** Normalize minimal desktop v4 data without changing stored records or expiry. */
type RecordValue = Record<string, unknown>;
export interface ViewerPrescription {
  patient: string;
  doctor: string;
  date?: string;
  medications: { name: string; dosage?: string; instructions: string }[];
}
const record = (value: unknown): RecordValue | undefined =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? value as RecordValue : undefined;
const text = (value: unknown): string => typeof value === 'string' ? value.trim() : '';

export function normalizePrescription(value: unknown): ViewerPrescription | undefined {
  const rx = record(value);
  if (!rx) return undefined;
  const doctor = record(rx.doctor);
  const patient = record(rx.patient);
  const drugs = Array.isArray(rx.drugs) ? rx.drugs : rx.medications;
  if (!Array.isArray(drugs)) return undefined;
  return {
    doctor: doctor ? [text(doctor.name), text(doctor.specialty)].filter(Boolean).join(' · ') : text(rx.doctor),
    patient: patient ? text(patient.name) : text(rx.patient),
    date: text(rx.date) || undefined,
    medications: drugs.flatMap(value => {
      const med = record(value);
      if (!med) return [];
      const brand = text(med.brand_name);
      const scientific = text(med.generic_name);
      const name = [brand, scientific && scientific !== brand ? scientific : ''].filter(Boolean).join(' — ') || text(med.name);
      if (!name) return [];
      return [{
        name,
        dosage: text(med.dosage) || undefined,
        instructions: [text(med.instructions), text(med.frequency), text(med.duration),
          text(med.notes), text(med.quantity) ? `Quantity: ${text(med.quantity)}` : ''].filter(Boolean).join(' · '),
      }];
    }),
  };
}
