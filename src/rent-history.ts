export interface RentPeriod { effective_date: string; amount: number }
export function localDate(today = new Date()): string {
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
}
export function readRentHistory(value?: string): RentPeriod[] {
  return value ? JSON.parse(value) : [];
}
export function rentOnDate(history: RentPeriod[], date: string): number | undefined {
  return [...history].reverse().find(period => period.effective_date <= date)?.amount;
}
export function periodRent(history: RentPeriod[], year: number, month: number, dueDay: number): number | undefined {
  const day = Math.min(Math.max(dueDay || 1, 1), new Date(year, month + 1, 0).getDate());
  return rentOnDate(history, `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`);
}
export function updatedRentHistory(previous: any, property: any): string {
  const history = readRentHistory(previous?.rent_history);
  const amount = Number(property.rent_amount);
  if (!Number.isFinite(amount) || amount < 0) throw new Error('Kira tutarı sıfır veya pozitif olmalıdır.');
  if (previous && Number(previous.rent_amount) === amount && !property.rent_effective_date) return JSON.stringify(history);
  const date = property.rent_effective_date || (!previous ? property.contract_start_date || localDate() : '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isFinite(Date.parse(date)) || new Date(date).toISOString().slice(0, 10) !== date)
    throw new Error('Geçerli bir sözleşme güncelleme tarihi girin.');
  if (history.length && date <= history[history.length - 1].effective_date)
    throw new Error('Yeni tarih son sözleşme tarihinden sonra olmalıdır.');
  // Eski tutarın geçmişte ne zaman başladığı bilinmez; yalnızca bugünden itibaren koru.
  if (!history.length && previous && localDate() < date)
    history.push({ effective_date: localDate(), amount: Number(previous.rent_amount) });
  history.push({ effective_date: date, amount });
  return JSON.stringify(history);
}
