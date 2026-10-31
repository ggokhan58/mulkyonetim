interface PaymentSummary {
  payment_status: string;
  last_amount_paid?: number | null;
}

// Tahsilat, güncel sözleşme tutarından yeniden hesaplanmaz.
export function collectedAmount(property: PaymentSummary): number {
  return property.payment_status === 'Ödendi' ? Number(property.last_amount_paid ?? 0) : 0;
}

export function initialPaymentAmount(
  payment: { amount_paid: number | null } | undefined,
  currentRent: string | number,
  year: number,
  monthIndex: number,
  today = new Date(),
): string {
  if (payment) return String(payment.amount_paid ?? 0);
  // Geçmiş dönemin sözleşme bedeli bilinmiyorsa kullanıcı gerçek tutarı girer.
  if (year * 12 + monthIndex < today.getFullYear() * 12 + today.getMonth()) return '';
  return String(Number(currentRent) || 0);
}
