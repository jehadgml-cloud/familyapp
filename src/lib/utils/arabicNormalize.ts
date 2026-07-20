/**
 * Normalizes Arabic string characters (e.g. Alif variants, Teh Marbuta, Yeh/Alef Maksura)
 * to standard base letters for accurate matching across names.
 */
export function normalizeArabicName(text: string): string {
  if (!text) return ''
  return text
    .trim()
    .toLowerCase()
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/[ًٌٍَُِّْ]/g, '') // Remove Harakat (diacritics)
}
