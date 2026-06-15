/** Tiny crypto-based nanoid replacement (no external dep needed) */
export function nanoid(size = 21): string {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
  const bytes = crypto.getRandomValues(new Uint8Array(size));
  return Array.from(bytes, (b) => alphabet[b % alphabet.length]).join("");
}
