/**
 * Decode a base64 string to bytes using the web-standard `atob` — no `node:`
 * dependency, so it works the same in Node, browsers, and edge runtimes.
 * (`Uint8Array.fromBase64` is not yet available on our Node 24 target.)
 */
export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}
