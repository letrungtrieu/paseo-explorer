export function buildWordPreviewHtml(convertedWordHtml: string): string {
  // Mammoth does not sanitize document links. Block active content and remote assets in addition to the host sandbox.
  return `<!doctype html><html lang="vi"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body { margin: 0; padding: 24px; background: #e5e7eb; color: #111827; font: 15px/1.6 system-ui, sans-serif; }
main { box-sizing: border-box; max-width: 850px; min-height: 90vh; margin: auto; padding: 40px; background: white; overflow-wrap: anywhere; }
table { border-collapse: collapse; width: 100%; } td, th { border: 1px solid #9ca3af; padding: 6px; }
img { max-width: 100%; height: auto; } pre { white-space: pre-wrap; }
@media (max-width: 600px) { body { padding: 8px; } main { padding: 16px; } }
</style></head><body><main>${convertedWordHtml}</main></body></html>`;
}
