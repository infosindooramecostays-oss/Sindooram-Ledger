// Turns system-prompt.md into a JS module the function imports, so the
// prompt is bundled into the function itself — edit the .md, push, and
// Netlify rebuilds with the new prompt.
import { readFileSync, writeFileSync } from 'node:fs';

const prompt = readFileSync(new URL('../system-prompt.md', import.meta.url), 'utf8');
writeFileSync(
  new URL('../netlify/functions/prompt.generated.mjs', import.meta.url),
  '// Generated from system-prompt.md by scripts/build-prompt.mjs — do not edit.\n' +
  'export default ' + JSON.stringify(prompt) + ';\n'
);
console.log('system prompt bundled (' + prompt.length + ' chars)');
