// Extracts the GC object from moadilim-poc.html and writes gameContent.json
// Run: node export_game_content.js

const fs = require('fs');
const path = require('path');

const htmlPath = path.join(__dirname, 'moadilim-poc.html');
const outPath  = path.join(__dirname, 'src', 'data', 'gameContent.json');

const html = fs.readFileSync(htmlPath, 'utf8');

const marker = 'const GC = ';
const start  = html.indexOf(marker);
if (start === -1) {
  console.error('ERROR: Could not find "const GC = " in moadilim-poc.html');
  process.exit(1);
}

// Walk forward with bracket counting to find the matching closing }
const fromGC = html.slice(start + marker.length);
let depth = 0, i = 0, inString = false, escaped = false;

for (; i < fromGC.length; i++) {
  const ch = fromGC[i];
  if (escaped)          { escaped = false; continue; }
  if (ch === '\\' && inString) { escaped = true; continue; }
  if (ch === '"')       { inString = !inString; continue; }
  if (inString)         continue;
  if (ch === '{')       depth++;
  else if (ch === '}')  { depth--; if (depth === 0) { i++; break; } }
}

const gcStr = fromGC.slice(0, i);

let gc;
try {
  gc = JSON.parse(gcStr);
} catch (e) {
  console.error('ERROR: Failed to parse GC object as JSON:', e.message);
  process.exit(1);
}

fs.writeFileSync(outPath, JSON.stringify(gc, null, 2), 'utf8');
console.log(`✓ gameContent.json updated (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
