const fs = require('fs');
const content = fs.readFileSync('c:/Users/hemam/OneDrive/Desktop/rezolt/src/App.jsx', 'utf8');
const lines = content.split('\n');

let divDepth = 0;
let braceDepth = 0;
let inLandingPage = false;

for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const n = i + 1;
    
    if (line.includes('function LandingPage')) {
        inLandingPage = true;
        braceDepth = 1; // Function start
    }
    
    if (!inLandingPage) continue;

    // Count braces (ignoring strings/regex for now)
    const lineWithoutStrings = line.replace(/(["'])(?:(?=(\\?))\2.)*?\1/g, "");
    const ob = (lineWithoutStrings.match(/\{/g) || []).length;
    const cb = (lineWithoutStrings.match(/\}/g) || []).length;
    braceDepth += (ob - cb);

    // Only count div tags in JSX (simplified)
    const odiv = (line.match(/<div/g) || []).length;
    const cdiv = (line.match(/<\/div>/g) || []).length;
    divDepth += (odiv - cdiv);

    if (n >= 2327 && n <= 3016) {
        if (odiv !== 0 || cdiv !== 0 || ob !== 0 || cb !== 0) {
            console.log(`L${n} | D:${divDepth} B:${braceDepth} | ${line.trim()}`);
        }
    }

    if (braceDepth === 0 && line.includes('}')) {
        console.log(`POENTIAL FUNCTION END AT L${n}`);
        break;
    }
}
