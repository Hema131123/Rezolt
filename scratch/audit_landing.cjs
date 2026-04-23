const fs = require('fs');
const content = fs.readFileSync('c:/Users/hemam/OneDrive/Desktop/rezolt/src/App.jsx', 'utf8');
const lines = content.split('\n');

let inLandingPage = false;
let depth = 0;

console.log('--- STARTING AUDIT ---');

for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const n = i + 1;
    
    if (line.includes('function LandingPage')) {
        inLandingPage = true;
        console.log(`L${n}: Found function LandingPage`);
        continue;
    }
    
    if (!inLandingPage) continue;
    
    // Check for return statement - depth should be 0 here!
    if (line.includes('return (') && n < 2400) {
        console.log(`L${n}: ENTERING RETURN BLOCK. Current Depth: ${depth}`);
    }

    // Ignore comments
    if (line.trim().startsWith('//') || line.trim().startsWith('{/*')) continue;

    // Count div tags
    const opens = (line.match(/<div/g) || []).length;
    const closes = (line.match(/<\/div>/g) || []).length;
    
    // Check for self-closing divs <div ... />
    const selfClosers = (line.match(/<div[^>]*\/>/g) || []).length;
    
    const effectiveOpens = opens - selfClosers;
    
    const oldDepth = depth;
    depth += (effectiveOpens - closes);
    
    if (depth !== oldDepth && inLandingPage) {
       // Only log if we are before the return or in a suspected problematic area
       if (n < 2330 || depth < 0 || depth > 20) {
           console.log(`L${n}: ${oldDepth} -> ${depth} | ${line.trim()}`);
       }
    }

    if (line.includes('function ArticlesPage')) break;
}

console.log('--- AUDIT FINISHED ---');
console.log('Final Depth:', depth);
