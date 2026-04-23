const fs = require('fs');
const content = fs.readFileSync('c:/Users/hemam/OneDrive/Desktop/rezolt/src/App.jsx', 'utf8');
const lines = content.split('\n');

let depth = 0;
let inLandingPage = false;

lines.forEach((line, i) => {
    const lineNum = i + 1;
    if (line.includes('function LandingPage')) inLandingPage = true;
    if (!inLandingPage) return;

    // Simple regex for divs
    const opens = (line.match(/<div/g) || []).length;
    const closes = (line.match(/<\/div>/g) || []).length;
    
    const prevDepth = depth;
    depth += (opens - closes);
    
    if (lineNum >= 2118 && lineNum <= 3020 && (opens !== 0 || closes !== 0)) {
        console.log(`L${lineNum}: ${prevDepth} -> ${depth} | ${line.trim()}`);
    }
    
    if (line.includes(');') && depth === 0 && lineNum > 3000) {
        console.log(`POENTIAL RETURN END AT L${lineNum}`);
        inLandingPage = false;
    }
});
console.log('Final depth:', depth);
