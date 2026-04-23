const fs = require('fs');
const content = fs.readFileSync('c:/Users/hemam/OneDrive/Desktop/rezolt/src/App.jsx', 'utf8');
const lines = content.split('\n');

let inLandingPage = false;
let depth = 0;

for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const n = i + 1;
    
    if (line.includes('function LandingPage')) {
        inLandingPage = true;
        continue;
    }
    
    if (!inLandingPage) continue;
    
    // Ignore comments
    if (line.trim().startsWith('//') || line.trim().startsWith('{/*')) continue;

    const opens = (line.match(/<div/g) || []).length;
    const closes = (line.match(/<\/div>/g) || []).length;
    const selfClosers = (line.match(/<div[^>]*\/>/g) || []).length;
    
    depth += (opens - selfClosers - closes);
    
    if (n >= 2327 && n <= 3020) {
        if (opens || closes) {
            console.log(`L${n} | D:${depth} | ${line.trim()}`);
        }
    }
    
    if (line.includes('function ArticlesPage')) break;
}
console.log('Final Depth:', depth);
