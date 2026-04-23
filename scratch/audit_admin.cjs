const fs = require('fs');
const content = fs.readFileSync('src/App.jsx', 'utf8');
const lines = content.split('\n');

let depth = 0;
let inAdmin = false;

for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.includes('function AdminPage')) inAdmin = true;
    if (line.includes('function syncPaymentProfile')) inAdmin = false; // Next function starts

    if (!inAdmin) continue;

    const openCount = (line.match(/<div|<section|<header|<footer|<main|<aside/g) || []).length;
    const closeCount = (line.match(/<\/div>|<\/section>|<\/header>|<\/footer>|<\/main>|<\/aside>/g) || []).length;
    const oldDepth = depth;
    depth += openCount - closeCount;

    if (openCount > 0 || closeCount > 0) {
        console.log(`L${i + 1} | D:${oldDepth} -> ${depth} | ${line.trim().substring(0, 40)}`);
    }
}
console.log("Final Admin Depth:", depth);
