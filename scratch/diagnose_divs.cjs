const fs = require('fs');
const content = fs.readFileSync('c:/Users/hemam/OneDrive/Desktop/rezolt/src/App.jsx', 'utf8');
const lines = content.split('\n');

let stack = [];

lines.forEach((l, i) => {
    let pos = 0;
    while (pos < l.length) {
        let openIdx = l.indexOf('<div', pos);
        let closeIdx = l.indexOf('</div>', pos);

        if (openIdx === -1 && closeIdx === -1) break;

        if (openIdx !== -1 && (closeIdx === -1 || openIdx < closeIdx)) {
            // Check if self-closing
            let endIdx = l.indexOf('>', openIdx);
            if (endIdx !== -1) {
                let tag = l.substring(openIdx, endIdx + 1);
                if (!tag.endsWith('/>')) {
                    stack.push({ line: i + 1, content: l.trim() });
                }
            }
            pos = openIdx + 4;
        } else {
            stack.pop();
            pos = closeIdx + 6;
        }
    }
});

console.log('Final Level:', stack.length);
stack.forEach(s => console.log(`Unclosed at line ${s.line}: ${s.content}`));
