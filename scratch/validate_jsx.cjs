const fs = require('fs');
const content = fs.readFileSync('c:/Users/hemam/OneDrive/Desktop/rezolt/src/App.jsx', 'utf8');
const lines = content.split('\n');

let stack = [];
let inReturn = false;

for (let i = 2325; i < lines.length; i++) {
    const line = lines[i];
    const n = i + 1;
    
    if (line.includes('return (')) {
        inReturn = true;
        console.log(`L${n}: ENTERING RETURN BLOCK`);
        continue;
    }
    
    if (!inReturn) continue;
    
    // Ignore map arrays and objects as much as possible, but we need the JSX inside
    // This is hard with regex, so we just look for tag-like strings.
    
    // Match tags: <div, <h1, </div, etc.
    const tags = line.match(/<\/?([a-zA-Z0-9-]+)/g) || [];
    tags.forEach(tag => {
        if (tag.startsWith('</')) {
            const tagName = tag.substring(2);
            if (stack.length === 0) {
                 console.log(`L${n}: UNEXPECTED CLOSER </${tagName}> | line: ${line.trim()}`);
            } else {
                const last = stack.pop();
                if (last !== tagName) {
                    console.log(`L${n}: TAG MISMATCH. Found </${tagName}>, expected </${last}> | line: ${line.trim()}`);
                    stack.push(last); // restore
                }
            }
        } else {
            const tagName = tag.substring(1);
            // Check for self-closing: ends with /> (need to check the whole line or just the tag block)
            // A better check: find the tag in the line and see if there is a /> before the next > or next tag
            const tagIdx = line.indexOf(tag);
            const nextClose = line.indexOf('>', tagIdx);
            if (nextClose !== -1) {
                const tagContent = line.substring(tagIdx, nextClose + 1);
                if (tagContent.endsWith('/>')) {
                    // self-closing, don't push
                } else {
                    stack.push(tagName);
                }
            }
        }
    });

    if (line.includes(');') && stack.length === 0) {
        console.log(`L${n}: End of return block (Balanced)`);
        break;
    }
    if (line.includes(');') && stack.length > 0) {
         console.log(`L${n}: End of return block with UNCLOSED TAGS:`, stack);
         break;
    }
    if (line.includes('function ArticlesPage')) break;
}
