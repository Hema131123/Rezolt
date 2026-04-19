const fs = require('fs');
let code = fs.readFileSync('src/App.jsx', 'utf8');
code = code.replace(/\\n\\nexport default function App/, '\n\nexport default function App');
fs.writeFileSync('src/App.jsx', code);
