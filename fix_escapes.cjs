const fs = require('fs');
const filePath = 'ui/src/pages/Calculator/calculator.jsx';
let content = fs.readFileSync(filePath, 'utf8');
const re = /\u([0-9a-fA-F]{3})([^0-9a-fA-F])/g;
const before = (content.match(re) || []).length;
const prefix = ['\', 'u', '0'].join('');
content = content.replace(re, function(m, hex, after) {
  return prefix + hex + after;
});
const remaining = (content.match(re) || []).length;
fs.writeFileSync(filePath, content, 'utf8');
console.log('Fixed: ' + before + ', remaining: ' + remaining);
