const fs = require('fs');
const jsdom = require('jsdom');
const reader = require('xlsx');

const file = fs.readFileSync(`test.html`);

const doc = new jsdom.JSDOM(file.toString(), 'text/html');
const r = doc.window.document.querySelectorAll('.prqacontent > pre');
const data = [];
for (let i = 0; i < r.length; i++) {
  let str = '';
  for (let j = 0; j < r[i]?.children.length; j++) {
    if (r[i]?.children[j].nodeName === 'BR') break;
    str += r[i]?.children[j].textContent.replace('(^)', '') + '^';
  }
  if (str) {
    const obj = {};
    ['(', ',', ')', 'M3CM', '^', '^', '^', '^'].forEach((i, idx) => {
      const foundIndex = i === '.' ? str.lastIndexOf(i) : str.indexOf(i);
      if (foundIndex > -1) {
        obj[idx] = str.substring(0, foundIndex);
        str = str.substring(
          i === 'M3CM' ? foundIndex : foundIndex + 1,
          str.length
        );
      }
    });
    obj['last'] = str;
    data.push(obj);
  }
}
// Reading our test file
const xlsx = reader.readFile('./result.xlsx');

const ws = reader.utils.json_to_sheet(data);

reader.utils.book_append_sheet(xlsx, ws, 'Sheet2');

// Writing to our file
reader.writeFile(xlsx, './result.xlsx');
