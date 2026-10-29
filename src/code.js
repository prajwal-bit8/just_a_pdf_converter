import {extension, fileStem} from './core.js';

export const CODE_LANGUAGES = {
  c:'C', h:'C header', cpp:'C++', cc:'C++', cxx:'C++', hpp:'C++ header', hxx:'C++ header',
  java:'Java', py:'Python', js:'JavaScript', mjs:'JavaScript', cjs:'JavaScript', jsx:'JSX',
  ts:'TypeScript', tsx:'TSX', cs:'C#', go:'Go', rs:'Rust', rb:'Ruby', php:'PHP',
  kt:'Kotlin', kts:'Kotlin', swift:'Swift', dart:'Dart', r:'R', scala:'Scala', lua:'Lua', pl:'Perl',
  html:'HTML', htm:'HTML', css:'CSS', scss:'SCSS', vue:'Vue', svelte:'Svelte',
  sql:'SQL', sh:'Shell', bash:'Bash', ps1:'PowerShell', bat:'Batch',
  asm:'Assembly', s:'Assembly', v:'Verilog', vh:'Verilog', sv:'SystemVerilog', svh:'SystemVerilog', ino:'Arduino',
  json:'JSON', xml:'XML', yaml:'YAML', yml:'YAML', md:'Markdown', txt:'Plain text'
};
export const CODE_EXTENSIONS = Object.keys(CODE_LANGUAGES);
export const CODE_FILE_LIMIT = 1024 * 1024;
export const isCodeTool = id => ['code-pdf','code-word','code-pages'].includes(id);
export const isCodeFile = name => CODE_EXTENSIONS.includes(extension(name));

export async function readCodeFile(file) {
  if (!isCodeFile(file.name)) throw new Error('Choose a source-code or plain-text file, such as .c, .cpp, .java, .py or .txt.');
  if (!file.size) throw new Error(`“${file.name}” is empty.`);
  if (file.size > CODE_FILE_LIMIT) throw new Error('Code files can be up to 1 MB each. Choose a smaller source file.');
  const bytes = new Uint8Array(await file.arrayBuffer());
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le' : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8';
  let text;
  try { text = new TextDecoder(encoding, {fatal:true}).decode(bytes); }
  catch { throw new Error(`“${file.name}” is not readable text. Save it as UTF-8 in your code editor and try again.`); }
  if (/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/.test(text)) throw new Error(`“${file.name}” contains binary data or unsupported control characters. Choose a source-code text file.`);
  if (!text.trim()) throw new Error(`“${file.name}” has no code or text to convert.`);
  return text.replace(/\r\n?/g, '\n');
}

export function codeOptions(opt = {}) {
  const size = Number(opt.fontSize || 10), tabSize = Number(opt.tabSize || 4);
  if (![8,10,12].includes(size) || ![2,4,8].includes(tabSize)) throw new Error('Choose a supported font and tab size.');
  const paper = opt.paper === 'LETTER' ? 'LETTER' : 'A4';
  const landscape = opt.orientation === 'landscape';
  const dimensions = paper === 'LETTER' ? [612,792] : [595.28,841.89];
  if (landscape) dimensions.reverse();
  return {fontSize:size, tabSize, paper, landscape, dimensions, lineNumbers:opt.lineNumbers !== 'no', title:String(opt.codeTitle || '').trim().slice(0,120)};
}

export function expandTabs(line, tabSize) {
  let column = 0, result = '';
  for (const c of line) {
    if (c === '\t') { const spaces = tabSize - column % tabSize; result += ' '.repeat(spaces); column += spaces; }
    else { result += c; column++; }
  }
  return result;
}

export function wrapCode(text, opt) {
  const lines = text.split('\n'), digits = String(lines.length).length;
  const gutter = opt.lineNumbers ? digits + 3 : 0;
  // A little extra room accommodates both the embedded PDF font and Word's Courier New.
  const columns = Math.max(20, Math.floor((opt.dimensions[0] - 84) / (opt.fontSize * .62)) - gutter);
  const rows = [];
  lines.forEach((line,index) => {
    const chars = Array.from(expandTabs(line,opt.tabSize));
    for (let start=0; start < Math.max(1,chars.length); start+=columns) {
      rows.push({number:start === 0 ? String(index+1) : '', text:chars.slice(start,start+columns).join('')});
      if (rows.length > 8000) throw new Error('This code is too long to format. Keep the task under 5,000 source lines and 8,000 wrapped lines.');
    }
  });
  return {rows, digits, lineCount:lines.length};
}

export async function prepareCode(files, options = {}) {
  if (!files.length || files.length > 10) throw new Error('Choose between 1 and 10 code files per document.');
  const opt = codeOptions(options), documents = [];
  let lines = 0, rows = 0, characters = 0;
  for (const file of files) {
    const text = await readCodeFile(file);
    lines += text.split('\n').length; characters += text.length;
    if (lines > 5000 || characters > 400000) throw new Error('Keep one code document under 5,000 source lines and 400,000 characters. Split larger projects into smaller tasks.');
    const layout = wrapCode(text,opt); rows += layout.rows.length;
    if (rows > 8000) throw new Error('Too many wrapped lines. Choose landscape, a smaller font, or fewer files.');
    documents.push({name:file.name, language:CODE_LANGUAGES[extension(file.name)], text, ...layout});
  }
  return {opt, documents, name:files.length === 1 ? fileStem(files[0].name) : 'code-listing'};
}

// Check the bundled TrueType font before PDF creation so unsupported characters
// produce an actionable message instead of silently becoming missing-glyph boxes.
export function missingPdfGlyph(fontBytes, text) {
  const view = new DataView(fontBytes.buffer, fontBytes.byteOffset, fontBytes.byteLength);
  let cmap = 0;
  for (let i=0; i<view.getUint16(4); i++) {
    const p=12+i*16;
    if (view.getUint32(p) === 0x636d6170) cmap=view.getUint32(p+8);
  }
  const tables=[];
  for (let i=0; i<view.getUint16(cmap+2); i++) {
    const p=cmap+4+i*8, platform=view.getUint16(p), encoding=view.getUint16(p+2);
    if (platform===0 || platform===3 && [1,10].includes(encoding)) tables.push(cmap+view.getUint32(p+4));
  }
  function hasGlyph(code) {
    return tables.some(p=>{
      const format=view.getUint16(p);
      if (format===12) {
        for (let i=0; i<view.getUint32(p+12); i++) {
          const q=p+16+i*12, first=view.getUint32(q), last=view.getUint32(q+4);
          if (code>=first && code<=last) return view.getUint32(q+8)+code-first!==0;
        }
      } else if (format===4 && code<=0xffff) {
        const n=view.getUint16(p+6)/2, end=p+14, start=end+n*2+2, delta=start+n*2, range=delta+n*2;
        for (let i=0; i<n; i++) if (code>=view.getUint16(start+i*2) && code<=view.getUint16(end+i*2)) {
          const d=view.getInt16(delta+i*2), offset=view.getUint16(range+i*2);
          if (!offset) return (code+d & 0xffff)!==0;
          const glyph=view.getUint16(range+i*2+offset+2*(code-view.getUint16(start+i*2)));
          return glyph!==0 && (glyph+d & 0xffff)!==0;
        }
      }
      return false;
    });
  }
  return [...new Set(text)].find(c=>c!=='\n' && c!=='\t' && !hasGlyph(c.codePointAt(0)));
}
