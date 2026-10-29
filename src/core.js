export const MAX_FILE_SIZE = 50 * 1024 * 1024;
export const MAX_TOTAL_SIZE = 150 * 1024 * 1024;
export function parseRanges(input, count) {
  if (!Number.isInteger(count) || count < 1) throw new Error('This PDF has no pages.');
  if (!String(input).trim()) return Array.from({length: count}, (_, i) => i);
  const pages = [];
  for (const token of String(input).split(',')) {
    const match = token.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!match) throw new Error('Use page numbers like 1, 3-5, 8.');
    const start = Number(match[1]), end = match[2] ? Number(match[2]) : start;
    if (start < 1 || end > count || start > end) throw new Error(`Page ranges must be between 1 and ${count}, in ascending order.`);
    for (let n = start; n <= end; n++) pages.push(n - 1);
  }
  return [...new Set(pages)];
}
export function fileStem(name) { return name.replace(/\.[^.]+$/, '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') || 'document'; }
export function humanSize(n) { return n < 1024 ? `${n} B` : n < 1048576 ? `${(n/1024).toFixed(1)} KB` : `${(n/1048576).toFixed(1)} MB`; }
export function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
export function extension(name) { return name.split('.').pop().toLowerCase(); }
export function validateFiles(files, allowed, existing = []) {
  if (!files.length) throw new Error('Choose a file to get started.');
  for (const file of files) {
    if (!allowed.includes(extension(file.name))) throw new Error(`“${file.name}” is not supported here. Choose ${allowed.map(x=>'.'+x).join(', ')} files.`);
    if (!file.size) throw new Error(`“${file.name}” is empty.`);
    if (file.size > MAX_FILE_SIZE) throw new Error(`“${file.name}” is over 50 MB. Choose a smaller file.`);
  }
  if ([...existing, ...files].reduce((n,f)=>n+f.size,0) > MAX_TOTAL_SIZE) throw new Error('Keep the combined file size under 150 MB.');
}
export function pdfLines(items) {
  const lines = [];
  for (const item of items) {
    if (typeof item.str !== 'string') continue;
    const y = item.transform?.[5] || 0, x = item.transform?.[4] || 0;
    let row = lines.find(r=>Math.abs(r.y-y)<2);
    if(!row) { row = {y, items:[]}; lines.push(row); }
    row.items.push({x, str:item.str, width:item.width||0, size:Math.abs(item.transform?.[3]||11)});
  }
  return lines.sort((a,b)=>b.y-a.y).map(row=>{
    const a=row.items.sort((x,y)=>x.x-y.x); let text='';
    a.forEach((item,i)=>{ if(i && !text.endsWith(' ') && !item.str.startsWith(' ') && item.x-a[i-1].x-a[i-1].width>1) text+=' '; text+=item.str; });
    return {text:text.trim(),size:Math.max(...a.map(x=>x.size))};
  }).filter(r=>r.text);
}
