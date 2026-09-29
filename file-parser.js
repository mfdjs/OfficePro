

const path = require('path');
const fs = require('fs');
const zlib = require('zlib');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const JSZip = require('jszip');

let WordExtractor = null;
try { WordExtractor = require('word-extractor'); } catch(e) { WordExtractor = null; }

const TEXT_EXTENSIONS = [
  '.txt','.text','.md','.markdown','.mdx','.rst','.org','.adoc','.asciidoc','.tex','.bib','.typ',
  '.csv','.tsv','.tab','.psv','.json','.jsonl','.ndjson','.jsonc','.geojson','.json5','.yaml','.yml','.toml','.ini','.conf','.config','.cfg','.properties','.prop','.env','.env.local',
  '.xml','.xsd','.xsl','.xslt','.svgz-text','.plist','.rss','.atom','.opml','.drawio','.mm','.gml','.kml','.gpx','.tcx',
  '.html','.htm','.xhtml','.shtml','.htm','.mht','.mhtml','.eml','.msg',
  '.rtf','.log','.err','.out','.dmesg','.nfo','.ans',

  '.js','.mjs','.cjs','.jsx','.ts','.tsx','.vue','.svelte','.astro','.py','.pyw','.ipynb','.java','.kt','.kts','.scala','.groovy','.clj','.cljs','.cljc','.ex','.exs','.erl','.hrl','.hs','.lhs','.elm','.jl','.lua','.rb','.php','.go','.rs','.swift','.objc','.m','.mm','.c','.h','.cpp','.cc','.cxx','.hpp','.hh','.hxx','.cs','.vb','.fs','.dart','.r','.R','.m','.asm','.s','.S','.v','.sv','.svh','.vhd','.vhdl','.uc','.sol','.move','.zig','.nim','.cr','.ml','.fsx','.fsi','.f90','.f95','.f03','.pas','.p','.pp','.d','.tcl','.pl','.pm','.lua','.sh','.bash','.zsh','.fish','.bat','.cmd','.ps1','.psm1','.vbs','.ahk','.awk','.sed','.vim','.el','.scm','.lisp','.rkt','.exs','.gleam',

  '.css','.scss','.sass','.less','.styl','.postcss','.graphql','.gql','.proto','.thrift','.wasm','.wat','.prisma',

  '.sql','.ddl','.gradle','.maven','.pom','.cmake','.mk','.mak','.makefile','.dockerfile','.dockerignore','.gitignore','.gitattributes','.editorconfig','.sln','.csproj','.vbproj','.fsproj','.vcxproj','.pbxproj','.gradle','.gemspec','.rake','.podspec','.manifest','.lock','.diff','.patch','.rej','.map','.sourcemap','.min','.lic','.license','.notice','.authors','.changelog','.todo','.faq',

  '.srt','.ass','.ssa','.vtt','.lrc','.sub','.smi',

  '.abnf','.ebnf','.bnf','.regex','.csv','.ics','.vcf','.cron','.sysctl','.resolv'
];

const IMAGE_EXTENSIONS = ['.jpg','.jpeg','.png','.bmp','.gif','.webp','.tiff','.tif','.ico','.svg','.heic','.heif','.avif','.jxl','.jp2','.wmf','.emf','.psd','.ai','.sketch'];

const WORD_EXTENSIONS = ['.docx','.doc','.docm','.dotx','.dotm','.dot','.odt','.ott','.wps','.wpt','.rtf','.pages'];

const SHEET_EXTENSIONS = ['.xlsx','.xls','.xlsm','.xlsb','.xlt','.xltx','.xltm','.ods','.ots','.csv','.tsv','.et','.ett','.numbers','.dif','.prn','.slk'];

const SLIDE_EXTENSIONS = ['.pptx','.ppt','.pptm','.pot','.potx','.potm','.pps','.ppsx','.odp','.otp','.dps','.dpt','.key'];

const PDF_EXTENSIONS = ['.pdf','.xps','.oxps'];

const EBOOK_EXTENSIONS = ['.epub','.mobi','.azw','.azw3','.azw4','.fb2','.fb2.zip','.lit','.pdb','.txt'];

const ARCHIVE_EXTENSIONS = ['.zip','.zipx','.jar','.war','.ear','.apk','.aab','.ipa','.crx','.xpi','.7z','.rar','.tar','.gz','.tgz','.bz2','.tbz2','.xz','.txz','.z','.lz','.lzma','.cab','.iso','.dmg','.xmind','.vsdx'];

const AUDIO_EXTENSIONS = ['.mp3','.wav','.flac','.ape','.aac','.ogg','.oga','.m4a','.wma','.amr','.opus','.aiff','.aif','.mid','.midi','.au','.ra','.wv'];

const VIDEO_EXTENSIONS = ['.mp4','.m4v','.mov','.mkv','.avi','.wmv','.flv','.webm','.mpg','.mpeg','.m2ts','.mts','.ts','.3gp','.rmvb','.rm','.vob','.ogv'];

const FONT_EXTENSIONS = ['.ttf','.otf','.woff','.woff2','.eot','.ttc'];

const MARKDOWN_EXTENSIONS = ['.md','.markdown','.mdx','.rst','.org'];

const MIME = {
  '.jpg':'image/jpeg','.jpeg':'image/jpeg','.png':'image/png','.gif':'image/gif','.webp':'image/webp',
  '.bmp':'image/bmp','.svg':'image/svg+xml','.ico':'image/x-icon','.tif':'image/tiff','.tiff':'image/tiff',
  '.avif':'image/avif','.jfif':'image/jpeg','.heic':'image/heic','.heif':'image/heif',
  '.mp3':'audio/mpeg','.wav':'audio/wav','.flac':'audio/flac','.aac':'audio/aac','.ogg':'audio/ogg',
  '.m4a':'audio/mp4','.wma':'audio/x-ms-wma','.opus':'audio/ogg','.mid':'audio/midi','.aiff':'audio/aiff',
  '.mp4':'video/mp4','.m4v':'video/mp4','.mov':'video/quicktime','.mkv':'video/x-matroska','.avi':'video/x-msvideo',
  '.wmv':'video/x-ms-wmv','.flv':'video/x-flv','.webm':'video/webm','.mpg':'video/mpeg','.mpeg':'video/mpeg',
  '.3gp':'video/3gpp','.ogv':'video/ogg','.ts':'video/mp2t'
};

function decodeTextBuffer(buf) {
  if (!buf || !buf.length) return '';

  if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) return buf.slice(3).toString('utf8');
  if (buf.length >= 2 && buf[0] === 0xFF && buf[1] === 0xFE) return new TextDecoder('utf-16le').decode(buf.slice(2));
  if (buf.length >= 2 && buf[0] === 0xFE && buf[1] === 0xFF) return new TextDecoder('utf-16be').decode(buf.slice(2));

  let utf8str = null;
  try { utf8str = new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { utf8str = null; }
  let gbkStr = null;
  try { gbkStr = new TextDecoder('gbk').decode(buf); } catch (e) { gbkStr = null; }
  const cjkCount = s => (s.match(/[一-鿿　-〿＀-￯]/g) || []).length;
  if (utf8str != null) {

    if (gbkStr && gbkStr.indexOf('\uFFFD') === -1 && cjkCount(gbkStr) > cjkCount(utf8str) && cjkCount(utf8str) === 0) return gbkStr;
    return utf8str;
  }

  if (gbkStr != null) return gbkStr;
  for (const enc of ['big5','shift-jis','windows-1252']) {
    try { return new TextDecoder(enc).decode(buf); } catch (e) {}
  }
  return buf.toString('latin1');
}

function escapeHtml(text) {
  return String(text == null ? '' : text)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function htmlToText(html) {
  return String(html)
    .replace(/<\s*br\s*\/?\s*>/gi,'\n')
    .replace(/<\/(p|div|h[1-6]|li|tr|section|article)>/gi,'\n')
    .replace(/<\/t[dh]>/gi,'\t')
    .replace(/<[^>]+>/g,'')
    .replace(/&nbsp;/g,' ').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
    .replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(+n))
    .replace(/\n{3,}/g,'\n\n').trim();
}

function isProbablyText(buf) {
  if (!buf || !buf.length) return true;
  const sample = buf.slice(0, Math.min(buf.length, 65536));
  let nul = 0, ctrl = 0, high = 0;
  for (let i = 0; i < sample.length; i++) {
    const b = sample[i];
    if (b === 0) nul++;
    else if (b < 9 || (b > 13 && b < 32)) ctrl++;
    else if (b >= 128) high++;
  }
  if (nul / sample.length > 0.01) return false;
  if (ctrl / sample.length > 0.05) return false;
  return true;
}

function hexDump(buf, maxBytes) {
  const n = Math.min(buf.length, maxBytes || 4096);
  const rows = [];
  for (let off = 0; off < n; off += 16) {
    const chunk = buf.slice(off, off + 16);
    const hex = [], asc = [];
    for (let i = 0; i < 16; i++) {
      if (i < chunk.length) {
        hex.push(chunk[i].toString(16).padStart(2,'0'));
        asc.push(chunk[i] >= 32 && chunk[i] < 127 ? String.fromCharCode(chunk[i]) : '·');
      } else { hex.push('  '); asc.push(' '); }
    }
    rows.push('<div class="hex-row"><span class="hex-off">'+off.toString(16).padStart(8,'0')+'</span><span class="hex-bytes">'+hex.join(' ')+'</span><span class="hex-ascii">'+asc.join('')+'</span></div>');
  }
  return rows.join('');
}

function extractStrings(buf, minLen) {
  const out = [];

  let cur = '';
  for (let i = 0; i < buf.length; i++) {
    const b = buf[i];
    if (b >= 32 && b < 127) cur += String.fromCharCode(b);
    else { if (cur.length >= (minLen||5)) out.push(cur); cur = ''; }
  }
  if (cur.length >= (minLen||5)) out.push(cur);
  return out.slice(0, 400);
}

function wordHtml(content, fileName, fileSize, extra) {
  return Object.assign({ type:'word', content, fileName, fileSize }, extra||{});
}
function pre(text, mono) {
  const fam = mono ? 'Consolas,"Courier New",monospace;font-size:13px;' : 'inherit;';
  return '<pre style="white-space:pre-wrap;word-break:break-word;font-family:'+fam+'line-height:1.8;margin:0">'+escapeHtml(text)+'</pre>';
}

function plainTextToHtml(text) {
  if (!text || !String(text).trim()) return '<p>(文档内容为空)</p>';
  return String(text).replace(/\r\n?/g,'\n').split('\n').map(line=>{
    if (!line.trim()) return '';
    return '<p style="line-height:1.9;margin:.35em 0;white-space:pre-wrap;tab-size:8;">'+escapeHtml(line)+'</p>';
  }).join('');
}

const CP1252_FORWARD = {0x80:0x20AC,0x82:0x201A,0x83:0x192,0x84:0x201E,0x85:0x2026,0x86:0x2020,0x87:0x2021,0x88:0x2C6,0x89:0x2030,0x8A:0x160,0x8B:0x2039,0x8C:0x152,0x8E:0x17D,0x91:0x2018,0x92:0x2019,0x93:0x201C,0x94:0x201D,0x95:0x2022,0x96:0x2013,0x97:0x2014,0x98:0x2DC,0x99:0x2122,0x9A:0x161,0x9B:0x203A,0x9C:0x153,0x9E:0x17E,0x9F:0x178};
const CP1252_REV = {}; Object.keys(CP1252_FORWARD).forEach(b=>{ CP1252_REV[CP1252_FORWARD[b]] = Number(b); });
function cjkCountOf(s){ return (String(s).match(/[一-鿿]/g)||[]).length; }

function repairCp1252Gbk(text) {
  if (!text || text.length < 4) return text;
  let bad = 0;
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if ((c>=0xC0&&c<=0xFF)||(c>=0x100&&c<=0x17F)||(c>=0x250&&c<=0x2FF)) bad++;
  }
  const cjk = cjkCountOf(text);
  if (bad < 4 || bad <= cjk) return text;
  const bytes = [];
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if (CP1252_REV[c] != null) bytes.push(CP1252_REV[c]);
    else if (c <= 0xFF) bytes.push(c);
    else return text;
  }
  const fixed = decodeTextBuffer(Buffer.from(bytes));
  const fc = cjkCountOf(fixed);
  const fffd = (fixed.match(/\uFFFD/g)||[]).length;
  let fixedBad = 0;
  for (const ch of fixed) { const c=ch.codePointAt(0); if((c>=0xC0&&c<=0xFF)||(c>=0x100&&c<=0x17F)) fixedBad++; }
  if (fffd===0 && fc>=4 && fc>cjk && fc>=bad*0.6 && fixedBad<bad*0.2) return fixed;
  return text;
}
function humanSize(n) {
  if (n == null) return '';
  if (n < 1024) return n+' B';
  if (n < 1048576) return (n/1024).toFixed(1)+' KB';
  if (n < 1073741824) return (n/1048576).toFixed(2)+' MB';
  return (n/1073741824).toFixed(2)+' GB';
}

function utf16be(buf) {
  const out = Buffer.from(buf);
  for (let i=0;i+1<out.length;i+=2){ const t=out[i]; out[i]=out[i+1]; out[i+1]=t; }
  return out.toString('utf16le').replace(/\0+$/,'');
}

function getFileType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (WORD_EXTENSIONS.includes(ext)) return 'word';
  if (SHEET_EXTENSIONS.includes(ext)) return 'sheet';
  if (SLIDE_EXTENSIONS.includes(ext)) return 'slide';
  if (PDF_EXTENSIONS.includes(ext)) return 'pdf';
  if (EBOOK_EXTENSIONS.includes(ext)) return 'ebook';
  if (IMAGE_EXTENSIONS.includes(ext)) return 'image';
  if (ARCHIVE_EXTENSIONS.includes(ext)) return 'archive';
  if (AUDIO_EXTENSIONS.includes(ext)) return 'audio';
  if (VIDEO_EXTENSIONS.includes(ext)) return 'video';
  if (FONT_EXTENSIONS.includes(ext)) return 'font';
  if (MARKDOWN_EXTENSIONS.includes(ext)) return 'markdown';
  if (TEXT_EXTENSIONS.includes(ext)) return 'text';
  return 'unknown';
}

async function zipXmlText(zip, filePattern, tagPattern) {
  const keys = Object.keys(zip.files).filter(f => filePattern.test(f));
  const parts = [];
  for (const k of keys) {
    const xml = await zip.files[k].async('string');
    const ms = xml.match(tagPattern);
    if (ms) parts.push(ms.map(m=>m.replace(/<[^>]+>/g,'')).filter(t=>t.trim()).join('\n'));
  }
  return parts;
}

async function parseOdt(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.files['content.xml'].async('string');

  let html = '';
  const blocks = xml.match(/<text:(h|p)[^>]*>[\s\S]*?<\/text:\1>/g) || [];
  for (const b of blocks) {
    const tag = b.match(/^<text:(h|p)/)[1];
    const outline = b.match(/text:outline-level="(\d+)"/);
    let txt = (b.match(/<text:[\w-]+[^>]*>([\s\S]*?)<\/text:[\w-]+>/g)||[])
      .map(s=>s.replace(/<[^>]+>/g,'')).join('');
    if (!txt) txt = b.replace(/<[^>]+>/g,'');
    txt = escapeHtml(txt);
    if (tag === 'h') html += '<h'+(outline?outline[1]:2)+' style="margin:.8em 0 .4em">'+txt+'</h'+(outline?outline[1]:2)+'>';
    else if (txt.trim()) html += '<p style="line-height:1.8;margin:.4em 0">'+txt+'</p>';

    if (/<text:list-item/.test(b)) html = html;
  }
  if (!html) html = '<p>'+escapeHtml(htmlToText(xml))+'</p>';
  return html;
}

async function parseOds(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.files['content.xml'].async('string');
  const sheets = {}, names = [];
  const tables = xml.match(/<table:table[\s\S]*?<\/table:table>/g) || [];
  tables.forEach((tb, idx) => {
    const nm = (tb.match(/table:name="([^"]*)"/)||[])[1] || ('Sheet'+(idx+1));
    const rows = [];
    const trs = tb.match(/<table:table-row[\s\S]*?<\/table:table-row>/g) || [];
    trs.forEach(tr => {
      const cells = [];
      const tds = tr.match(/<table:table-cell[\s\S]*?<\/table:table-cell>/g) || [];
      tds.forEach(td => {
        const repeat = +((td.match(/table:number-columns-repeated="(\d+)"/)||[])[1]||1);
        let v = (td.match(/<text:p[^>]*>([\s\S]*?)<\/text:p>/g)||[]).map(s=>s.replace(/<[^>]+>/g,'')).join(' ');
        if (v === '') v = (td.match(/office:value="([^"]*)"/)||[])[1]||'';
        for (let r=0;r<Math.min(repeat,50);r++) cells.push(v);
      });
      if (cells.some(c=>c!=='')) rows.push(cells.join('\t'));
    });
    names.push(nm); sheets[nm] = rows.join('\n');
  });
  return { sheets, sheetNames: names, content: sheets[names[0]]||'' };
}

async function parseOdp(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.files['content.xml'].async('string');
  const pages = xml.match(/<draw:page[\s\S]*?<\/draw:page>/g) || [];
  const slides = pages.map(p => {
    return (p.match(/<text:(span|p|list-item)[^>]*>([\s\S]*?)<\/text:\1>/g)||[])
      .map(s=>s.replace(/<[^>]+>/g,'')).filter(t=>t.trim()).join('\n');
  }).filter(s=>s.trim());
  return slides;
}

async function parseIwork(buf, kind) {
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files);
  const preview = names.find(n=>/preview\.pdf$/i.test(n)) || names.find(n=>/^.*\.pdf$/i.test(n));
  if (preview) {
    const pb = await zip.files[preview].async('nodebuffer');
    const data = await pdfParse(pb);
    return { viaPdf:true, text:data.text, numpages:data.numpages };
  }
  const idx = names.find(n=>/index\.xml$/i.test(n)) || names.find(n=>/\.xml$/i.test(n));
  if (idx) {
    const xml = await zip.files[idx].async('string');
    return { viaPdf:false, text: htmlToText(xml) };
  }
  return { viaPdf:false, text:'' };
}

function parseRtf(text) {
  if (!text) return '';
  let cp = 936;
  const cpm = text.match(/ansicpg(\d+)/i);
  if (cpm) cp = +cpm[1];
  const byteBuf = [];
  const flushBytes = () => {
    if (!byteBuf.length) return '';
    const b = Buffer.from(byteBuf); byteBuf.length = 0;
    let enc = 'gbk';
    if (cp === 65001) enc = 'utf-8';
    else if (cp === 1252) enc = 'windows-1252';
    else if (cp === 950) enc = 'big5';
    else if (cp === 932) enc = 'shift-jis';
    try { return new TextDecoder(enc).decode(b); } catch (e) { return b.toString('latin1'); }
  };
  let out = '', i = 0, skipGroupDepth = 0;
  const groupOpens = [];
  while (i < text.length) {
    const ch = text[i];
    if (ch === '\\') {
      const c1 = text[i+1];

      if (c1 === "'") {
        const v = parseInt(text.substr(i+2,2),16);
        if (!isNaN(v)) byteBuf.push(v);
        i += 4; continue;
      }

      out += flushBytes();
      if (c1 === '\\') { out += '\\'; i += 2; continue; }
      if (c1 === '{') { out += '{'; i += 2; continue; }
      if (c1 === '}') { out += '}'; i += 2; continue; }
      if (c1 === '~') { out += ' '; i += 2; continue; }
      if (c1 === '-') { i += 2; continue; }
      if (c1 === '_') { out += '-'; i += 2; continue; }

      const um = text.slice(i+1).match(/^u(-?\d+)\s?.?/);
      if (um) {
        let code = +um[1]; if (code < 0) code += 65536;
        out += String.fromCharCode(code);
        i += 1 + um[0].length; continue;
      }

      const wm = text.slice(i+1).match(/^([a-zA-Z]+)(-?\d+)?[ ]?/);
      if (wm) {
        const w = wm[1].toLowerCase();
        if (w === 'par' || w === 'line') out += '\n';
        else if (w === 'tab') out += '\t';
        else if (w === 'bullet') out += '•';
        else if (w === 'emdash') out += '—';
        else if (w === 'endash') out += '–';
        else if (w === 'lquote' || w === 'rquote') out += '’';
        else if (w === 'ldblquote' || w === 'rdblquote') out += '“';
        else if (w === 'lcbracket') out += '{';
        else if (w === 'rcbracket') out += '}';
        i += 1 + wm[0].length; continue;
      }
      i++; continue;
    }
    if (ch === '{') {
      out += flushBytes();
      const head = text.slice(i+1, i+140);
      const starM = head.match(/^\\\*\\?([a-zA-Z]+)/);
      const dirM = head.match(/^\\([a-zA-Z]+)/);
      const SKIP_DEST = ['fonttbl','colortbl','stylesheet','styles','info','pict','object','objdata','fldinst','fldchars','themefontdata','themedata','datastore','generator','rsidtbl','listtable','listoverridetable','latentstyles','colorschememapping','xmlnstbl','mtef','fontemb','nonshppict','shppict','background'];
      let skipGroup = false;
      if (starM) skipGroup = true;
      else if (dirM && SKIP_DEST.includes(dirM[1].toLowerCase())) skipGroup = true;
      if (/\\ch?ftnsep|\\ch?aftnsep/i.test(head)) skipGroup = true;
      if (skipGroup) {
        let depth = 1, j = i + 1;
        while (j < text.length && depth > 0) {
          if (text[j] === '\\') j += 2;
          else { if (text[j] === '{') depth++; else if (text[j] === '}') depth--; j++; }
        }
        i = j; continue;
      }
      i++; continue;
    }
    if (ch === '}') { out += flushBytes(); i++; continue; }
    if (ch === '\r') { i++; continue; }
    if (ch === '\n') { i++; continue; }
    const cc = ch.charCodeAt(0);
    if (cc < 128) {

      out += flushBytes();
      if (cc >= 32 || ch === '\t') out += ch;
      i++;
    } else {

      byteBuf.push(cc & 0xff);
      i++;
    }
  }
  out += flushBytes();
  return out.replace(/[ \t]+\n/g,'\n').replace(/\n{3,}/g,'\n\n').trim();
}

function mdToHtml(md) {
  const lines = String(md).replace(/\r\n/g,'\n').split('\n');
  let html = '', inCode = false, inUl = false, inOl = false, inQuote = false;
  const inline = s => escapeHtml(s)
    .replace(/`([^`]+)`/g,'<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>')
    .replace(/__([^_]+)__/g,'<strong>$1</strong>')
    .replace(/\*([^*]+)\*/g,'<em>$1</em>')
    .replace(/~~([^~]+)~~/g,'<del>$1</del>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g,'<a href="$2">$1</a>');
  const closeLists = () => { if(inUl){html+='</ul>';inUl=false;} if(inOl){html+='</ol>';inOl=false;} };
  for (let raw of lines) {
    const line = raw;
    if (/^```/.test(line)) { closeLists(); if(!inCode){html+='<pre style="background:#f5f7fa;padding:12px;border-radius:6px;overflow:auto;font-family:Consolas,monospace;font-size:13px">';inCode=true;} else {html+='</pre>';inCode=false;} continue; }
    if (inCode) { html += escapeHtml(line)+'\n'; continue; }
    let m;
    if (m = line.match(/^(#{1,6})\s+(.*)$/)) { closeLists(); const lv=m[1].length; html+='<h'+lv+' style="margin:.8em 0 .4em">'+inline(m[2])+'</h'+lv+'>'; continue; }
    if (/^\s*([-*_])\s*\1\s*\1[\s\S]*$/.test(line)) { closeLists(); html+='<hr>'; continue; }
    if (m = line.match(/^\s*>\s?(.*)$/)) { closeLists(); html+='<blockquote style="border-left:3px solid #ccc;margin:.5em 0;padding:.2em 1em;color:#555">'+inline(m[1])+'</blockquote>'; continue; }
    if (m = line.match(/^\s*[-*+]\s+(.*)$/)) { if(!inUl){closeLists();html+='<ul style="line-height:1.8">';inUl=true;} html+='<li>'+inline(m[1])+'</li>'; continue; }
    if (m = line.match(/^\s*\d+[.)]\s+(.*)$/)) { if(!inOl){closeLists();html+='<ol style="line-height:1.8">';inOl=true;} html+='<li>'+inline(m[1])+'</li>'; continue; }
    if (m = line.match(/^\s*\|(.*)\|\s*$/)) {
      closeLists();
      const cells = m[1].split('|').map(c=>c.trim());
      if (cells.every(c=>/^:?-{2,}:?$/.test(c))) continue;
      html += '<div style="display:flex">'+cells.map(c=>'<div style="flex:1;padding:4px 10px;border:1px solid #e3e6eb">'+inline(c)+'</div>').join('')+'</div>';
      continue;
    }
    if (!line.trim()) { closeLists(); html+=''; continue; }
    closeLists();
    html += '<p style="line-height:1.8;margin:.5em 0">'+inline(line)+'</p>';
  }
  closeLists(); if(inCode) html+='</pre>';
  return html;
}

async function parseArchive(buf, ext) {
  const rows = [];
  if (ext === '.tar') return listTar(buf);
  if (ext === '.gz' || ext === '.tgz') {
    try {
      const inner = zlib.gunzipSync(buf);
      if (ext === '.tgz') return listTar(inner);
      return '<p>GZIP 单文件压缩，解压后大小 '+humanSize(inner.length)+'</p>';
    } catch(e){ return null; }
  }
  if (ext === '.bz2' || ext === '.tbz2') {
    try { const inner = zlib.brotliDecompressSync ? zlib.brotliDecompressSync(buf) : null; return inner ? '<p>Brotli/BZ2 解压后 '+humanSize(inner.length)+'</p>' : null; } catch(e){ return null; }
  }
  if (ext === '.xz') return null;

  if (['.7z','.rar','.cab','.iso','.dmg','.lzma','.z'].includes(ext)) return null;
  try {
    const zip = await JSZip.loadAsync(buf);
    const entries = Object.values(zip.files);
    let total = 0, dirs = 0;
    entries.forEach(f=>{
      if (f.dir) { dirs++; return; }
      total += f._data ? f._data.uncompressedSize : 0;
    });
    rows.push('<p style="color:#666">压缩包内共 '+entries.length+' 项（文件夹 '+dirs+' 个），解压总大小约 '+humanSize(total)+'</p>');
    rows.push('<div style="font-family:Consolas,monospace;font-size:13px;line-height:1.7">');
    entries.slice(0, 2000).forEach(f=>{
      const sz = f.dir ? '—' : humanSize(f._data?f._data.uncompressedSize:0);
      const icon = f.dir ? '📁' : '📄';
      rows.push('<div style="display:flex;gap:12px;padding:1px 0"><span style="width:90px;color:#888;text-align:right">'+sz+'</span><span>'+icon+' '+escapeHtml(f.name)+'</span></div>');
    });
    if (entries.length > 2000) rows.push('<p>… 其余 '+(entries.length-2000)+' 项省略</p>');
    rows.push('</div>');
    return rows.join('');
  } catch(e){ return null; }
}
function listTar(buf) {
  const rows = ['<div style="font-family:Consolas,monospace;font-size:13px;line-height:1.7">'];
  let off = 0, count = 0, total = 0;
  while (off + 512 <= buf.length) {
    const header = buf.slice(off, off+512);
    if (header.every(b=>b===0)) break;
    const name = header.slice(0,100).toString('utf8').replace(/\0+$/,'').trim();
    let size = parseInt(header.slice(124,136).toString('utf8').replace(/\0/g,'').trim()||'0',8)||0;
    const type = String.fromCharCode(header[156]);
    if (name) {
      count++; if(type!=='5') total += size;
      rows.push('<div><span style="width:110px;display:inline-block;color:#888;text-align:right">'+(type==='5'?'—':humanSize(size))+'</span> '+(type==='5'?'📁':'📄')+' '+escapeHtml(name)+'</div>');
    }
    off += 512 + Math.ceil(size/512)*512;
  }
  rows.unshift('<p style="color:#666">TAR 包内共 '+count+' 项，总大小约 '+humanSize(total)+'</p>');
  rows.push('</div>');
  return rows.join('');
}

function parseEml(raw) {
  const text = decodeTextBuffer(raw);
  const headEnd = text.search(/\r?\n\r?\n/);
  const head = headEnd >= 0 ? text.slice(0, headEnd) : text;
  let body = headEnd >= 0 ? text.slice(headEnd).replace(/^\r?\n\r?\n/,'') : '';
  const getH = k => { const m = head.match(new RegExp('^'+k+':\\s*(.*)$','im')); return m?decodeMimeWords(m[1].replace(/\r?\n\s+/g,' ')):''; };

  const bm = head.match(/boundary="?([^";\r\n]+)"?/i);
  let plain = '', htmlBody = '';
  if (bm) {
    const parts = body.split('--'+bm[1]);
    for (const p of parts) {
      const pm = p.match(/Content-Type:\s*([^;\r\n]+)/i);
      const cte = (p.match(/Content-Transfer-Encoding:\s*(\S+)/i)||[])[1]||'';
      const pBody = p.replace(/^[\s\S]*?\r?\n\r?\n/,'');
      const decoded = cte.toLowerCase()==='base64' ? safeB64ToText(pBody) : cte.toLowerCase()==='quoted-printable' ? decodeQP(pBody) : pBody;
      if (pm && /text\/html/i.test(pm[1])) htmlBody = decoded;
      else if (pm && /text\/plain/i.test(pm[1])) plain = decoded;
    }
  } else {
    const cte = (head.match(/Content-Transfer-Encoding:\s*(\S+)/i)||[])[1]||'';
    body = cte.toLowerCase()==='base64'?safeB64ToText(body):cte.toLowerCase()==='quoted-printable'?decodeQP(body):body;
    if (/text\/html/i.test(head)) htmlBody = body; else plain = body;
  }
  const attaches = [];
  const am = text.match(/Content-Disposition:[^;]*;\s*filename="?([^"\r\n;]+)"?/gi)||[];
  am.forEach(a=>{const m=a.match(/filename="?([^";\r\n]+)"?/i);if(m)attaches.push(decodeMimeWords(m[1]));});
  let h = '<div style="border:1px solid #e3e6eb;border-radius:8px;padding:14px 18px;margin-bottom:14px;background:#fafbfc;line-height:1.9">';
  h += '<div><b>主题：</b>'+escapeHtml(getH('Subject')||'(无主题)')+'</div>';
  h += '<div><b>发件人：</b>'+escapeHtml(getH('From'))+'</div>';
  h += '<div><b>收件人：</b>'+escapeHtml(getH('To'))+'</div>';
  if(getH('Cc')) h += '<div><b>抄送：</b>'+escapeHtml(getH('Cc'))+'</div>';
  h += '<div><b>时间：</b>'+escapeHtml(getH('Date'))+'</div>';
  if(attaches.length) h += '<div><b>附件：</b>'+attaches.map(escapeHtml).join('、')+'</div>';
  h += '</div>';
  if (htmlBody) {
    h += htmlBody.replace(/<script[\s\S]*?<\/script>/gi,'');
  } else {
    h += '<pre style="white-space:pre-wrap;line-height:1.8;font-family:inherit">'+escapeHtml(decodeMimeWords(plain))+'</pre>';
  }
  return h;
}
function decodeMimeWords(s) {
  return String(s).replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (_,cs,enc,data)=>{
    try {
      let buf;
      if (enc.toLowerCase()==='b') buf = Buffer.from(data,'base64');
      else buf = Buffer.from(data.replace(/_/g,' ').replace(/=([0-9A-Fa-f]{2})/g,(_,h)=>String.fromCharCode(parseInt(h,16))),'binary');
      return new TextDecoder(cs.toLowerCase().replace('_','-')).decode(buf);
    } catch(e){ return data; }
  });
}
function decodeQP(s){ return s.replace(/=\r?\n/g,'').replace(/=([0-9A-Fa-f]{2})/g,(_,h)=>String.fromCharCode(parseInt(h,16))); }
function safeB64ToText(s){ try { return decodeTextBuffer(Buffer.from(String(s).replace(/\s/g,''),'base64')); } catch(e){ return ''; } }

function parseFont(buf, ext) {
  const info = { family:'未知', subfamily:'', tables:[] };
  try {
    if (ext === '.woff2') {
      info.family = '（WOFF2 网页字体，压缩格式）';
    } else if (ext === '.woff') {

      const numTables = buf.readUInt16BE(12);
      info.tables.push('表数量:'+numTables);
    } else {

      let base = 0;
      const tag0 = buf.slice(0,4).toString('latin1');
      if (tag0 === 'ttcf') base = buf.readUInt32BE(8);
      const numTables = buf.readUInt16BE(base+4);
      const tableDir = [];
      for (let t=0;t<numTables;t++) {
        const o = base+12+t*16;
        tableDir.push({ tag:buf.slice(o,o+4).toString('latin1'), off:buf.readUInt32BE(o+8), len:buf.readUInt32BE(o+12) });
      }
      const nameT = tableDir.find(x=>x.tag==='name');
      if (nameT) {
        const no = nameT.off;
        const count = buf.readUInt16BE(no+2), strOff = no+buf.readUInt16BE(no+4);
        const names = {};
        for (let r=0;r<count;r++) {
          const ro = no+6+r*12;
          const pid = buf.readUInt16BE(ro+2), nameId = buf.readUInt16BE(ro+6), len=buf.readUInt16BE(ro+8), off2=buf.readUInt16BE(ro+10);
          if (names[nameId]) continue;
          const raw = buf.slice(strOff+off2, strOff+off2+len);

          const decodeName = (pid, raw) => (pid===0||pid===3) ? utf16be(raw) : raw.toString('latin1');
          names[nameId] = decodeName(pid, raw);
        }
        info.family = names[1]||info.family; info.subfamily = names[2]||'';
        info.full = names[4]||''; info.version = names[5]||''; info.designer = names[9]||''; info.vendor = names[8]||'';
      }
    }
  } catch(e){ info.err = e.message; }
  let h = '<div style="line-height:2">';
  h += '<h3 style="margin:.2em 0 .6em">'+escapeHtml(info.family)+(info.subfamily?' '+escapeHtml(info.subfamily):'')+'</h3>';
  if(info.full) h+='<div><b>完整名称：</b>'+escapeHtml(info.full)+'</div>';
  if(info.version) h+='<div><b>版本：</b>'+escapeHtml(info.version)+'</div>';
  if(info.designer) h+='<div><b>设计师：</b>'+escapeHtml(info.designer)+'</div>';
  if(info.vendor) h+='<div><b>厂商：</b>'+escapeHtml(info.vendor)+'</div>';
  h += '<div><b>文件大小：</b>'+humanSize(buf.length)+'</div>';

  h += '<div style="margin-top:16px;padding:18px;border:1px solid #e3e6eb;border-radius:8px"><div style="font-size:40px">永 OfficePro 字体预览 ABCDEFG abcdefg 0123456789</div><div style="font-size:20px;margin-top:10px">中文：永和九年岁在癸丑暮春之初，会于会稽山阴之兰亭。</div></div>';
  h += '</div>';
  return h;
}

function readOleStreams(buf) {
  if (buf.length < 512 || buf.readUInt32LE(0) !== 0xE011CFD0) return {};
  try {
    const sectorShift = buf.readUInt16LE(30), secSize = 1<<sectorShift, miniSize = 64;
    const firstDir = buf.readUInt32LE(48), miniCutoff = buf.readUInt32LE(56);
    const firstMiniFAT = buf.readUInt32LE(60), numMiniFAT = buf.readUInt32LE(64);
    const firstDIFAT = buf.readUInt32LE(68), numDIFAT = buf.readUInt32LE(72);
    const secOff = s => (s+1)*secSize;
    const difat = [];
    for (let i=0;i<109;i++) difat.push(buf.readUInt32LE(76+i*4));
    let d = firstDIFAT;
    for (let k=0;k<numDIFAT && d!==0xFFFFFFFF && d!==0xFFFFFFFE;k++){
      const o=secOff(d), cnt=secSize/4-1;
      for (let i=0;i<cnt;i++) difat.push(buf.readUInt32LE(o+i*4));
      d=buf.readUInt32LE(o+cnt*4);
    }
    const FAT=[];
    for (const fsSec of difat){ if(fsSec===0xFFFFFFFF||fsSec===0xFFFFFFFE) continue; const o=secOff(fsSec); for(let i=0;i<secSize/4;i++) FAT.push(buf.readUInt32LE(o+i*4)); }
    const chain=(start,arr)=>{const out=[];let s=start,g=0;while(s!==0xFFFFFFFF&&s!==0xFFFFFFFE&&s<arr.length&&g++<200000){out.push(s);s=arr[s];}return out;};
    const byChain=(start,size,arr)=>Buffer.concat(chain(start,arr).map(s=>buf.slice(secOff(s),secOff(s)+secSize))).slice(0,size);
    const dirBuf=byChain(firstDir,FAT.length*secSize,FAT);
    const entries=[];
    for(let off=0;off+128<=dirBuf.length;off+=128){
      const nameLen=dirBuf.readUInt16LE(off+64);
      const name=nameLen>0?dirBuf.toString('utf16le',off,off+Math.max(0,nameLen-2)):'';
      entries.push({name,type:dirBuf[off+66],start:dirBuf.readUInt32LE(off+116),size:dirBuf.readUInt32LE(off+120)});
    }
    const root=entries.find(e=>e.type===5);
    let miniStream=Buffer.alloc(0),miniFAT=[];
    if(root) miniStream=byChain(root.start,root.size,FAT);
    if(numMiniFAT>0){const mb=byChain(firstMiniFAT,numMiniFAT*secSize,FAT);for(let i=0;i<mb.length/4;i++)miniFAT.push(mb.readUInt32LE(i*4));}
    const streams={};
    for(const e of entries){
      if(e.type!==2) continue;
      let data;
      if(e.size<miniCutoff) data=Buffer.concat(chain(e.start,miniFAT).map(s=>miniStream.slice(s*miniSize,s*miniSize+miniSize))).slice(0,e.size);
      else data=byChain(e.start,e.size,FAT);
      streams[e.name]=data;
    }
    return streams;
  } catch(e){ return {}; }
}

function extractPptOle(buf) {
  try {
    const streams=readOleStreams(buf);
    const key=Object.keys(streams).find(n=>/PowerPoint Document/i.test(n));
    if(!key) throw new Error('no ppt stream');
    const pb=streams[key], slides=[]; let cur=null;
    const clean=t=>t.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/\r/g,'\n').replace(/\n{2,}/g,'\n').trim();
    const walk=(start,end)=>{
      let p=start;
      while(p+8<=end){
        const verInst=pb.readUInt16LE(p), type=pb.readUInt16LE(p+2), len=pb.readUInt32LE(p+4);
        const ver=verInst&0xF, body=p+8;
        if(body+len>end) break;
        if(type===1006){ cur=[]; slides.push(cur); walk(body,body+len); }
        else if(ver===0xF) walk(body,body+len);
        else if(type===4000){ const t=clean(pb.toString('utf16le',body,body+len)); if(t){cur=cur||[];cur.push(t);} }
        else if(type===4008){ let t=''; try{ t=new TextDecoder('gbk').decode(pb.slice(body,body+len)); }catch(e){ t=pb.toString('latin1',body,body+len); } t=clean(t); if(t){cur=cur||[];cur.push(t);} }
        p=body+len;
      }
    };
    walk(0,pb.length);
    const result=slides.map(a=>clean(a.join('\n'))).filter(Boolean);
    if(result.length) return result;
  } catch(e){}

  const txt=extractOleText(buf);
  return txt?txt.split(/\n{2,}/).map(s=>s.trim()).filter(Boolean):[];
}

async function parseFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const fileName = path.basename(filePath);
  let fileSize = 0;
  try { fileSize = fs.statSync(filePath).size; } catch(e){}
  const type = getFileType(filePath);
  const buf = () => fs.readFileSync(filePath);

  try {

    if (ext === '.pdf') {
      const data = await pdfParse(buf());
      return { type:'pdf', content:data.text, fileName, fileSize, numpages:data.numpages };
    }
    if (ext === '.xps' || ext === '.oxps') {

      try {
        const zip = await JSZip.loadAsync(buf());
        const pageKeys = Object.keys(zip.files).filter(f=>/\/documents\/1\/pages\/\d+\.fpage$/.test(f)).sort();
        const texts = [];
        for (const k of pageKeys) {
          const x = await zip.files[k].async('string');
          const t = (x.match(/<Glyphs[^>]*UnicodeString="([^"]*)"/g)||[]).map(m=>(m.match(/UnicodeString="([^"]*)"/)||[])[1]||'').join('');
          texts.push(t);
        }
        return wordHtml(pre(texts.join('\n\n——————\n\n'), false), fileName, fileSize, {warning:'xps'});
      } catch(e){ return awaitFallback(filePath, ext, fileName, fileSize); }
    }

    if (['.docx','.docm','.dotx','.dotm'].includes(ext)) {
      let mh = '';
      try { const result = await mammoth.convertToHtml({ path:filePath }); mh = result.value || ''; } catch(e) { mh = ''; }
      if (mh.replace(/<[^>]+>/g,'').trim()) return wordHtml(mh, fileName, fileSize);

      try {
        const zip = await JSZip.loadAsync(buf());
        if (zip.files['word/document.xml']) {
          const xml = await zip.files['word/document.xml'].async('string');
          const paras = xml.split(/<\/w:p>/).map(p=>(p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g)||[]).map(m=>m.replace(/<\/?w:t[^>]*>/g,'')).join('')).filter(x=>x.trim());
          const text = paras.join('\n');
          if (text.trim()) return wordHtml(plainTextToHtml(text), fileName, fileSize, { warning:'docx-fallback', note:'文档结构特殊，已提取纯文本' });
        }
      } catch(e2) {}

      const head = buf().slice(0,8);
      const isOle = head.length>=8 && head.readUInt32LE(0)===0xE011CFD0;
      const tip = isOle
        ? '该 Word 文档已加密或受密码保护。请先在 Word/WPS 中输入密码打开，再另存为未加密文档后用本软件打开。'
        : '该 Word 文档已损坏或结构无法识别，无法提取内容。请用 Word/WPS 打开检查，或另存为新文档后再试。';
      return wordHtml('<p style="color:#a15c00;line-height:1.9">'+escapeHtml(tip)+'</p>', fileName, fileSize, { warning:'docx-unreadable' });
    }
    if (ext === '.odt' || ext === '.ott') {
      const html = await parseOdt(buf());
      return wordHtml(html, fileName, fileSize);
    }
    if (ext === '.pages') {
      const r = await parseIwork(buf(),'pages');
      if (r.viaPdf) return { type:'pdf', content:r.text, fileName, fileSize, numpages:r.numpages };
      return wordHtml(pre(r.text), fileName, fileSize);
    }
    if (['.doc','.wps','.wpt','.dot'].includes(ext)) {

      if (WordExtractor) {
        try {
          const ex = new WordExtractor();
          const d = await ex.extract(filePath);
          let body = (d.getBody ? d.getBody() : '') || '';

          body = repairCp1252Gbk(body);
          if (body.trim()) {
            const extra = [];
            try { const fn = d.getFootnotes && d.getFootnotes(); if (fn && fn.trim()) extra.push('【脚注】\n' + repairCp1252Gbk(fn)); } catch(_) {}
            try { const en = d.getEndnotes && d.getEndnotes(); if (en && en.trim()) extra.push('【尾注】\n' + repairCp1252Gbk(en)); } catch(_) {}
            const full = extra.length ? body + '\n\n' + extra.join('\n\n') : body;
            return wordHtml(plainTextToHtml(full), fileName, fileSize);
          }
        } catch(we) { }
      }

      try {
        const result = await mammoth.convertToHtml({ path:filePath });
        if (result.value && result.value.replace(/<[^>]+>/g,'').trim()) return wordHtml(result.value, fileName, fileSize);
      } catch(e) {}

      const b = buf();
      let text = extractOleText(b);
      if (!text.trim()) text = decodeTextBuffer(b);
      return wordHtml(pre(text), fileName, fileSize, { warning:'legacy-ole' });
    }

    if (['.xlsx','.xls','.xlsm','.xlsb','.xlt','.xltx','.xltm','.ods','.ots','.et','.ett','.numbers','.dif','.slk'].includes(ext)) {

      if (ext==='.ods'||ext==='.ots') {
        try {
          const o = await parseOds(buf());
          if (o.sheetNames.length) return { type:'sheet', content:o.content, sheets:o.sheets, sheetNames:o.sheetNames, fileName, fileSize };
        } catch(e){}
      }
      if (ext === '.numbers') {
        const r = await parseIwork(buf(),'numbers');
        if (r.viaPdf) return { type:'pdf', content:r.text, fileName, fileSize, numpages:r.numpages };
      }
      let workbook;
      try {
        workbook = XLSX.readFile(filePath, { codepage:936, cellDates:true });
      } catch(xe) {
        const head = buf().slice(0,8);
        const isOle = head.length>=8 && head.readUInt32LE(0)===0xE011CFD0;
        const tip = isOle ? '该表格已加密或受密码保护，请先在 Excel/WPS 中输入密码打开并另存为未加密文件。' : '该表格文件已损坏或结构无法识别，请用 Excel/WPS 打开检查或另存为新文件。';
        return { type:'sheet', content:tip, sheets:{ 'Sheet1':tip }, sheetNames:['Sheet1'], fileName, fileSize, warning:'xlsx-unreadable' };
      }
      const sheets = {}, names = workbook.SheetNames;
      names.forEach(n=>{ sheets[n] = XLSX.utils.sheet_to_csv(workbook.Sheets[n], { FS:'\t' }); });
      return { type:'sheet', content:sheets[names[0]]||'', sheets, sheetNames:names, fileName, fileSize };
    }
    if (['.csv','.tsv','.prn'].includes(ext)) {
      const text = decodeTextBuffer(buf());
      return { type:'sheet', content:text, sheets:{ 'Sheet1':text }, sheetNames:['Sheet1'], fileName, fileSize };
    }

    if (['.pptx','.ppsx','.pptm','.potx','.potm','.dps'].includes(ext)) {
      try {
        const zip = await JSZip.loadAsync(buf());
        const slideFiles = Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f))
          .sort((a,b)=>parseInt(a.match(/\d+/)[0])-parseInt(b.match(/\d+/)[0]));
        const slides = [];
        for (const sf of slideFiles) {
          const xml = await zip.files[sf].async('string');
          const ms = xml.match(/<a:t>([^<]*)<\/a:t>/g);
          slides.push(ms?ms.map(m=>m.replace(/<\/?a:t>/g,'')).filter(t=>t.trim()).join('\n'):'');
        }
        return { type:'slide', content:JSON.stringify(slides), slides, fileName, fileSize };
      } catch(pe) {
        const head = buf().slice(0,8);
        const isOle = head.length>=8 && head.readUInt32LE(0)===0xE011CFD0;
        const tip = isOle ? '该演示文稿已加密或受密码保护，请先在 PowerPoint/WPS 中解密后另存。' : '该演示文稿已损坏或结构无法识别，请用 PowerPoint/WPS 打开检查或另存为新文件。';
        return { type:'slide', content:JSON.stringify([tip]), slides:[tip], fileName, fileSize, warning:'pptx-unreadable' };
      }
    }
    if (ext === '.odp' || ext === '.otp') {
      const slides = await parseOdp(buf());
      return { type:'slide', content:JSON.stringify(slides), slides, fileName, fileSize };
    }
    if (ext === '.key') {
      const r = await parseIwork(buf(),'key');
      if (r.viaPdf) return { type:'pdf', content:r.text, fileName, fileSize, numpages:r.numpages };
      return { type:'slide', content:'[]', slides:r.text?r.text.split(/\n{2,}/):[], fileName, fileSize };
    }
    if (['.ppt','.pps','.pot','.dpt','.dps'].includes(ext)) {
      const slides = extractPptOle(buf());
      return { type:'slide', content:JSON.stringify(slides), slides, fileName, fileSize, warning:slides.length?'legacy-ole':'legacy-empty' };
    }

    if (ext === '.epub') {
      const zip = await JSZip.loadAsync(buf());
      let opfPath = 'OEBPS/content.opf';
      try { const c = await zip.files['META-INF/container.xml'].async('string'); const m=c.match(/full-path="([^"]+)"/); if(m) opfPath=m[1]; } catch(e){}
      let html = '';
      try {
        const opf = await zip.files[opfPath].async('string');
        const manifest = {};
        (opf.match(/<item[^>]*>/g)||[]).forEach(item=>{
          const id=(item.match(/id="([^"]*)"/)||[])[1], href=(item.match(/href="([^"]*)"/)||[])[1];
          if(id&&href) manifest[id]=href;
        });
        const base = path.posix.dirname(opfPath.replace(/\\/g,'/'));
        const spines = opf.match(/<itemref[^>]*idref="([^"]*)"[^>]*>/g)||[];
        for (const sm of spines) {
          const idref = sm.match(/idref="([^"]*)"/)[1];
          let href = manifest[idref]; if(!href) continue;
          let fp = base?base+'/'+href:href; fp=fp.replace(/\/+/g,'/');
          try {
            const x = await zip.files[fp].async('string');
            const bm = x.match(/<body[^>]*>([\s\S]*)<\/body>/i);
            html += bm?bm[1]:x;
          } catch(e){}
        }
      } catch(e){}
      return wordHtml(html||'<p>(电子书内容为空或为DRM加密)</p>', fileName, fileSize);
    }
    if (ext === '.fb2') {
      const xml = decodeTextBuffer(buf());
      const secs = xml.match(/<(section|p)[^>]*>[\s\S]*?<\/\1>/g)||[];
      const txt = secs.map(s=>htmlToText(s)).filter(Boolean).join('\n\n');
      return wordHtml(pre(txt||htmlToText(xml)), fileName, fileSize);
    }
    if (['.mobi','.azw','.azw3','.azw4','.lit','.pdb'].includes(ext)) {
      const b = buf();
      const strings = extractStrings(b, 6).filter(s=>/[\u4e00-\u9fa5]|[a-zA-Z]{4,}/.test(s));
      const text = strings.join('\n');
      return wordHtml(pre(text, false), fileName, fileSize, { warning:'ebook-raw', note:'该电子书为专有二进制格式，以下为尽力提取的文本，建议安装对应阅读器获得最佳效果' });
    }

    if (IMAGE_EXTENSIONS.includes(ext)) {
      const b = buf();
      if (['.psd','.ai','.wmf','.emf','.sketch'].includes(ext)) {

        const note = ({'.psd':'Photoshop 文档','.ai':'Illustrator 矢量文档','.wmf':'Windows 图元文件','.emf':'增强图元文件','.sketch':'Sketch 设计稿'})[ext];
        return wordHtml('<p>这是 '+note+'（'+humanSize(b.length)+'），属于专业设计格式，建议用对应设计软件打开。</p>', fileName, fileSize, {warning:'design-format'});
      }
      const mime = MIME[ext] || 'image/jpeg';
      return { type:'image', content:'data:'+mime+';base64,'+b.toString('base64'), fileName, fileSize };
    }

    if (ARCHIVE_EXTENSIONS.includes(ext)) {
      const html = await parseArchive(buf(), ext);
      if (html) return wordHtml('<h3 style="margin:.2em 0 .6em">📦 '+escapeHtml(fileName)+' 压缩包内容</h3>'+html, fileName, fileSize);

      const b = buf();
      const sig = b.slice(0,8).toString('hex');
      return wordHtml('<h3>📦 '+escapeHtml(fileName)+'</h3><p>这是 '+ext.toUpperCase()+' 压缩格式（'+humanSize(b.length)+'，文件头 '+sig+'）。</p><p>该压缩算法需专用解压软件，已为你显示文件基本信息。</p>', fileName, fileSize, {warning:'archive-unsupported'});
    }

    if (AUDIO_EXTENSIONS.includes(ext) || VIDEO_EXTENSIONS.includes(ext)) {
      const isAudio = AUDIO_EXTENSIONS.includes(ext);
      const url = 'file:///'+filePath.replace(/\\/g,'/').split('/').map(encodeURIComponent).join('/');
      return { type:'media', mediaType:isAudio?'audio':'video', mime:MIME[ext]||(isAudio?'audio/mpeg':'video/mp4'), url, fileName, fileSize };
    }

    if (FONT_EXTENSIONS.includes(ext)) {
      return wordHtml(parseFont(buf(),ext), fileName, fileSize);
    }

    if (ext === '.eml') return wordHtml(parseEml(buf()), fileName, fileSize);
    if (ext === '.msg') {
      const b = buf(); const t = extractOleText(b);
      return wordHtml(pre(t||'Outlook 邮件（.msg），建议用 Outlook 打开'), fileName, fileSize, {warning:'msg'});
    }

    if (ext === '.rtf') {
      const raw = decodeTextBuffer(buf());
      return wordHtml(pre(parseRtf(raw)), fileName, fileSize);
    }

    if (MARKDOWN_EXTENSIONS.includes(ext)) {
      return wordHtml(mdToHtml(decodeTextBuffer(buf())), fileName, fileSize);
    }

    if (['.html','.htm','.xhtml','.shtml'].includes(ext)) {
      const text = decodeTextBuffer(buf());
      const bm = text.match(/<body[^>]*>([\s\S]*)<\/body>/i);
      let body = bm?bm[1]:text;
      body = body.replace(/<script[\s\S]*?<\/script>/gi,'').replace(/<style[\s\S]*?<\/style>/gi,'');
      return wordHtml(body, fileName, fileSize);
    }
    if (ext === '.mht' || ext === '.mhtml') {
      const raw = decodeTextBuffer(buf());
      const hm = raw.match(/Content-Type:\s*text\/html[\s\S]*?\r?\n\r?\n([\s\S]*?)(?:\r?\n--|$)/i);
      let body = hm?hm[1]:raw;
      const cte=(raw.match(/Content-Transfer-Encoding:\s*(\S+)/i)||[])[1]||'';
      if(/base64/i.test(cte)) body = safeB64ToText(body);
      else if(/quoted-printable/i.test(cte)) body = decodeQP(body);
      body = body.replace(/<script[\s\S]*?<\/script>/gi,'').replace(/<style[\s\S]*?<\/style>/gi,'');
      return wordHtml(body, fileName, fileSize);
    }

    if (TEXT_EXTENSIONS.includes(ext)) {
      const b = buf();
      let text = decodeTextBuffer(b);

      if (ext === '.json' || ext === '.jsonl' || ext === '.ndjson' || ext === '.geojson' || ext === '.json5') {
        try {
          if (ext === '.jsonl' || ext === '.ndjson') {
            text = text.split(/\r?\n/).filter(Boolean).map(l=>JSON.stringify(JSON.parse(l),null,2)).join('\n,\n');
          } else text = JSON.stringify(JSON.parse(text),null,2);
        } catch(e){}
        return wordHtml(pre(text,true), fileName, fileSize);
      }

      if (ext === '.ipynb') {
        try {
          const nb = JSON.parse(text); const cells=[];
          (nb.cells||[]).forEach(c=>{
            const src = Array.isArray(c.source)?c.source.join(''):'';
            cells.push((c.cell_type==='markdown'?'# [Markdown] ':'# ['+(c.execution_count!=null?'输出':'代码')+'] ')+src);
          });
          return wordHtml(pre(cells.join('\n\n'),true), fileName, fileSize);
        } catch(e){ return wordHtml(pre(text,true), fileName, fileSize); }
      }
      const mono = !['.txt','.text','.log','.org','.rst','.adoc','.srt','.ass','.vtt','.lrc','.ics','.vcf'].includes(ext);
      return wordHtml(pre(text, mono), fileName, fileSize);
    }

    return awaitFallback(filePath, ext, fileName, fileSize);

  } catch(e) {

    try { return awaitFallback(filePath, ext, fileName, fileSize, e.message); }
    catch(e2) {
      return { type:'word', content:'<p>文件读取失败：'+escapeHtml(e.message)+'</p>', fileName, fileSize, error:e.message };
    }
  }
}

async function awaitFallback(filePath, ext, fileName, fileSize, errMsg) {
  const b = fs.readFileSync(filePath);
  if (isProbablyText(b)) {
    const text = decodeTextBuffer(b);
    return wordHtml(pre(text, true), fileName, fileSize, { warning:'unknown-as-text', note:'未识别的扩展名，已按文本方式打开' });
  }
  const strings = extractStrings(b, 6).slice(0,120);
  let h = '<div style="line-height:1.8">';
  h += '<h3 style="margin:.2em 0 .6em">二进制文件预览：'+escapeHtml(fileName)+'</h3>';
  h += '<table style="border-collapse:collapse;margin-bottom:12px;font-size:13px">';
  h += '<tr><td style="padding:2px 14px 2px 0;color:#888">类型</td><td>'+escapeHtml(ext||'无扩展名')+' 二进制文件</td></tr>';
  h += '<tr><td style="padding:2px 14px 2px 0;color:#888">大小</td><td>'+humanSize(b.length)+'</td></tr>';
  h += '<tr><td style="padding:2px 14px 2px 0;color:#888">文件头</td><td style="font-family:monospace">'+b.slice(0,16).toString('hex')+'</td></tr>';
  if(errMsg) h+='<tr><td style="padding:2px 14px 2px 0;color:#888">说明</td><td>'+escapeHtml(errMsg)+'</td></tr>';
  h += '</table>';
  h += '<p style="color:#a15c00;background:#fff7e6;border:1px solid #ffe1a8;padding:8px 12px;border-radius:6px">这是一个二进制文件，无法直接预览文字内容。请在文件上右键选择“打开方式”，用对应的专用软件（如压缩软件、专业查看器等）打开。</p>';

  if (strings.length) {
    h += '<h4 style="margin:14px 0 6px">其中可识别的字符串</h4>';
    h += '<div style="font-family:Consolas,monospace;font-size:12px;color:#445;max-height:240px;overflow:auto">'+strings.map(s=>'<div>'+escapeHtml(s)+'</div>').join('')+'</div>';
  }
  h += '</div>';
  return wordHtml(h, fileName, fileSize, { warning:'binary-hex' });
}

function extractOleText(buf) {
  const chunks = [];
  let cur = [];
  for (let i = 0; i + 1 < buf.length; i += 2) {
    const lo = buf[i], hi = buf[i+1];
    const ok = (hi === 0 && ((lo >= 0x20 && lo <= 0x7e) || lo === 0x0a || lo === 0x0d || lo === 0x09))
            || (hi >= 0x4e && hi <= 0x9f);
    if (ok) cur.push(buf[i], buf[i+1]);
    else { if (cur.length >= 8) chunks.push(Buffer.from(cur).toString('utf16le')); cur = []; }
  }
  if (cur.length >= 8) chunks.push(Buffer.from(cur).toString('utf16le'));
  let text = chunks.join('\n');
  text = text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/[ \t]{3,}/g,' ').replace(/\n{3,}/g,'\n\n').trim();
  return text;
}

module.exports = { parseFile, getFileType, decodeTextBuffer, extractOleText, escapeHtml, isProbablyText, humanSize, repairCp1252Gbk, plainTextToHtml };
