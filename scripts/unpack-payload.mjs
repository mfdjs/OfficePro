import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import crypto from 'node:crypto';

const root = process.cwd();
const payloadDir = path.join(root, '.payload');
const metaPath = path.join(payloadDir, 'meta.json');

if (!fs.existsSync(metaPath)) {
  console.log('SKIP: .payload/meta.json not found, nothing to unpack');
  process.exit(0);
}

const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
const parts = [];
for (let i = 0; i < meta.parts; i++) {
  const name = 'part.' + String(i).padStart(3, '0');
  const p = path.join(payloadDir, name);
  if (!fs.existsSync(p)) throw new Error('missing part: ' + name);
  parts.push(fs.readFileSync(p, 'utf8').replace(/\s+/g, ''));
}

const b64 = parts.join('');
if (b64.length !== meta.gzB64Len) {
  throw new Error('b64 length mismatch: got ' + b64.length + ', expected ' + meta.gzB64Len);
}

const gz = Buffer.from(b64, 'base64');
const hash = crypto.createHash('sha256').update(gz).digest('hex');
if (hash !== meta.sha256gz) {
  throw new Error('sha256 mismatch: got ' + hash + ', expected ' + meta.sha256gz);
}

const raw = zlib.gunzipSync(gz);
if (raw.length !== meta.rawBytes) {
  throw new Error('raw length mismatch: got ' + raw.length + ', expected ' + meta.rawBytes);
}

const manifest = JSON.parse(raw.toString('utf8'));
let count = 0;
for (const f of manifest.files) {
  const dest = path.join(root, f.p);
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.writeFileSync(dest, Buffer.from(f.b64, 'base64'));
  count++;
}

if (count !== meta.fileCount) {
  throw new Error('file count mismatch: got ' + count + ', expected ' + meta.fileCount);
}

console.log('OK: unpacked ' + count + ' files');
