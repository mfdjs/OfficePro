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
  if (!fs.existsSync(p)) {
    console.log('WAIT: missing ' + name);
    process.exit(0);
  }
  parts.push(fs.readFileSync(p, 'utf8').replace(/\s+/g, ''));
}

const b64 = parts.join('');
if (b64.length !== meta.gzB64Len) {
  console.log('WAIT: b64 length ' + b64.length + '/' + meta.gzB64Len);
  process.exit(0);
}

const gz = Buffer.from(b64, 'base64');
const hash = crypto.createHash('sha256').update(gz).digest('hex');
if (hash !== meta.sha256gz) {
  console.log('WAIT: sha256 mismatch ' + hash);
  process.exit(0);
}

const raw = zlib.gunzipSync(gz);
if (raw.length !== meta.rawBytes) {
  console.log('WAIT: raw length mismatch ' + raw.length + '/' + meta.rawBytes);
  process.exit(0);
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
  console.log('WAIT: file count ' + count + '/' + meta.fileCount);
  process.exit(0);
}

console.log('OK: unpacked ' + count + ' files');
