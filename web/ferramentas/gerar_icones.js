'use strict';
// Gera os ícones do app instalável (PWA) em public/icones/, sem dependências.
// Desenho: quadrado verde da casa com três barras subindo sobre a linha do horizonte.
// Uso: node web/ferramentas/gerar_icones.js   (o resultado é sempre o mesmo, byte a byte)

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const VERDE = [0x0f, 0x6e, 0x56];
const BRANCO = [0xff, 0xff, 0xff];
const MENTA = [0x9f, 0xe1, 0xcb];

// Formas em coordenadas de 0 a 1 dentro da "caixa do desenho".
const FORMAS = [
  { x0: 0.12, y0: 0.78, x1: 0.88, y1: 0.83, r: 0.025, cor: MENTA },    // horizonte
  { x0: 0.19, y0: 0.55, x1: 0.33, y1: 0.78, r: 0.03, cor: BRANCO },
  { x0: 0.43, y0: 0.40, x1: 0.57, y1: 0.78, r: 0.03, cor: BRANCO },
  { x0: 0.67, y0: 0.22, x1: 0.81, y1: 0.78, r: 0.03, cor: BRANCO },
];

function dentroRet(x, y, f) {
  if (x < f.x0 || x > f.x1 || y < f.y0 || y > f.y1) return false;
  const cx = Math.min(Math.max(x, f.x0 + f.r), f.x1 - f.r);
  const cy = Math.min(Math.max(y, f.y0 + f.r), f.y1 - f.r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= f.r * f.r;
}

// fundo: 'arredondado' (ícone comum) ou 'cheio' (maskable e iPhone, que recortam sozinhos).
function desenhar(tam, fundo, escala) {
  const px = Buffer.alloc(tam * tam * 4);
  const S = 4;   // 4×4 amostras por pixel: bordas suaves
  const raio = 0.2;
  const margem = (1 - escala) / 2;
  for (let py = 0; py < tam; py++) {
    for (let pxl = 0; pxl < tam; pxl++) {
      let a = 0, r = 0, g = 0, b = 0;
      for (let sy = 0; sy < S; sy++) {
        for (let sx = 0; sx < S; sx++) {
          const x = (pxl + (sx + 0.5) / S) / tam;
          const y = (py + (sy + 0.5) / S) / tam;
          if (fundo === 'arredondado' && !dentroRet(x, y, { x0: 0, y0: 0, x1: 1, y1: 1, r: raio })) continue;
          let cor = VERDE;
          const gx = (x - margem) / escala, gy = (y - margem) / escala;
          for (const f of FORMAS) if (dentroRet(gx, gy, f)) cor = f.cor;
          a++; r += cor[0]; g += cor[1]; b += cor[2];
        }
      }
      const i = (py * tam + pxl) * 4;
      if (a) { px[i] = Math.round(r / a); px[i + 1] = Math.round(g / a); px[i + 2] = Math.round(b / a); }
      px[i + 3] = Math.round((255 * a) / (S * S));
    }
  }
  return png(tam, px);
}

const CRC = new Int32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c; });
function crc32(buf) { let c = -1; for (const b of buf) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ -1) >>> 0; }
function bloco(tipo, dados) {
  const t = Buffer.from(tipo, 'ascii');
  const tam = Buffer.alloc(4); tam.writeUInt32BE(dados.length);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(Buffer.concat([t, dados])));
  return Buffer.concat([tam, t, dados, crc]);
}
function png(tam, rgba) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(tam, 0); ihdr.writeUInt32BE(tam, 4);
  ihdr[8] = 8; ihdr[9] = 6;   // 8 bits, RGBA
  const linhas = Buffer.alloc(tam * (tam * 4 + 1));
  for (let y = 0; y < tam; y++) rgba.copy(linhas, y * (tam * 4 + 1) + 1, y * tam * 4, (y + 1) * tam * 4);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), bloco('IHDR', ihdr),
    bloco('IDAT', zlib.deflateSync(linhas, { level: 9 })), bloco('IEND', Buffer.alloc(0))]);
}

const destino = path.join(__dirname, '..', '..', 'public', 'icones');
fs.mkdirSync(destino, { recursive: true });
const arquivos = {
  'icone-192.png': desenhar(192, 'arredondado', 1),
  'icone-512.png': desenhar(512, 'arredondado', 1),
  'icone-maskable-512.png': desenhar(512, 'cheio', 0.72),   // área segura do Android: círculo de 80%
  'apple-touch-icon.png': desenhar(180, 'cheio', 0.9),
  'favicone-32.png': desenhar(32, 'arredondado', 1),
};
for (const [nome, buf] of Object.entries(arquivos)) {
  fs.writeFileSync(path.join(destino, nome), buf);
  console.log(`${nome}: ${buf.length} bytes`);
}
