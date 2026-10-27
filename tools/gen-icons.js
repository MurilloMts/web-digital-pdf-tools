/**
 * Gerador de ícones — Web Digital · PDF Tools
 * Rasteriza assets/icon.svg (marca Web Digital) com o próprio Chromium do
 * Electron e empacota em PNG (256) e ICO multi-resolução (PNG embutido).
 *
 * Uso: npm run gen:icons   (equivale a: electron tools/gen-icons.js)
 */
const { app, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

const assetsDir = path.join(__dirname, '..', 'assets');
const sizes = [256, 128, 64, 48, 32, 24, 16];

// ---- Encoder ICO (PNG embutido) ----
function encodeICO(entries) {
    const header = Buffer.alloc(6);
    header.writeUInt16LE(0, 0); header.writeUInt16LE(1, 2); header.writeUInt16LE(entries.length, 4);
    const dir = Buffer.alloc(16 * entries.length);
    let offset = 6 + 16 * entries.length;
    entries.forEach((e, i) => {
        const o = i * 16;
        dir[o] = e.size >= 256 ? 0 : e.size;
        dir[o + 1] = e.size >= 256 ? 0 : e.size;
        dir[o + 2] = 0; dir[o + 3] = 0;
        dir.writeUInt16LE(1, o + 4); dir.writeUInt16LE(32, o + 6);
        dir.writeUInt32LE(e.png.length, o + 8); dir.writeUInt32LE(offset, o + 12);
        offset += e.png.length;
    });
    return Buffer.concat([header, dir, ...entries.map(e => e.png)]);
}

app.whenReady().then(async () => {
    const svg = fs.readFileSync(path.join(assetsDir, 'icon.svg'), 'utf8');
    const win = new BrowserWindow({ show: false, webPreferences: { offscreen: true } });
    await win.loadURL('data:text/html,<!doctype html><html><body></body></html>');

    // Desenha o SVG num canvas em cada tamanho e devolve os PNGs em base64
    const pngs = await win.webContents.executeJavaScript(`
        (async () => {
            const img = new Image();
            img.src = 'data:image/svg+xml;base64,' + ${JSON.stringify(Buffer.from(svg).toString('base64'))};
            await img.decode();
            return ${JSON.stringify(sizes)}.map(size => {
                const c = document.createElement('canvas');
                c.width = c.height = size;
                const ctx = c.getContext('2d');
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(img, 0, 0, size, size);
                return c.toDataURL('image/png').split(',')[1];
            });
        })()
    `);

    const entries = sizes.map((size, i) => ({ size, png: Buffer.from(pngs[i], 'base64') }));
    fs.writeFileSync(path.join(assetsDir, 'icon.png'), entries[0].png);
    fs.writeFileSync(path.join(assetsDir, 'icon.ico'), encodeICO(entries));
    console.log('OK  icon.png (256) e icon.ico (' + sizes.join(',') + ') gerados em assets/');
    app.quit();
});
