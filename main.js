/**
 * Web Digital · PDF Tools
 * Desenvolvido por: Murillo Matos
 * Powered by: Web Digital (https://webdigital.app)
 *
 * Converte documentos Office, junta com PDFs em um único arquivo e edita PDFs.
 */

const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');

let mainWindow;

// Executa um script PowerShell gravado num .ps1 temporário (com BOM UTF-8, para
// suportar acentos). Valores variáveis — como caminhos de arquivo — vão por
// variáveis de ambiente ($env:...), nunca interpolados no script: assim nomes
// com $, `, aspas ou acentos não quebram (nem injetam código) no PowerShell.
function runPowerShell(script, { env = {}, timeout = 180000 } = {}) {
    const tempScript = path.join(
        app.getPath('temp'),
        `wd_pdftools_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.ps1`
    );
    fs.writeFileSync(tempScript, '﻿' + script, { encoding: 'utf8' });

    return new Promise((resolve) => {
        exec(`powershell -NoProfile -ExecutionPolicy Bypass -File "${tempScript}"`, {
            timeout,
            windowsHide: true,
            maxBuffer: 1024 * 1024,
            env: { ...process.env, ...env }
        }, (error, stdout, stderr) => {
            try { fs.unlinkSync(tempScript); } catch (e) {}
            resolve({ error, stdout: (stdout || '').trim(), stderr: (stderr || '').trim() });
        });
    });
}

// Verificar se o Office (Word, Excel ou PowerPoint) está instalado.
// Usa o registro do Windows (ProgID COM registrado) — rápido e sem abrir o Office.
// Assíncrono para não travar a janela enquanto o PowerShell inicia.
async function checkOfficeInstalled() {
    const script = `$ErrorActionPreference = "SilentlyContinue"
$progIds = @("Word.Application", "Excel.Application", "PowerPoint.Application")
$found = $false
foreach ($p in $progIds) {
    if (Test-Path "Registry::HKEY_CLASSES_ROOT\\$p") { $found = $true; break }
}
if ($found) { exit 0 } else { exit 1 }`;

    try {
        const { error } = await runPowerShell(script, { timeout: 20000 });
        return { installed: !error };
    } catch (e) {
        return { installed: false };
    }
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 950,
        height: 750,
        minWidth: 700,
        minHeight: 550,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        },
        icon: path.join(__dirname, 'assets', 'icon.png'),
        titleBarStyle: 'default',
        show: false
    });

    mainWindow.loadFile('index.html');
    mainWindow.setMenuBarVisibility(false);
    
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    // Impede o Electron de navegar para arquivos arrastados (file:// URLs)
    mainWindow.webContents.on('will-navigate', (event) => {
        event.preventDefault();
    });
    // Nenhuma janela nova dentro do app: links http(s) abrem no navegador padrão
    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        if (/^https?:\/\//i.test(url)) shell.openExternal(url);
        return { action: 'deny' };
    });

    // Confirmação antes de fechar se houver trabalho não salvo
    mainWindow.on('close', async (event) => {
        event.preventDefault();

        let hasUnsaved = false;
        try {
            hasUnsaved = await mainWindow.webContents.executeJavaScript(`
                (function() {
                    const hasCombineFiles = typeof state !== 'undefined'
                        && Array.isArray(state.files) && state.files.length > 0;
                    const hasEdits = typeof editorState !== 'undefined'
                        && !!editorState.pdfDoc
                        && Array.isArray(editorState.edits)
                        && editorState.edits.some(p => Array.isArray(p) && p.length > 0);
                    return hasCombineFiles || hasEdits;
                })()
            `);
        } catch (e) {
            hasUnsaved = false;
        }

        if (hasUnsaved) {
            const { response } = await dialog.showMessageBox(mainWindow, {
                type: 'question',
                buttons: ['Fechar assim mesmo', 'Cancelar'],
                defaultId: 1,
                cancelId: 1,
                title: 'Fechar — Web Digital · PDF Tools',
                message: 'Você tem trabalho não salvo.',
                detail: 'Se fechar agora vai perder as alterações. Deseja continuar?',
                icon: path.join(__dirname, 'assets', 'icon.png')
            });
            if (response === 0) mainWindow.destroy();
            // response === 1 → cancelar, janela permanece aberta
        } else {
            mainWindow.destroy();
        }
    });
}

// ---- Atualização automática (GitHub Releases, via electron-updater) ----
// Só funciona no app instalado (Setup/NSIS): em desenvolvimento não há o que
// atualizar e a versão portátil não consegue substituir o próprio .exe.
const UPDATE_CHECK_INTERVAL = 4 * 60 * 60 * 1000; // a cada 4 horas

function sendUpdateStatus(status, data = {}) {
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('update-status', { status, ...data });
    }
}

function setupAutoUpdater() {
    if (!app.isPackaged || process.env.PORTABLE_EXECUTABLE_DIR) return;

    const { autoUpdater } = require('electron-updater');
    autoUpdater.autoDownload = true;          // baixa em segundo plano
    autoUpdater.autoInstallOnAppQuit = true;  // se o usuário só fechar, instala ao sair

    autoUpdater.on('update-available', (info) => sendUpdateStatus('downloading', { version: info.version }));
    autoUpdater.on('download-progress', (p) => sendUpdateStatus('downloading', { percent: Math.round(p.percent) }));
    autoUpdater.on('update-downloaded', (info) => sendUpdateStatus('ready', { version: info.version }));
    // Sem internet / GitHub fora do ar: falha silenciosa, tenta de novo no próximo ciclo
    autoUpdater.on('error', (err) => console.log('Atualização:', err && err.message));

    const check = () => autoUpdater.checkForUpdates().catch(() => {});
    check();
    setInterval(check, UPDATE_CHECK_INTERVAL);

    ipcMain.handle('install-update', () => {
        // Reinicia já com a nova versão (instalação silenciosa, reabre o app)
        mainWindow.removeAllListeners('close');
        autoUpdater.quitAndInstall(true, true);
    });
}

app.whenReady().then(() => {
    createWindow();
    setupAutoUpdater();
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// IPC Handlers
ipcMain.handle('check-office', async () => checkOfficeInstalled());

ipcMain.handle('select-files', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile', 'multiSelections'],
        filters: [
            { name: 'Documentos e PDFs', extensions: ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'rtf', 'txt'] }
        ]
    });
    return result.filePaths;
});

ipcMain.handle('select-pdf', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [{ name: 'PDF', extensions: ['pdf'] }]
    });
    return result.filePaths;
});

ipcMain.handle('save-file', async (event, defaultName) => {
    const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: defaultName || 'documento_combinado.pdf',
        filters: [{ name: 'PDF', extensions: ['pdf'] }]
    });
    return result.filePath || null;
});

// Converter documento Office para PDF
// Os caminhos chegam ao script por $env:WD_INPUT / $env:WD_OUTPUT (ver runPowerShell).
const WORD_SCRIPT = `
$ErrorActionPreference = "Stop"
$word = $null
$doc = $null

try {
    # Criar instância do Word
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone

    # Abrir documento (ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false)
    $doc = $word.Documents.Open($env:WD_INPUT, $false, $true, $false)

    # Aguardar impressão/paginação em segundo plano (com limite de ~10s)
    $waited = 0
    while ($word.BackgroundPrintingStatus -gt 0 -and $waited -lt 100) {
        Start-Sleep -Milliseconds 100
        $waited++
    }

    # Salvar como PDF (FileFormat 17 = wdFormatPDF)
    $doc.ExportAsFixedFormat($env:WD_OUTPUT, 17, $false, 0, 0, 0, 0, 0, $false, $false, 0, $false, $true, $false)

    Write-Output "SUCCESS"
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
} finally {
    # Fechar documento
    if ($doc) {
        try { $doc.Close(0) } catch {}  # wdDoNotSaveChanges = 0
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
    }
    # Fechar Word
    if ($word) {
        try { $word.Quit(0) } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
    # Forçar garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}`;

const EXCEL_SCRIPT = `
$ErrorActionPreference = "Stop"
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($env:WD_INPUT, 0, $true)
    $workbook.ExportAsFixedFormat(0, $env:WD_OUTPUT)

    Write-Output "SUCCESS"
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
} finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}`;

const POWERPOINT_SCRIPT = `
$ErrorActionPreference = "Stop"
$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application

    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $presentation = $ppt.Presentations.Open($env:WD_INPUT, $true, $false, $false)

    # SaveAs com formato 32 = ppSaveAsPDF
    $presentation.SaveAs($env:WD_OUTPUT, 32)

    Write-Output "SUCCESS"
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
} finally {
    if ($presentation) {
        try { $presentation.Close() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
    }
    if ($ppt) {
        # O PowerPoint é instância única: só fecha se o usuário não tiver
        # outras apresentações abertas (senão fecharia o trabalho dele).
        try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}`;

ipcMain.handle('convert-office-to-pdf', async (event, { filePath, tempFolder }) => {
    const ext = path.extname(filePath).toLowerCase();
    const fileName = path.basename(filePath, ext);
    const outputPath = path.join(tempFolder, `${fileName}_${Date.now()}.pdf`);

    let psScript;
    if (['.doc', '.docx', '.rtf', '.txt'].includes(ext)) psScript = WORD_SCRIPT;
    else if (['.xls', '.xlsx'].includes(ext)) psScript = EXCEL_SCRIPT;
    else if (['.ppt', '.pptx'].includes(ext)) psScript = POWERPOINT_SCRIPT;
    else return { success: false, error: 'Formato não suportado' };

    if (!fs.existsSync(filePath)) {
        return { success: false, error: 'Arquivo não encontrado (foi movido ou excluído?)' };
    }

    const { error, stdout: output, stderr: errorOutput } = await runPowerShell(psScript, {
        env: { WD_INPUT: filePath, WD_OUTPUT: outputPath }
    });

    if (errorOutput) console.log('PowerShell stderr:', errorOutput);

    if (output.includes('SUCCESS')) {
        // Verificar se o PDF foi criado
        if (fs.existsSync(outputPath)) return { success: true, pdfPath: outputPath };
        return { success: false, error: 'PDF não foi gerado no caminho esperado' };
    }
    if (output.includes('ERROR:')) {
        return { success: false, error: output.split('ERROR:')[1]?.trim() || 'Erro na conversão' };
    }
    if (error && error.killed) {
        return { success: false, error: 'Tempo esgotado na conversão (o Office pode estar aguardando uma resposta)' };
    }
    return { success: false, error: error?.message || errorOutput || 'Erro desconhecido na conversão' };
});

// Ler arquivo PDF e devolver como base64.
// Dados binários trafegam como string (base64) através do contextBridge de forma confiável.
ipcMain.handle('read-pdf-file', async (event, filePath) => {
    try {
        const buffer = fs.readFileSync(filePath);
        return { success: true, data: buffer.toString('base64') };
    } catch (e) {
        return { success: false, error: e.message };
    }
});

// Salvar PDF final (recebe base64)
ipcMain.handle('save-pdf', async (event, { data, outputPath }) => {
    try {
        fs.writeFileSync(outputPath, Buffer.from(data, 'base64'));
        return { success: true };
    } catch (e) {
        if (e.code === 'EBUSY' || e.code === 'EPERM') {
            return { success: false, error: 'O arquivo está aberto em outro programa ou sem permissão de escrita. Feche-o e tente novamente.' };
        }
        return { success: false, error: e.message };
    }
});

// Escolher a pasta de destino dos arquivos separados
ipcMain.handle('select-folder', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory', 'createDirectory']
    });
    return result.canceled ? null : result.filePaths[0];
});

// Salvar um PDF numa pasta sem sobrescrever nada: o nome é higienizado
// (só o nome do arquivo, sem caminho) e ganha " (2)", " (3)"... se já existir.
ipcMain.handle('save-pdf-in-folder', async (event, { data, folder, fileName }) => {
    try {
        let base = path.basename(String(fileName || ''))
            .replace(/\.pdf$/i, '')
            .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
            .replace(/[. ]+$/, '')
            .trim() || 'documento';
        let outputPath = path.join(folder, base + '.pdf');
        for (let n = 2; fs.existsSync(outputPath); n++) {
            outputPath = path.join(folder, `${base} (${n}).pdf`);
        }
        fs.writeFileSync(outputPath, Buffer.from(data, 'base64'));
        return { success: true, path: outputPath };
    } catch (e) {
        return { success: false, error: e.message };
    }
});

// Limpar arquivo temporário
ipcMain.handle('delete-temp-file', async (event, filePath) => {
    try { fs.unlinkSync(filePath); } catch (e) {}
});

// Obter pasta temp
ipcMain.handle('get-temp-folder', async () => app.getPath('temp'));

// Versão do app (vem do package.json — o rodapé nunca fica desatualizado)
ipcMain.handle('get-version', async () => app.getVersion());

// Abrir arquivo/pasta
ipcMain.handle('open-file', async (event, filePath) => shell.openPath(filePath));
ipcMain.handle('open-folder', async (event, folderPath) => shell.showItemInFolder(folderPath));
ipcMain.handle('open-external', async (event, url) => {
    // Só abre links web — nunca file://, programas ou outros protocolos
    if (/^https?:\/\//i.test(String(url))) return shell.openExternal(url);
});
