/**
 * Web Digital · PDF Tools — Renderer
 * Desenvolvido por: Murillo Matos
 * Powered by: Web Digital (https://webdigital.app)
 */

// ---- Helpers de caminho (sem Node no renderer: contextIsolation ativo) ----
function extname(p) {
    const base = String(p).split(/[\\/]/).pop();
    const dot = base.lastIndexOf('.');
    return dot > 0 ? base.slice(dot) : '';
}
function basename(p) {
    return String(p).split(/[\\/]/).pop();
}
function hexToRgb(hex) {
    return {
        r: parseInt(hex.slice(1, 3), 16) / 255,
        g: parseInt(hex.slice(3, 5), 16) / 255,
        b: parseInt(hex.slice(5, 7), 16) / 255
    };
}
// Escapa texto vindo do sistema de arquivos antes de inserir no HTML (evita injeção)
function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, c => (
        { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
    ));
}

// Conversão base64 <-> bytes (PDFs trafegam como base64 pelo contextBridge)
function base64ToBytes(b64) {
    const bin = atob(b64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}
function bytesToBase64(bytes) {
    let bin = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(bin);
}

const state = {
    files: [],
    officeInstalled: false,
    isProcessing: false,
    outputPath: null,
    tempFiles: []
};

const OFFICE_EXTENSIONS = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.rtf', '.txt'];
const PDF_EXTENSION = '.pdf';

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileList = document.getElementById('fileList');
const actionBar = document.getElementById('actionBar');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const progressPercent = document.getElementById('progressPercent');
const resultsSection = document.getElementById('resultsSection');
const resultsText = document.getElementById('resultsText');
const statusBadge = document.getElementById('statusBadge');
const officeAlert = document.getElementById('officeAlert');
const combineBtn = document.getElementById('combineBtn');

async function init() {
    setupUpdates();
    await checkOffice();
    setupDropZone();
    setupDragAndDrop();
}

// ---- Atualização automática ----
function setupUpdates() {
    window.api.getVersion().then(v => {
        document.getElementById('appVersion').textContent = 'v' + v;
    });

    const badge = document.getElementById('updateBadge');
    const text = badge.querySelector('.update-text');
    window.api.onUpdateStatus(({ status, version, percent }) => {
        badge.style.display = 'flex';
        if (status === 'downloading') {
            badge.disabled = true;
            badge.classList.remove('ready');
            text.textContent = percent != null
                ? `Baixando atualização... ${percent}%`
                : `Baixando v${version}...`;
        } else if (status === 'ready') {
            badge.disabled = false;
            badge.classList.add('ready');
            badge.title = `Versão ${version} pronta. Se preferir, ela também é instalada ao fechar o app.`;
            text.textContent = `Reiniciar e atualizar (v${version})`;
        }
    });
}

function hasUnsavedWork() {
    const hasEdits = !!editorState.pdfDoc
        && editorState.edits.some(p => Array.isArray(p) && p.length > 0);
    return state.files.length > 0 || hasEdits || state.isProcessing;
}

function installUpdate() {
    if (hasUnsavedWork() &&
        !confirm('Você tem trabalho não salvo que será perdido ao reiniciar. Atualizar agora mesmo assim?')) {
        return;
    }
    window.api.installUpdate();
}

async function checkOffice() {
    const result = await window.api.checkOffice();
    state.officeInstalled = result.installed;

    const statusDot = statusBadge.querySelector('.status-dot');
    const statusText = statusBadge.querySelector('.status-text');

    if (result.installed) {
        statusDot.classList.add('online');
        statusText.textContent = 'Office OK';
        officeAlert.style.display = 'none';
    } else {
        statusDot.classList.add('offline');
        statusText.textContent = 'Office não encontrado';
        officeAlert.style.display = 'flex';
    }
}

function setupDropZone() {
    dropZone.addEventListener('click', async () => {
        const files = await window.api.selectFiles();
        if (files.length > 0) addFiles(files);
    });
}

function setupDragAndDrop() {
    // Nível window: impede o Chromium de navegar para arquivos arrastados de fora
    window.addEventListener('dragover',  (e) => { e.preventDefault(); e.stopPropagation(); });
    window.addEventListener('drop',      (e) => { e.preventDefault(); e.stopPropagation(); });
    window.addEventListener('dragenter', (e) => { e.preventDefault(); e.stopPropagation(); });

    // Drop zone da aba Combinar
    dropZone.addEventListener('dragenter', (e) => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); e.stopPropagation();
        e.dataTransfer.dropEffect = 'copy';
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', (e) => {
        e.stopPropagation();
        if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('dragover');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragover');
        const paths = Array.from(e.dataTransfer.files)
            .filter(f => isValidFile(f.name))
            .map(f => f.path)
            .filter(Boolean);
        if (paths.length > 0) addFiles(paths);
    });
}

function isValidFile(filename) {
    const ext = extname(filename).toLowerCase();
    return ext === PDF_EXTENSION || OFFICE_EXTENSIONS.includes(ext);
}

function addFiles(filePaths) {
    for (const filePath of filePaths) {
        if (!state.files.find(f => f.path === filePath)) {
            const ext = extname(filePath).toLowerCase();
            state.files.push({
                path: filePath,
                name: basename(filePath),
                ext: ext,
                isPDF: ext === PDF_EXTENSION,
                status: 'pending'
            });
        }
    }
    updateUI();
}

function updateUI() {
    renderFileList();
    const hasFiles = state.files.length > 0;
    actionBar.style.display = hasFiles && !state.isProcessing ? 'flex' : 'none';
    document.getElementById('sortBar').style.display = hasFiles && !state.isProcessing ? 'flex' : 'none';

    // Verificar se pode combinar
    const hasOfficeFiles = state.files.some(f => !f.isPDF);
    combineBtn.disabled = state.isProcessing || (hasOfficeFiles && !state.officeInstalled);
}

// Funções de ordenação
function sortFiles(method) {
    if (state.isProcessing) return;
    switch (method) {
        case 'name-asc':
            state.files.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { numeric: true }));
            break;
        case 'name-desc':
            state.files.sort((a, b) => b.name.localeCompare(a.name, 'pt-BR', { numeric: true }));
            break;
        case 'number':
            state.files.sort((a, b) => {
                const numA = extractNumber(a.name);
                const numB = extractNumber(b.name);
                if (numA !== null && numB !== null) return numA - numB;
                if (numA !== null) return -1;
                if (numB !== null) return 1;
                return a.name.localeCompare(b.name, 'pt-BR', { numeric: true });
            });
            break;
        case 'type':
            state.files.sort((a, b) => {
                // PDF primeiro, depois por nome
                if (a.isPDF && !b.isPDF) return -1;
                if (!a.isPDF && b.isPDF) return 1;
                return a.name.localeCompare(b.name, 'pt-BR', { numeric: true });
            });
            break;
    }
    renderFileList();
}

function extractNumber(filename) {
    // Extrai o primeiro número do nome do arquivo
    const match = filename.match(/^(\d+)/);
    if (match) return parseInt(match[1]);

    // Tenta encontrar número após hífen, ponto, underscore ou espaço
    const match2 = filename.match(/[-._\s](\d+)/);
    if (match2) return parseInt(match2[1]);

    return null;
}

function renderFileList() {
    if (state.files.length === 0) {
        fileList.innerHTML = '';
        return;
    }

    fileList.innerHTML = state.files.map((file, index) => `
        <div class="file-item ${file.status}" draggable="${!state.isProcessing}" data-index="${index}">
            <div class="file-drag-handle">⋮⋮</div>
            <div class="file-icon">${getFileIcon(file.ext)}</div>
            <div class="file-info">
                <div class="file-name">${escapeHtml(file.name)}</div>
                <div class="file-type">${getFileType(file.ext)}</div>
            </div>
            <div class="file-status-indicator">
                ${file.status === 'converting' ? '<div class="spinner-small"></div>' : ''}
                ${file.status === 'ready' ? '<span class="status-icon success">✓</span>' : ''}
                ${file.status === 'error' ? '<span class="status-icon error" title="' + escapeHtml(file.error || '') + '">✕</span>' : ''}
            </div>
            <button class="btn-icon remove" onclick="removeFile(${index})" title="Remover" ${state.isProcessing ? 'disabled' : ''}>✕</button>
        </div>
    `).join('');

    setupFileDragReorder();
}

function getFileIcon(ext) {
    const icons = {
        '.pdf': '📕', '.doc': '📘', '.docx': '📘',
        '.xls': '📗', '.xlsx': '📗', '.ppt': '📙', '.pptx': '📙',
        '.rtf': '📄', '.txt': '📄'
    };
    return icons[ext] || '📄';
}

function getFileType(ext) {
    const types = {
        '.pdf': 'PDF', '.doc': 'Word', '.docx': 'Word',
        '.xls': 'Excel', '.xlsx': 'Excel', '.ppt': 'PowerPoint', '.pptx': 'PowerPoint',
        '.rtf': 'Rich Text', '.txt': 'Texto'
    };
    return types[ext] || 'Documento';
}

function setupFileDragReorder() {
    const items = fileList.querySelectorAll('.file-item');
    items.forEach(item => {
        item.addEventListener('dragstart', (e) => {
            item.classList.add('dragging');
            e.dataTransfer.setData('text/plain', item.dataset.index);
        });
        item.addEventListener('dragend', () => item.classList.remove('dragging'));
        item.addEventListener('dragover', (e) => {
            e.preventDefault();
            item.classList.add('drag-over');
        });
        item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
        item.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            item.classList.remove('drag-over');
            if (state.isProcessing) return;
            // Arquivos vindos do Explorer soltos sobre a lista: adiciona em vez de reordenar
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                const paths = Array.from(e.dataTransfer.files)
                    .filter(f => isValidFile(f.name))
                    .map(f => f.path)
                    .filter(Boolean);
                if (paths.length > 0) addFiles(paths);
                return;
            }
            const fromIndex = parseInt(e.dataTransfer.getData('text/plain'), 10);
            const toIndex = parseInt(item.dataset.index, 10);
            if (Number.isInteger(fromIndex) && Number.isInteger(toIndex) && fromIndex !== toIndex) {
                const [moved] = state.files.splice(fromIndex, 1);
                state.files.splice(toIndex, 0, moved);
                renderFileList();
            }
        });
    });
}

function removeFile(index) {
    if (state.isProcessing) return;
    state.files.splice(index, 1);
    updateUI();
}

function clearFiles() {
    if (state.isProcessing) return;
    state.files = [];
    resultsSection.style.display = 'none';
    progressSection.style.display = 'none';
    updateUI();
}

function updateProgress(percent, text) {
    progressFill.style.width = `${percent}%`;
    progressPercent.textContent = `${Math.round(percent)}%`;
    if (text) progressText.textContent = text;
}

async function combineAndSave() {
    if (state.files.length === 0) return;

    // Pedir onde salvar
    const outputPath = await window.api.saveFile('documento_combinado.pdf');
    if (!outputPath) return;

    state.outputPath = outputPath;
    state.isProcessing = true;
    state.tempFiles = [];

    // Recomeça os indicadores (uma nova tentativa não herda status antigos)
    state.files.forEach(f => { f.status = 'pending'; f.error = null; });

    progressSection.style.display = 'block';
    actionBar.style.display = 'none';
    document.getElementById('sortBar').style.display = 'none';
    resultsSection.style.display = 'none';
    updateProgress(0, 'Iniciando...');

    // try/finally: mesmo com erro inesperado a interface nunca fica travada em "processando"
    try {
        await processAndMerge();
    } catch (e) {
        alert('Erro inesperado: ' + e.message);
    } finally {
        // Limpar arquivos temporários
        for (const tempFile of state.tempFiles) {
            await window.api.deleteTempFile(tempFile);
        }
        state.tempFiles = [];
        state.isProcessing = false;
        progressSection.style.display = 'none';
        updateUI();
    }
}

async function processAndMerge() {
    const tempFolder = await window.api.getTempFolder();
    const pdfBuffers = [];
    // Cópia: a ordem não pode mudar no meio do processamento
    const files = state.files.slice();
    const totalFiles = files.length;
    let processedFiles = 0;
    let errors = 0;

    // Processar cada arquivo
    for (const file of files) {
        const fileProgress = (processedFiles / totalFiles) * 100;

        if (file.isPDF) {
            // Já é PDF, só ler
            updateProgress(fileProgress, `Lendo: ${file.name}`);
            file.status = 'converting';
            renderFileList();

            const result = await window.api.readPdfFile(file.path);
            if (result.success) {
                pdfBuffers.push({ file, name: file.name, data: base64ToBytes(result.data) });
                file.status = 'ready';
            } else {
                file.status = 'error';
                file.error = result.error;
                errors++;
            }
        } else {
            // Converter Office para PDF
            updateProgress(fileProgress, `Convertendo: ${file.name}`);
            file.status = 'converting';
            renderFileList();

            const result = await window.api.convertOfficeToPdf({
                filePath: file.path,
                tempFolder: tempFolder
            });

            if (result.success) {
                const pdfResult = await window.api.readPdfFile(result.pdfPath);
                if (pdfResult.success) {
                    pdfBuffers.push({ file, name: file.name, data: base64ToBytes(pdfResult.data) });
                    file.status = 'ready';
                    state.tempFiles.push(result.pdfPath);
                } else {
                    file.status = 'error';
                    file.error = 'Erro ao ler PDF convertido';
                    errors++;
                }
            } else {
                file.status = 'error';
                file.error = result.error;
                errors++;
            }
        }

        processedFiles++;
        renderFileList();
    }

    // Combinar PDFs
    if (pdfBuffers.length > 0) {
        updateProgress(90, 'Combinando PDFs...');

        try {
            const mergedPdf = await PDFLib.PDFDocument.create();
            let merged = 0;

            for (const pdfBuffer of pdfBuffers) {
                try {
                    // ignoreEncryption: PDFs só com restrição de edição (sem senha de abertura) ainda podem ser juntados
                    const pdf = await PDFLib.PDFDocument.load(pdfBuffer.data, { ignoreEncryption: true });
                    const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                    pages.forEach(page => mergedPdf.addPage(page));
                    merged++;
                } catch (e) {
                    console.error(`Erro ao processar ${pdfBuffer.name}:`, e);
                    pdfBuffer.file.status = 'error';
                    pdfBuffer.file.error = 'PDF inválido ou protegido: ' + e.message;
                    errors++;
                }
            }
            renderFileList();

            if (merged === 0) {
                alert('Nenhum arquivo pôde ser combinado. Passe o mouse sobre o ✕ para ver o motivo.');
                return;
            }

            const mergedBytes = await mergedPdf.save();

            updateProgress(95, 'Salvando...');
            const saveResult = await window.api.savePdf({
                data: bytesToBase64(mergedBytes),
                outputPath: state.outputPath
            });

            if (saveResult.success) {
                updateProgress(100, 'Concluído!');
                resultsText.textContent = errors > 0
                    ? `PDF criado com ${merged} arquivo(s) — ${errors} com erro (passe o mouse sobre o ✕ para ver o motivo)`
                    : `PDF criado com sucesso! (${merged} arquivos combinados)`;
                resultsSection.style.display = 'block';
            } else {
                alert('Erro ao salvar: ' + saveResult.error);
            }
        } catch (e) {
            alert('Erro ao combinar PDFs: ' + e.message);
        }
    } else {
        alert('Nenhum arquivo pôde ser preparado para combinar.');
    }
}

function openOutputFile() {
    if (state.outputPath) window.api.openFile(state.outputPath);
}

function openOutputFolder() {
    if (state.outputPath) window.api.openFolder(state.outputPath);
}

function startNew() {
    state.files = [];
    state.outputPath = null;
    resultsSection.style.display = 'none';
    updateUI();
}

function openExternal(url) {
    window.api.openExternal(url);
}

// ==================== EDITOR DE PDF ====================
//
// As edições são armazenadas em coordenadas de página do PDF (pontos, origem
// no canto inferior esquerdo), independentes de zoom e rotação. A conversão
// tela <-> PDF usa o viewport do pdf.js (convertToPdfPoint / convertToViewportPoint),
// garantindo que o que aparece na tela é exatamente o que é gravado no arquivo.

const editorState = {
    pdfDoc: null,
    pdfBytes: null,
    pdfPath: null,
    currentPage: 1,
    totalPages: 0,
    scale: 1.5,        // escala real do pdf.js (1.0 = 100%)
    pageRotations: [], // rotação extra aplicada pelo usuário, por página
    viewport: null,    // viewport atual (para conversão de coordenadas)
    baseWidth: 0,      // largura do viewport em escala 1 (para "ajustar à tela")
    baseHeight: 0,     // altura do viewport em escala 1 (para "ajustar à tela")
    currentTool: 'erase',
    edits: [],         // edits[pageIndex] = lista de edições em coordenadas PDF
    selectedEdit: -1,  // índice da edição selecionada na página atual (-1 = nenhuma)
    pendingText: null, // edição de texto em andamento (campo inline aberto)
    textFormat: { bold: false, italic: false, underline: false, highlight: false },
    isDrawing: false,
    startX: 0,
    startY: 0,
    isFullscreen: false,
    isDragging: false,  // arrastar edição selecionada para mover
    dragOffsetX: 0,
    dragOffsetY: 0,
    dpr: 1,             // devicePixelRatio usado no render atual
    renderTask: null,   // renderização em andamento do pdf.js (cancelável)
    renderSeq: 0        // descarta renderizações obsoletas
};

// Worker do pdf.js empacotado localmente (funciona offline)
pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

function switchTab(tabName) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    document.getElementById(`tab-${tabName}`).classList.add('active');

    updateEditingLayout();
}

// Ativa o layout de tela cheia do editor apenas quando a aba "Editar" está ativa
// e há um PDF aberto (faz o editor ocupar toda a tela).
function updateEditingLayout() {
    const editTabActive = document.getElementById('tab-edit').classList.contains('active');
    const editorOpen = !!editorState.pdfDoc &&
        document.getElementById('editorContainer').style.display !== 'none';
    document.body.classList.toggle('editing', editTabActive && editorOpen);
}

// Setup do editor
function setupEditor() {
    const editDropZone = document.getElementById('editDropZone');

    editDropZone.addEventListener('click', async () => {
        const files = await window.api.selectPdf();
        if (files && files.length > 0) loadPdfForEdit(files[0]);
    });

    editDropZone.addEventListener('dragenter', (e) => {
        e.preventDefault(); e.stopPropagation();
        editDropZone.classList.add('dragover');
    });
    editDropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); e.stopPropagation();
        e.dataTransfer.dropEffect = 'copy';
        editDropZone.classList.add('dragover');
    });
    editDropZone.addEventListener('dragleave', (e) => {
        e.stopPropagation();
        if (!editDropZone.contains(e.relatedTarget)) editDropZone.classList.remove('dragover');
    });
    editDropZone.addEventListener('drop', (e) => {
        e.preventDefault(); e.stopPropagation();
        editDropZone.classList.remove('dragover');
        const file = Array.from(e.dataTransfer.files).find(f => f.name.toLowerCase().endsWith('.pdf'));
        if (file) loadPdfForEdit(file.path);
    });

    // Setup canvas de edição
    const editCanvas = document.getElementById('editCanvas');
    editCanvas.addEventListener('mousedown', startDraw);
    editCanvas.addEventListener('mousemove', draw);
    editCanvas.addEventListener('mouseup', endDraw);
    editCanvas.addEventListener('mouseleave', endDraw);

    // Campo de texto inline: confirma no Enter / cancela no Esc
    const inlineInput = document.getElementById('inlineTextInput');
    inlineInput.addEventListener('keydown', (e) => {
        e.stopPropagation();
        if (e.key === 'Enter') { e.preventDefault(); commitInlineText(); }
        else if (e.key === 'Escape') { e.preventDefault(); cancelInlineText(); }
    });
    // Acompanha a largura do texto enquanto digita
    inlineInput.addEventListener('input', () => {
        inlineInput.style.width = (inlineInput.scrollWidth + 4) + 'px';
    });

    // Controles de formatação atualizam o texto em digitação ao vivo
    ['toolFont', 'toolSize', 'toolColor', 'toolHighlight'].forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('input', () => {
            if (editorState.pendingText) {
                editorState.pendingText.font = document.getElementById('toolFont').value;
                editorState.pendingText.size = parseInt(document.getElementById('toolSize').value) || 16;
                editorState.pendingText.color = document.getElementById('toolColor').value;
                editorState.pendingText.highlight = editorState.textFormat.highlight
                    ? document.getElementById('toolHighlight').value : null;
                styleInlineInput();
            }
        });
    });
}

async function loadPdfForEdit(filePath) {
    try {
        editorState.pdfPath = filePath;

        const result = await window.api.readPdfFile(filePath);
        if (!result.success) {
            alert('Erro ao carregar PDF: ' + result.error);
            return;
        }

        // Decodificar base64 -> bytes. Decodifica duas vezes porque o pdf.js
        // "consome" (transfere) o buffer ao carregar, e o pdf-lib precisa de uma cópia íntegra ao salvar.
        editorState.pdfBytes = base64ToBytes(result.data);
        editorState.pdfDoc = await pdfjsLib.getDocument({ data: base64ToBytes(result.data) }).promise;
        editorState.totalPages = editorState.pdfDoc.numPages;
        editorState.currentPage = 1;
        editorState.edits = [];
        editorState.pageRotations = [];

        // Inicializar arrays para cada página
        for (let i = 0; i < editorState.totalPages; i++) {
            editorState.edits.push([]);
            editorState.pageRotations.push(0);
        }

        document.getElementById('editDropZone').style.display = 'none';
        document.getElementById('editorContainer').style.display = 'flex'; // flex: o canvas ocupa a altura restante
        updateEditingLayout();

        // Medir a página e renderizar já ajustado à tela. O requestAnimationFrame
        // garante que o layout amplo de edição já assentou antes de medir o contêiner.
        const probe = await editorState.pdfDoc.getPage(1);
        const baseVp = probe.getViewport({ scale: 1, rotation: probe.rotate || 0 });
        editorState.baseWidth = baseVp.width;
        editorState.baseHeight = baseVp.height;
        updatePageIndicator();
        requestAnimationFrame(() => zoomFit());
    } catch (e) {
        alert('Erro ao carregar PDF: ' + e.message);
    }
}

async function renderCurrentPage() {
    // Confirma texto em digitação enquanto o viewport atual ainda é válido
    if (editorState.pendingText) commitInlineText();
    if (!editorState.pdfDoc) return;

    // Cliques rápidos em zoom/página disparam várias renderizações: só a última vale
    const seq = ++editorState.renderSeq;
    const page = await editorState.pdfDoc.getPage(editorState.currentPage);
    if (seq !== editorState.renderSeq || !editorState.pdfDoc) return;
    if (editorState.renderTask) editorState.renderTask.cancel();

    const extra = editorState.pageRotations[editorState.currentPage - 1] || 0;
    const rotation = ((page.rotate || 0) + extra) % 360;
    const viewport = page.getViewport({ scale: editorState.scale, rotation });

    editorState.viewport = viewport;
    const baseVp = page.getViewport({ scale: 1, rotation });
    editorState.baseWidth = baseVp.width;
    editorState.baseHeight = baseVp.height;

    const pdfCanvas = document.getElementById('pdfCanvas');
    const editCanvas = document.getElementById('editCanvas');
    const wrapper = document.getElementById('canvasWrapper');

    // Canvas em resolução física (nítido em telas HiDPI / escala do Windows > 100%).
    // O tamanho CSS continua em unidades do viewport, então coordenadas de mouse não mudam.
    const dpr = window.devicePixelRatio || 1;
    editorState.dpr = dpr;
    for (const c of [pdfCanvas, editCanvas]) {
        c.width = Math.floor(viewport.width * dpr);
        c.height = Math.floor(viewport.height * dpr);
        c.style.width = viewport.width + 'px';
        c.style.height = viewport.height + 'px';
    }

    // Ajustar wrapper
    wrapper.style.width = viewport.width + 'px';
    wrapper.style.height = viewport.height + 'px';

    // Renderizar PDF
    const ctx = pdfCanvas.getContext('2d');
    ctx.clearRect(0, 0, pdfCanvas.width, pdfCanvas.height);
    const task = page.render({
        canvasContext: ctx,
        viewport: viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null
    });
    editorState.renderTask = task;
    try {
        await task.promise;
    } catch (e) {
        if (e && e.name === 'RenderingCancelledException') return;
        throw e;
    } finally {
        if (editorState.renderTask === task) editorState.renderTask = null;
    }

    // Renderizar edições existentes
    renderEdits();
}

// Converte um retângulo em coordenadas PDF para a caixa equivalente no viewport
function pdfRectToViewport(vp, x, y, w, h) {
    const c = [
        vp.convertToViewportPoint(x, y),
        vp.convertToViewportPoint(x + w, y),
        vp.convertToViewportPoint(x, y + h),
        vp.convertToViewportPoint(x + w, y + h)
    ];
    const xs = c.map(p => p[0]);
    const ys = c.map(p => p[1]);
    const minX = Math.min(...xs), minY = Math.min(...ys);
    return { x: minX, y: minY, width: Math.max(...xs) - minX, height: Math.max(...ys) - minY };
}

function renderEdits() {
    const editCanvas = document.getElementById('editCanvas');
    const ctx = editCanvas.getContext('2d');
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, editCanvas.width, editCanvas.height);

    const vp = editorState.viewport;
    if (!vp) return;
    // Desenha em unidades do viewport; a escala física (HiDPI) fica na transformação
    const dpr = editorState.dpr || 1;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    const pageEdits = editorState.edits[editorState.currentPage - 1] || [];

    for (const edit of pageEdits) {
        if (edit.type === 'erase' || edit.type === 'rect') {
            const r = pdfRectToViewport(vp, edit.x, edit.y, edit.width, edit.height);
            ctx.fillStyle = edit.color;
            ctx.fillRect(r.x, r.y, r.width, r.height);
        } else if (edit.type === 'highlight') {
            const r = pdfRectToViewport(vp, edit.x, edit.y, edit.width, edit.height);
            ctx.save();
            ctx.globalAlpha = 0.45;
            ctx.fillStyle = edit.color;
            ctx.fillRect(r.x, r.y, r.width, r.height);
            ctx.restore();
        } else if (edit.type === 'text') {
            const p = vp.convertToViewportPoint(edit.x, edit.y);
            const fontPx = edit.size * editorState.scale;
            ctx.font = canvasFontString(edit, editorState.scale);
            ctx.textBaseline = 'top';   // alinha com o campo inline (canto superior esq.)
            const w = ctx.measureText(edit.text).width;
            // Marca-texto (fundo)
            if (edit.highlight) {
                ctx.save();
                ctx.globalAlpha = 0.45;
                ctx.fillStyle = edit.highlight;
                ctx.fillRect(p[0], p[1], w, fontPx);
                ctx.restore();
            }
            ctx.fillStyle = edit.color;
            ctx.fillText(edit.text, p[0], p[1]);
            // Sublinhado
            if (edit.underline) {
                ctx.strokeStyle = edit.color;
                ctx.lineWidth = Math.max(1, fontPx * 0.06);
                const uy = p[1] + fontPx * 0.92;
                ctx.beginPath();
                ctx.moveTo(p[0], uy);
                ctx.lineTo(p[0] + w, uy);
                ctx.stroke();
            }
        }
    }

    // Destaque da edição selecionada (tracejado + alças nos cantos)
    if (editorState.selectedEdit >= 0 && editorState.selectedEdit < pageEdits.length) {
        const edit = pageEdits[editorState.selectedEdit];
        let box;
        if (edit.type === 'text') {
            const p = vp.convertToViewportPoint(edit.x, edit.y);
            ctx.font = canvasFontString(edit, editorState.scale);
            box = { x: p[0], y: p[1], width: ctx.measureText(edit.text).width, height: edit.size * editorState.scale };
        } else {
            box = pdfRectToViewport(vp, edit.x, edit.y, edit.width, edit.height);
        }
        ctx.save();
        ctx.strokeStyle = '#c09e6f';
        ctx.lineWidth = 2;
        ctx.setLineDash([6, 4]);
        ctx.strokeRect(box.x - 3, box.y - 3, box.width + 6, box.height + 6);
        ctx.setLineDash([]);
        ctx.fillStyle = '#c09e6f';
        const hs = 5;
        [[box.x - 3, box.y - 3], [box.x + box.width + 3, box.y - 3],
         [box.x - 3, box.y + box.height + 3], [box.x + box.width + 3, box.y + box.height + 3]]
            .forEach(c => ctx.fillRect(c[0] - hs / 2, c[1] - hs / 2, hs, hs));
        ctx.restore();
    }
}

function setTool(tool) {
    editorState.currentTool = tool;
    editorState.selectedEdit = -1;
    commitInlineText();
    document.querySelectorAll('.tool-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelector(`[data-tool="${tool}"]`).classList.add('active');

    // Cor padrão sensata ao trocar de ferramenta
    const colorInput = document.getElementById('toolColor');
    if (tool === 'erase') colorInput.value = '#ffffff';
    else if (tool === 'text') colorInput.value = '#000000';
    else if (tool === 'highlight') colorInput.value = '#ffff00';

    // Cursor conforme a ferramenta
    const editCanvas = document.getElementById('editCanvas');
    if (editCanvas) editCanvas.style.cursor = (tool === 'select') ? 'pointer' : 'crosshair';

    if (editorState.viewport) renderEdits();
}

// ---- Formatação de texto ----
function canvasFontFamily(family) {
    if (family === 'Times') return "'Times New Roman', Times, serif";
    if (family === 'Courier') return "'Courier New', Courier, monospace";
    return 'Arial, Helvetica, sans-serif';
}
function canvasFontString(edit, scale) {
    let s = '';
    if (edit.italic) s += 'italic ';
    if (edit.bold) s += 'bold ';
    s += (edit.size * scale) + 'px ' + canvasFontFamily(edit.font);
    return s;
}
function pdfFontName(family, bold, italic) {
    const F = PDFLib.StandardFonts;
    if (family === 'Times') {
        if (bold && italic) return F.TimesRomanBoldItalic;
        if (bold) return F.TimesRomanBold;
        if (italic) return F.TimesRomanItalic;
        return F.TimesRoman;
    }
    if (family === 'Courier') {
        if (bold && italic) return F.CourierBoldOblique;
        if (bold) return F.CourierBold;
        if (italic) return F.CourierOblique;
        return F.Courier;
    }
    if (bold && italic) return F.HelveticaBoldOblique;
    if (bold) return F.HelveticaBold;
    if (italic) return F.HelveticaOblique;
    return F.Helvetica;
}
function hexToRgba(hex, a) {
    const c = hexToRgb(hex);
    return `rgba(${Math.round(c.r * 255)},${Math.round(c.g * 255)},${Math.round(c.b * 255)},${a})`;
}

// Lê o formato atual dos controles da barra de ferramentas
function readTextFormat() {
    const fontEl = document.getElementById('toolFont');
    return {
        font: fontEl ? fontEl.value : 'Helvetica',
        size: parseInt(document.getElementById('toolSize').value) || 16,
        color: document.getElementById('toolColor').value,
        bold: editorState.textFormat.bold,
        italic: editorState.textFormat.italic,
        underline: editorState.textFormat.underline,
        highlight: editorState.textFormat.highlight ? document.getElementById('toolHighlight').value : null
    };
}

// Aplica o formato pendente ao campo inline (pré-visualização ao vivo)
function styleInlineInput() {
    const input = document.getElementById('inlineTextInput');
    const f = editorState.pendingText;
    if (!f) return;
    input.style.fontFamily = canvasFontFamily(f.font);
    input.style.fontSize = (f.size * editorState.scale) + 'px';
    input.style.fontWeight = f.bold ? 'bold' : 'normal';
    input.style.fontStyle = f.italic ? 'italic' : 'normal';
    input.style.textDecoration = f.underline ? 'underline' : 'none';
    input.style.color = f.color;
    input.style.background = f.highlight ? hexToRgba(f.highlight, 0.45) : 'transparent';
}

// ---- Texto inline: campo editável posicionado no local do clique ----
function showInlineTextInput(vx, vy) {
    const input = document.getElementById('inlineTextInput');
    editorState.pendingText = Object.assign({ vx, vy }, readTextFormat());

    input.value = '';
    input.style.display = 'block';
    input.style.left = vx + 'px';
    input.style.top = vy + 'px';
    input.style.width = '60px';
    styleInlineInput();
    setTimeout(() => input.focus(), 0);
}

function commitInlineText() {
    const input = document.getElementById('inlineTextInput');
    if (input.style.display === 'none' || !editorState.pendingText) return;

    const text = input.value;
    const pt = editorState.pendingText;
    input.style.display = 'none';
    input.value = '';
    editorState.pendingText = null;

    if (text && text.trim() && editorState.viewport) {
        // (vx, vy) é o canto superior esquerdo do texto na tela -> ponto PDF
        const p = editorState.viewport.convertToPdfPoint(pt.vx, pt.vy);
        editorState.edits[editorState.currentPage - 1].push({
            type: 'text',
            x: p[0], y: p[1],
            text: text,
            color: pt.color,
            size: pt.size,        // tamanho em pontos do PDF
            font: pt.font,
            bold: pt.bold,
            italic: pt.italic,
            underline: pt.underline,
            highlight: pt.highlight
        });
        renderEdits();
    }
}

function cancelInlineText() {
    const input = document.getElementById('inlineTextInput');
    input.style.display = 'none';
    input.value = '';
    editorState.pendingText = null;
}

// Alterna negrito/itálico/sublinhado/marca-texto (aplica ao vivo se estiver digitando)
function toggleFormat(key) {
    editorState.textFormat[key] = !editorState.textFormat[key];
    const btn = document.getElementById('fmt-' + key);
    if (btn) btn.classList.toggle('active', editorState.textFormat[key]);
    if (editorState.pendingText) {
        if (key === 'highlight') {
            editorState.pendingText.highlight = editorState.textFormat.highlight
                ? document.getElementById('toolHighlight').value : null;
        } else {
            editorState.pendingText[key] = editorState.textFormat[key];
        }
        styleInlineInput();
        document.getElementById('inlineTextInput').focus();
    }
}

// Localiza a edição sob o ponto (coords de tela). Retorna o índice ou -1.
function hitTestEdit(vx, vy) {
    const vp = editorState.viewport;
    if (!vp) return -1;
    const pageEdits = editorState.edits[editorState.currentPage - 1] || [];
    const ctx = document.getElementById('editCanvas').getContext('2d');
    for (let i = pageEdits.length - 1; i >= 0; i--) {
        const edit = pageEdits[i];
        if (edit.type === 'text') {
            const p = vp.convertToViewportPoint(edit.x, edit.y);
            ctx.font = canvasFontString(edit, editorState.scale);
            const w = ctx.measureText(edit.text).width;
            const h = edit.size * editorState.scale;
            if (vx >= p[0] - 2 && vx <= p[0] + w + 2 && vy >= p[1] - 2 && vy <= p[1] + h + 2) return i;
        } else {
            const r = pdfRectToViewport(vp, edit.x, edit.y, edit.width, edit.height);
            if (vx >= r.x && vx <= r.x + r.width && vy >= r.y && vy <= r.y + r.height) return i;
        }
    }
    return -1;
}

function deleteSelected() {
    const pageEdits = editorState.edits[editorState.currentPage - 1];
    if (pageEdits && editorState.selectedEdit >= 0 && editorState.selectedEdit < pageEdits.length) {
        pageEdits.splice(editorState.selectedEdit, 1);
        editorState.selectedEdit = -1;
        renderEdits();
    }
}

function startDraw(e) {
    // Se há um texto em digitação, confirma antes de tratar o novo clique
    if (editorState.pendingText) commitInlineText();

    const rect = e.target.getBoundingClientRect();
    editorState.startX = e.clientX - rect.left;
    editorState.startY = e.clientY - rect.top;

    // Selecionar: escolhe a edição sob o cursor e prepara drag para mover
    if (editorState.currentTool === 'select') {
        const hitIdx = hitTestEdit(editorState.startX, editorState.startY);
        editorState.selectedEdit = hitIdx;
        if (hitIdx >= 0 && editorState.viewport) {
            const edit = editorState.edits[editorState.currentPage - 1][hitIdx];
            const pdfPt = editorState.viewport.convertToPdfPoint(editorState.startX, editorState.startY);
            editorState.isDragging = true;
            editorState.dragOffsetX = pdfPt[0] - edit.x;
            editorState.dragOffsetY = pdfPt[1] - edit.y;
        }
        renderEdits();
        return;
    }

    // Texto: abre o campo editável no local do clique
    if (editorState.currentTool === 'text') {
        showInlineTextInput(editorState.startX, editorState.startY);
        return;
    }

    // Apagar / Retângulo: inicia o desenho
    editorState.isDrawing = true;
}

function draw(e) {
    // Mover edição selecionada (drag)
    if (editorState.currentTool === 'select' && editorState.isDragging && editorState.selectedEdit >= 0) {
        const rect = e.target.getBoundingClientRect();
        const mx = e.clientX - rect.left;
        const my = e.clientY - rect.top;
        const pdfPt = editorState.viewport.convertToPdfPoint(mx, my);
        const edit = editorState.edits[editorState.currentPage - 1][editorState.selectedEdit];
        edit.x = pdfPt[0] - editorState.dragOffsetX;
        edit.y = pdfPt[1] - editorState.dragOffsetY;
        renderEdits();
        return;
    }
    if (!editorState.isDrawing || editorState.currentTool === 'text') return;

    const rect = e.target.getBoundingClientRect();
    const currentX = e.clientX - rect.left;
    const currentY = e.clientY - rect.top;

    const editCanvas = document.getElementById('editCanvas');
    const ctx = editCanvas.getContext('2d');

    // Redesenhar edições existentes
    renderEdits();

    // Desenhar preview (em coordenadas de tela)
    const color = document.getElementById('toolColor').value;
    ctx.fillStyle = color;
    ctx.globalAlpha = 0.5;

    const x = Math.min(editorState.startX, currentX);
    const y = Math.min(editorState.startY, currentY);
    const width = Math.abs(currentX - editorState.startX);
    const height = Math.abs(currentY - editorState.startY);

    ctx.fillRect(x, y, width, height);
    ctx.globalAlpha = 1;
}

function endDraw(e) {
    if (editorState.currentTool === 'select') {
        editorState.isDragging = false;
        return;
    }
    if (!editorState.isDrawing || editorState.currentTool === 'text') return;

    const rect = e.target.getBoundingClientRect();
    const endX = e.clientX - rect.left;
    const endY = e.clientY - rect.top;

    const dxC = Math.abs(endX - editorState.startX);
    const dyC = Math.abs(endY - editorState.startY);

    if (dxC > 5 && dyC > 5) {
        const color = document.getElementById('toolColor').value;
        // Converter os dois cantos para coordenadas do PDF e normalizar
        const p0 = editorState.viewport.convertToPdfPoint(editorState.startX, editorState.startY);
        const p1 = editorState.viewport.convertToPdfPoint(endX, endY);
        const x = Math.min(p0[0], p1[0]);
        const y = Math.min(p0[1], p1[1]);
        const width = Math.abs(p1[0] - p0[0]);
        const height = Math.abs(p1[1] - p0[1]);
        editorState.edits[editorState.currentPage - 1].push({
            type: editorState.currentTool,
            x, y, width, height, color
        });
    }

    editorState.isDrawing = false;
    renderEdits();
}

function undoEdit() {
    const pageEdits = editorState.edits[editorState.currentPage - 1];
    if (pageEdits && pageEdits.length > 0) {
        pageEdits.pop();
        editorState.selectedEdit = -1;
        renderEdits();
    }
}

function prevPage() {
    if (editorState.currentPage > 1) {
        editorState.currentPage--;
        editorState.selectedEdit = -1;
        renderCurrentPage();
        updatePageIndicator();
    }
}

function nextPage() {
    if (editorState.currentPage < editorState.totalPages) {
        editorState.currentPage++;
        editorState.selectedEdit = -1;
        renderCurrentPage();
        updatePageIndicator();
    }
}

function updatePageIndicator() {
    document.getElementById('pageIndicator').textContent =
        `${editorState.currentPage} / ${editorState.totalPages}`;
    document.getElementById('pageInput').value = editorState.currentPage;
    document.getElementById('pageInput').max = editorState.totalPages;
}

function updateZoomLevel() {
    const percent = Math.round(editorState.scale * 100);
    document.getElementById('zoomLevel').textContent = `${percent}%`;
}

function zoomIn() {
    editorState.scale = Math.min(editorState.scale * 1.25, 5.0);
    renderCurrentPage();
    updateZoomLevel();
}

function zoomOut() {
    editorState.scale = Math.max(editorState.scale / 1.25, 0.25);
    renderCurrentPage();
    updateZoomLevel();
}

function zoomFit() {
    // "Ajustar à tela": cabe a página inteira no contêiner (centralizada).
    const container = document.getElementById('canvasContainer');
    const cw = container.clientWidth - 48;
    const ch = container.clientHeight - 48;
    if (editorState.baseWidth > 0 && editorState.baseHeight > 0 && cw > 0 && ch > 0) {
        const scaleW = cw / editorState.baseWidth;
        const scaleH = ch / editorState.baseHeight;
        editorState.scale = Math.max(0.25, Math.min(scaleW, scaleH, 4.0));
    }
    renderCurrentPage();
    updateZoomLevel();
}

function zoomActual() {
    editorState.scale = 1.0;
    renderCurrentPage();
    updateZoomLevel();
}

function goToPage(pageNum) {
    const page = parseInt(pageNum);
    if (page >= 1 && page <= editorState.totalPages) {
        editorState.currentPage = page;
        editorState.selectedEdit = -1;
        renderCurrentPage();
    }
    // Valor inválido volta para a página atual
    updatePageIndicator();
}

function rotatePage(degrees) {
    const pageIndex = editorState.currentPage - 1;
    editorState.pageRotations[pageIndex] = (editorState.pageRotations[pageIndex] + degrees + 360) % 360;
    renderCurrentPage();
}

function toggleFullscreen() {
    const container = document.getElementById('editorContainer');
    editorState.isFullscreen = !editorState.isFullscreen;
    container.classList.toggle('fullscreen', editorState.isFullscreen);
    // Reajusta o zoom à nova área disponível após o relayout
    requestAnimationFrame(() => zoomFit());
}

// Atalhos: ESC sai da tela cheia; Delete/Backspace apaga a edição selecionada
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && editorState.isFullscreen) {
        toggleFullscreen();
        return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && editorState.pdfDoc && editorState.selectedEdit >= 0) {
        const tag = (document.activeElement && document.activeElement.tagName) || '';
        if (tag !== 'INPUT' && tag !== 'TEXTAREA' && tag !== 'SELECT') {
            e.preventDefault();
            deleteSelected();
        }
    }
});

function cancelEdit() {
    cancelInlineText();
    editorState.selectedEdit = -1;
    // Libera o documento e o worker do pdf.js (evita acumular memória a cada PDF aberto)
    if (editorState.renderTask) editorState.renderTask.cancel();
    if (editorState.pdfDoc) editorState.pdfDoc.destroy().catch(() => {});
    editorState.pdfDoc = null;
    editorState.pdfBytes = null;
    editorState.edits = [];
    editorState.pageRotations = [];
    editorState.viewport = null;
    // Sair do modo tela cheia, se estiver ativo
    editorState.isFullscreen = false;
    document.getElementById('editorContainer').classList.remove('fullscreen');
    document.getElementById('editorContainer').style.display = 'none';
    document.getElementById('editDropZone').style.display = 'block';
    updateEditingLayout();
}

async function saveEditedPdf() {
    try {
        // Confirma texto em digitação antes de salvar
        if (editorState.pendingText) commitInlineText();

        const pdfDoc = await PDFLib.PDFDocument.load(editorState.pdfBytes);
        const pages = pdfDoc.getPages();

        // Cache de fontes embutidas (reaproveita a mesma fonte em todas as páginas)
        const fontCache = {};
        const getFont = async (name) => {
            if (!fontCache[name]) fontCache[name] = await pdfDoc.embedFont(name);
            return fontCache[name];
        };

        for (let i = 0; i < pages.length; i++) {
            const page = pages[i];
            const pageEdits = editorState.edits[i] || [];

            // Rotação total da página (original + extra do usuário)
            const origRotation = page.getRotation().angle;
            const extra = editorState.pageRotations[i] || 0;
            const totalRot = (origRotation + extra) % 360;
            if (extra) {
                page.setRotation(PDFLib.degrees(totalRot));
            }

            // Pré-calcular seno/cosseno uma vez por página (usados no texto)
            const rotRad = totalRot * Math.PI / 180;
            const sinR = Math.sin(rotRad);
            const cosR = Math.cos(rotRad);

            for (const edit of pageEdits) {
                if (edit.type === 'erase' || edit.type === 'rect') {
                    const col = hexToRgb(edit.color);
                    page.drawRectangle({
                        x: edit.x, y: edit.y, width: edit.width, height: edit.height,
                        color: PDFLib.rgb(col.r, col.g, col.b)
                    });
                } else if (edit.type === 'highlight') {
                    const col = hexToRgb(edit.color);
                    page.drawRectangle({
                        x: edit.x, y: edit.y, width: edit.width, height: edit.height,
                        color: PDFLib.rgb(col.r, col.g, col.b), opacity: 0.45
                    });
                } else if (edit.type === 'text') {
                    const col = hexToRgb(edit.color);
                    const font = await getFont(pdfFontName(edit.font, edit.bold, edit.italic));
                    // As fontes padrão do PDF (WinAnsi) não têm emojis e alguns símbolos:
                    // troca o que não for codificável por "?" em vez de falhar o salvamento inteiro
                    const charset = new Set(font.getCharacterSet());
                    const text = Array.from(edit.text)
                        .map(ch => charset.has(ch.codePointAt(0)) ? ch : '?')
                        .join('');
                    const textWidth = font.widthOfTextAtSize(text, edit.size);

                    // edit.y é o topo do texto em coord. PDF. O baseline fica 0.8*size
                    // na direção perpendicular ao fluxo do texto (que varia com a rotação).
                    // Para rotação R: offset perpendicular = (sinR, -cosR) * 0.8 * size
                    const baseX = edit.x + sinR * edit.size * 0.8;
                    const baseY = edit.y - cosR * edit.size * 0.8;

                    // Marca-texto (fundo translúcido), girado junto com o texto: o canto
                    // inferior-esquerdo da caixa fica a "size" do topo, na direção perpendicular
                    if (edit.highlight) {
                        const h = hexToRgb(edit.highlight);
                        page.drawRectangle({
                            x: edit.x + sinR * edit.size, y: edit.y - cosR * edit.size,
                            width: textWidth, height: edit.size,
                            rotate: PDFLib.degrees(totalRot),
                            color: PDFLib.rgb(h.r, h.g, h.b), opacity: 0.45
                        });
                    }
                    // Texto com a rotação correta para aparecer horizontal no viewer
                    page.drawText(text, {
                        x: baseX, y: baseY, size: edit.size, font,
                        color: PDFLib.rgb(col.r, col.g, col.b),
                        rotate: PDFLib.degrees(totalRot)
                    });
                    // Sublinhado (segmento ao longo da direção do texto)
                    if (edit.underline) {
                        const ulOff = edit.size * 0.12;
                        page.drawLine({
                            start: { x: baseX + sinR * ulOff,              y: baseY - cosR * ulOff },
                            end:   { x: baseX + sinR * ulOff + cosR * textWidth,
                                     y: baseY - cosR * ulOff + sinR * textWidth },
                            thickness: Math.max(0.5, edit.size * 0.06),
                            color: PDFLib.rgb(col.r, col.g, col.b)
                        });
                    }
                }
            }
        }

        const editedBytes = await pdfDoc.save();

        // Salvar arquivo (sugere nome baseado no original)
        const suggested = basename(editorState.pdfPath || 'documento.pdf').replace(/\.pdf$/i, '') + '_editado.pdf';
        const outputPath = await window.api.saveFile(suggested);
        if (outputPath) {
            const result = await window.api.savePdf({
                data: bytesToBase64(editedBytes),
                outputPath: outputPath
            });

            if (result.success) {
                alert('PDF salvo com sucesso!');
                cancelEdit();
            } else {
                alert('Erro ao salvar: ' + result.error);
            }
        }
    } catch (e) {
        alert('Erro ao salvar PDF: ' + e.message);
    }
}

// Inicialização
init();
setupEditor();
