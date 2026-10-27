# Web Digital · PDF Tools

Aplicativo desktop (Windows) para **combinar** documentos Office e PDFs em um
único arquivo e **editar** PDFs (cobrir áreas, anotar, adicionar texto e girar páginas).

**Desenvolvido por:** Murillo Matos
**Powered by:** [Web Digital](https://webdigital.app)

## Funcionalidades

### 📑 Combinar
- Junta vários arquivos em um único PDF, na ordem que você definir (arrastar para reordenar).
- Converte documentos Office para PDF automaticamente antes de juntar.
- Ordenação rápida por nome (A→Z / Z→A), por número e por tipo.

### ✏️ Editar PDF
- **Apagar / Retângulo:** cobre áreas do documento com a cor escolhida.
- **Texto:** adiciona anotações de texto.
- Zoom, ajuste à largura, navegação por páginas e **rotação** (a rotação é gravada no PDF salvo).

## Formatos suportados

| Tipo | Extensões |
|------|-----------|
| PDF | `.pdf` |
| Word | `.doc`, `.docx` |
| Excel | `.xls`, `.xlsx` |
| PowerPoint | `.ppt`, `.pptx` |
| Texto | `.rtf`, `.txt` |

## Requisitos

- **Windows 10/11**
- **Microsoft Office** (Word, Excel e/ou PowerPoint) — necessário **apenas** para converter
  documentos Office. A combinação e a edição de **PDFs** funcionam sem o Office.

> A conversão de Office para PDF é feita via automação COM do próprio Microsoft Office,
> preservando a fidelidade do layout original.

## Instalação

### Executar em modo desenvolvimento

```bash
npm install
npm start
```

### Gerar executável

```bash
npm install

# Instalador Windows (NSIS)
npm run build:win

# Versão portátil (sem instalação)
npm run build:portable
```

O executável é gerado na pasta `dist/`.

## Atualização automática (GitHub Releases)

A versão **instalada** (Setup) verifica atualizações ao abrir e a cada 4 horas em
`github.com/MurilloMts/web-digital-pdf-tools` (releases). Quando há versão nova, ela é
baixada em segundo plano e aparece o botão **"Reiniciar e atualizar"** no cabeçalho
(se o usuário apenas fechar o app, a atualização é instalada ao sair).

> A versão **portátil** não se atualiza sozinha — quem usar ela precisa baixar de novo.

### Configuração (uma vez)

1. Crie o repositório **público** `web-digital-pdf-tools` na sua conta do GitHub
   (para outro nome/conta, ajuste `build.publish` no `package.json`).
2. Crie um token em GitHub → Settings → Developer settings → Personal access tokens
   → *Fine-grained*, com acesso ao repositório e permissão **Contents: Read and write**.

### Publicar uma nova versão

1. Aumente `"version"` no `package.json` (ex.: `1.0.1` → `1.0.2`). **Sem isso os apps
   não detectam a atualização.**
2. No PowerShell:

   ```powershell
   $env:GH_TOKEN = "seu_token_aqui"
   npm run release
   ```

3. Pronto: o release é publicado com o instalador, a versão portátil e o `latest.yml`.
   Os apps instalados se atualizam na próxima vez que forem abertos.

> Mande para os amigos o link `https://github.com/MurilloMts/web-digital-pdf-tools/releases/latest`
> — eles instalam o **Setup** uma vez e daí em diante tudo é automático.

## Como usar

### Combinar
1. Abra a aba **Combinar**.
2. Arraste os arquivos (ou clique para selecionar).
3. Reordene arrastando e, se quiser, use os botões de ordenação.
4. Clique em **Combinar e Salvar** e escolha onde salvar o PDF final.

### Editar PDF
1. Abra a aba **Editar PDF**.
2. Arraste um PDF (ou clique para selecionar).
3. Escolha uma ferramenta (Apagar, Texto ou Retângulo), ajuste cor/tamanho e edite.
4. Clique em **Salvar PDF Editado**.

## Tecnologias

- [Electron](https://www.electronjs.org/) (com `contextIsolation` + `preload`)
- Automação COM do Microsoft Office (conversão Office → PDF)
- [pdf-lib](https://pdf-lib.js.org/) (combinação e gravação de edições)
- [pdf.js](https://mozilla.github.io/pdf.js/) (renderização para a edição)

> As bibliotecas pdf-lib e pdf.js são empacotadas localmente em `lib/`, então o
> aplicativo **funciona offline**.

## Estrutura do projeto

```
├── main.js          # Processo principal (janela, IPC, conversão via COM)
├── preload.js       # Ponte segura (contextBridge) entre main e interface
├── renderer.js      # Lógica da interface (combinar + editor)
├── index.html       # Interface
├── styles.css       # Estilos
├── lib/             # pdf-lib e pdf.js empacotados localmente
├── assets/          # Ícones (icon.svg, icon.png, icon.ico)
└── tools/           # Utilitários de build (gerador de ícones: npm run gen:icons)
```

## Licença

MIT License — livre para uso pessoal e comercial.

---

© 2025 Murillo Matos · Web Digital
