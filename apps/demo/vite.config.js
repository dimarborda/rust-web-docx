import { defineConfig } from 'vite';
import fs from 'node:fs';
import path from 'node:path';

const EXAMPLES_DIR = path.resolve(__dirname, '../../examples');
const isDocx = name => /\.docx$/i.test(name) && !name.startsWith('~$'); // skip Word lock files

// Lists and serves the test documents in the git-ignored examples/ folder at request time, so
// new files show up on the welcome screen with a page reload. Dev server only: production
// builds have no such endpoints and never bundle or link private documents.
function exampleDocuments() {
  return {
    name: 'example-documents',
    apply: 'serve',
    configureServer(server) {
      server.middlewares.use('/__examples.json', (req, res) => {
        const files = fs.existsSync(EXAMPLES_DIR)
          ? fs.readdirSync(EXAMPLES_DIR)
            .filter(isDocx)
            .sort((a, b) => a.localeCompare(b))
            .map(name => {
              const stat = fs.statSync(path.join(EXAMPLES_DIR, name));
              return { name, size: stat.size, modified: stat.mtimeMs };
            })
          : [];
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Cache-Control', 'no-store');
        res.end(JSON.stringify({ folder: 'examples', files }));
      });
      server.middlewares.use('/__examples/', (req, res, next) => {
        const name = path.basename(decodeURIComponent(req.url.split('?')[0]));
        const file = path.join(EXAMPLES_DIR, name);
        if (!isDocx(name) || !fs.existsSync(file)) return next();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        fs.createReadStream(file).pipe(res);
      });
    },
  };
}

const PACKAGES_DIR = path.resolve(__dirname, '../../packages');
const SITE = 'https://rust-web-docx.dimarborda.workers.dev';
const REPO = 'https://github.com/dimarborda/rust-web-docx/blob/master';

// Publishes /llms.txt (index) and /llms-full.txt (both AGENTS.md guides) with the build, generated
// from the packages' AGENTS.md so the site and the npm packages never drift apart.
function llmsTxt() {
  return {
    name: 'llms-txt',
    apply: 'build',
    generateBundle() {
      const guide = name => fs.readFileSync(path.join(PACKAGES_DIR, name, 'AGENTS.md'), 'utf8').trim();
      const index = `# rust-web-docx

> Two MIT-licensed npm packages that edit Office files losslessly with a Rust + WebAssembly engine, in the browser (canvas editors, framework-agnostic web components) and, for Excel, also headless in Node.js. Built for AI agents to edit documents while a person keeps working. No server: files never leave the page.

- \`@dimarborda/docx-editor\`: Word (.docx) editor for the browser. Templates and markers, logical sections, tables, images, PDF export, agent activity marks.
- \`@dimarborda/xlsx-editor\`: Excel (.xlsx) editor for the browser and headless engine (\`/engine\`) for Node.js. Formulas, formats, row/column operations, area context for agents.

Each package ships an AGENTS.md next to its package.json; import it from your project's AGENTS.md or CLAUDE.md with \`@node_modules/@dimarborda/<package>/AGENTS.md\`.

## Agent guides

- [docx-editor AGENTS.md](${REPO}/packages/docx-editor/AGENTS.md): API map, conventions and pitfalls
- [xlsx-editor AGENTS.md](${REPO}/packages/xlsx-editor/AGENTS.md): API map, conventions and pitfalls
- [Both guides in one file](${SITE}/llms-full.txt)

## Docs

- [docx-editor API](${REPO}/docs/docx-editor/en/api.md)
- [docx-editor guides](${REPO}/docs/docx-editor/en/guides.md): markers, sections, tables, PDF, images
- [docx-editor integrations](${REPO}/docs/docx-editor/en/integrations.md): React, Next.js, Tauri, CSP
- [xlsx-editor component](${REPO}/docs/xlsx-editor/en/component.md)
- [xlsx-editor engine](${REPO}/docs/xlsx-editor/en/engine.md)
- [xlsx-editor areas for agents](${REPO}/docs/xlsx-editor/en/agents.md)

## Optional

- [Demo](${SITE}/?demo)
- [Security policy](${REPO}/SECURITY.md)
`;
      const full = `# rust-web-docx: full agent guides\n\n${guide('docx-editor')}\n\n---\n\n${guide('xlsx-editor')}\n`;
      this.emitFile({ type: 'asset', fileName: 'llms.txt', source: index });
      this.emitFile({ type: 'asset', fileName: 'llms-full.txt', source: full });
    },
  };
}

export default defineConfig({
  plugins: [exampleDocuments(), llmsTxt()],
  server: {
    port: 5173,
    open: false,
  },
  build: {
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
        embed: path.resolve(__dirname, 'embed.html'),
        xlsx: path.resolve(__dirname, 'xlsx.html'),
      },
    },
  },
});
