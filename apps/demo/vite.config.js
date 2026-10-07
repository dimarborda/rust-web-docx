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

export default defineConfig({
  plugins: [exampleDocuments()],
  server: {
    port: 5173,
    open: false,
  },
  build: {
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
        embed: path.resolve(__dirname, 'embed.html'),
      },
    },
  },
});
