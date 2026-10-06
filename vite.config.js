import { defineConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import topLevelAwait from 'vite-plugin-top-level-await';
import fs from 'node:fs';
import path from 'node:path';

// Lists the test documents in the git-ignored examples/ folder at request time, so new files
// show up on the welcome screen with a page reload. Dev server only: production builds have
// no such endpoint and never bundle or link private documents.
function exampleDocuments() {
  return {
    name: 'example-documents',
    apply: 'serve',
    configureServer(server) {
      server.middlewares.use('/__examples.json', (req, res) => {
        const dir = path.resolve(__dirname, 'examples');
        const files = fs.existsSync(dir)
          ? fs.readdirSync(dir)
            .filter(name => /\.docx$/i.test(name) && !name.startsWith('~$')) // skip Word lock files
            .sort((a, b) => a.localeCompare(b))
            .map(name => {
              const stat = fs.statSync(path.join(dir, name));
              return { name, size: stat.size, modified: stat.mtimeMs };
            })
          : [];
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Cache-Control', 'no-store');
        res.end(JSON.stringify({ folder: 'examples', files }));
      });
    },
  };
}

export default defineConfig({
  plugins: [
    wasm(),
    topLevelAwait(),
    exampleDocuments()
  ],
  server: {
    port: 5173,
    open: false
  }
});
