#!/usr/bin/env node
// Publishes a new version of one of the monorepo's npm packages.
//
//   npm run release -- <patch|minor|major|x.y.z> [--pkg=docx|xlsx|xlsx-editor] [--dry-run] [--otp=123456]
//
// --pkg picks the package: docx (default, @dimarborda/docx-editor, tags vx.y.z), xlsx
// (@dimarborda/xlsx, tags xlsx-vx.y.z) or xlsx-editor (@dimarborda/xlsx-editor, tags
// xlsx-editor-vx.y.z). Each package is versioned and published on its own.
//
// 1. Checks: on master, no pending changes (except .claude/), up to date with origin, logged
//    in to npm and the version not published yet.
// 2. Rust tests, wasm build, JS tests and npm audit, as CI does.
// 3. Checks the package contents (engine .wasm, sources, types).
// 4. Bumps the version (package.json + package-lock.json) and publishes. If publishing fails
//    the version bump is undone.
// 5. Commits "chore(release): x.y.z", tags it (vx.y.z, or xlsx-vx.y.z for xlsx) and pushes both.
//
// --dry-run runs steps 1–3 and `npm publish --dry-run`, changing nothing (pending changes are
// only a warning there, to check a release before committing).
import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BRANCH = 'master';
/** The publishable packages: directory, wasm build script, tag prefix and the files the
 * published package cannot work without */
const PACKAGES = {
  docx: {
    dir: 'packages/docx-editor',
    buildScript: 'build:wasm',
    tagPrefix: 'v',
    required: ['wasm/docx_engine_bg.wasm', 'wasm/docx_engine.js', 'src/index.js', 'src/index.d.ts', 'src/style.css', 'README.md', 'LICENSE'],
  },
  xlsx: {
    dir: 'packages/xlsx',
    buildScript: 'build:wasm:xlsx',
    tagPrefix: 'xlsx-v',
    required: ['wasm/xlsx_engine_bg.wasm', 'wasm/xlsx_engine.js', 'src/index.js', 'src/index.d.ts', 'README.md', 'LICENSE'],
  },
  // No engine of its own: it depends on the published @dimarborda/xlsx, so release that first
  'xlsx-editor': {
    dir: 'packages/xlsx-editor',
    buildScript: 'build:wasm:xlsx',
    tagPrefix: 'xlsx-editor-v',
    required: ['src/index.js', 'src/index.d.ts', 'src/xlsx_editor.js', 'src/style.css', 'src/fonts.js', 'README.md', 'LICENSE'],
  },
};

const args = process.argv.slice(2);
const pkgKey = args.find(a => a.startsWith('--pkg='))?.slice('--pkg='.length) ?? 'docx';
const target = PACKAGES[pkgKey];
if (!target) {
  console.error(`\n✖ Paquete desconocido: ${pkgKey}. Usa --pkg=${Object.keys(PACKAGES).join(' o --pkg=')}.`);
  process.exit(1);
}
const pkgDir = target.dir;
const pkgPath = path.join(root, pkgDir, 'package.json');
const REQUIRED_FILES = target.required;
const tag = v => `${target.tagPrefix}${v}`;
const dryRun = args.includes('--dry-run');
const otp = args.find(a => a.startsWith('--otp='))?.slice('--otp='.length);
const bump = args.find(a => !a.startsWith('--'));

const step = msg => console.log(`\n▸ ${msg}`);
const fail = msg => {
  console.error(`\n✖ ${msg}`);
  process.exit(1);
};

/** Runs a command showing its output; stops the release when it fails */
function run(cmd, cmdArgs, { allowFail = false } = {}) {
  console.log(`  $ ${cmd} ${cmdArgs.join(' ')}`);
  const r = spawnSync(cmd, cmdArgs, { cwd: root, stdio: 'inherit' });
  if (r.status !== 0 && !allowFail) fail(`Falló: ${cmd} ${cmdArgs.join(' ')}`);
  return r.status === 0;
}

/** Runs a command and returns its trimmed stdout ('' when it fails and `allowFail`) */
function read(cmd, cmdArgs, { allowFail = false } = {}) {
  try {
    return execFileSync(cmd, cmdArgs, { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }).trim();
  } catch (err) {
    if (allowFail) return '';
    fail(`Falló: ${cmd} ${cmdArgs.join(' ')}\n${err.stderr || err.message}`);
  }
}

function nextVersion(current, kind) {
  if (/^\d+\.\d+\.\d+(-[0-9A-Za-z.-]+)?$/.test(kind)) return kind;
  const [major, minor, patch] = current.split('-')[0].split('.').map(Number);
  switch (kind) {
    case 'major': return `${major + 1}.0.0`;
    case 'minor': return `${major}.${minor + 1}.0`;
    case 'patch': return `${major}.${minor}.${patch + 1}`;
    default: fail(`Indica patch, minor, major o una versión x.y.z (recibido: ${kind ?? 'nada'}).\n  Uso: npm run release -- minor [--pkg=docx|xlsx|xlsx-editor] [--dry-run] [--otp=123456]`);
  }
}

const isNewer = (a, b) => {
  const [x, y] = [a, b].map(v => v.split('-')[0].split('.').map(Number));
  for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] > y[i];
  return false;
};

// ---------- 1. Preconditions ----------
const pkgText = readFileSync(pkgPath, 'utf8');
const pkg = JSON.parse(pkgText);
const version = nextVersion(pkg.version, bump);
if (!isNewer(version, pkg.version)) fail(`La versión ${version} no es mayor que la actual (${pkg.version}).`);
console.log(`${pkg.name}: ${pkg.version} → ${version}${dryRun ? '  (ensayo, no se publica nada)' : ''}`);

step('Comprobando el repositorio');
const branch = read('git', ['rev-parse', '--abbrev-ref', 'HEAD']);
if (branch !== BRANCH) fail(`Estás en la rama ${branch}; publica desde ${BRANCH}.`);
const pending = read('git', ['status', '--porcelain'])
  .split('\n')
  .filter(Boolean)
  // The status code is one or two columns (read() trims the first line's leading space)
  .filter(line => !line.replace(/^\S+\s+/, '').startsWith('.claude/'));
if (pending.length && dryRun) console.log(`  ⚠ Hay cambios sin confirmar (en el ensayo se permiten):\n${pending.join('\n')}`);
else if (pending.length) fail(`Hay cambios sin confirmar:\n${pending.join('\n')}`);
read('git', ['fetch', '--quiet', 'origin', BRANCH]);
const [behind, ahead] = read('git', ['rev-list', '--left-right', '--count', `origin/${BRANCH}...HEAD`]).split(/\s+/).map(Number);
if (behind) fail(`Tu ${BRANCH} está ${behind} commit(s) por detrás de origin; haz git pull primero.`);
if (ahead) console.log(`  (hay ${ahead} commit(s) sin subir; se subirán con la versión)`);
if (read('git', ['tag', '--list', tag(version)])) fail(`La etiqueta ${tag(version)} ya existe.`);

step('Comprobando npm');
const user = read('npm', ['whoami'], { allowFail: true });
if (user) console.log(`  sesión de npm: ${user}`);
else if (dryRun) console.log('  ⚠ No has iniciado sesión en npm (npm login); en el ensayo no hace falta.');
else fail('No has iniciado sesión en npm. Ejecuta: npm login');
if (read('npm', ['view', `${pkg.name}@${version}`, 'version'], { allowFail: true })) {
  fail(`${pkg.name}@${version} ya está publicado en npm.`);
}

// ---------- 2. Tests and build ----------
step('Pruebas de Rust');
run('cargo', ['test', '--workspace', '--quiet']);
step('Compilando el motor WebAssembly');
run('npm', ['run', target.buildScript]);
step('Pruebas de JavaScript');
run('npm', ['test', '-w', pkgDir]);
step('Auditoría de dependencias npm');
run('npm', ['audit', '--omit=dev', '--audit-level=high', '-w', pkgDir]);

// ---------- 3. Package contents ----------
step('Revisando el contenido del paquete');
const [packed] = JSON.parse(read('npm', ['pack', '--dry-run', '--json', '-w', pkgDir]));
const files = new Set(packed.files.map(f => f.path));
const missing = REQUIRED_FILES.filter(f => !files.has(f));
if (missing.length) fail(`Al paquete le faltan: ${missing.join(', ')}`);
const stray = [...files].filter(f => f.startsWith('test/'));
if (stray.length) fail(`El paquete incluiría archivos de prueba: ${stray.join(', ')}`);
console.log(`  ${packed.entryCount} archivos, ${(packed.size / 1024 / 1024).toFixed(2)} MB comprimido, ${(packed.unpackedSize / 1024 / 1024).toFixed(2)} MB descomprimido`);

if (dryRun) {
  step('Ensayo de publicación');
  run('npm', ['publish', '--dry-run', '--access', 'public', '-w', pkgDir]);
  console.log(`\n✔ Todo listo para publicar ${pkg.name}@${version}. Ejecuta sin --dry-run para hacerlo.`);
  process.exit(0);
}

// ---------- 4. Version bump and publish ----------
step(`Subiendo la versión a ${version}`);
const lockPath = path.join(root, 'package-lock.json');
const lockText = readFileSync(lockPath, 'utf8');
const restore = () => {
  writeFileSync(pkgPath, pkgText);
  writeFileSync(lockPath, lockText);
};
writeFileSync(pkgPath, pkgText.replace(`"version": "${pkg.version}"`, `"version": "${version}"`));
if (JSON.parse(readFileSync(pkgPath, 'utf8')).version !== version) {
  restore();
  fail('No se pudo escribir la versión en package.json.');
}
if (!run('npm', ['install', '--package-lock-only', '--ignore-scripts', '--no-audit', '--no-fund'], { allowFail: true })) {
  restore();
  fail('No se pudo actualizar package-lock.json; no se cambió nada.');
}

step('Publicando en npm');
const publishArgs = ['publish', '--access', 'public', '-w', pkgDir, ...(otp ? [`--otp=${otp}`] : [])];
if (!run('npm', publishArgs, { allowFail: true })) {
  restore();
  fail('npm publish falló; se deshizo el cambio de versión. Si pide código de verificación, repite con --otp=123456.');
}

// ---------- 5. Commit, tag and push ----------
step('Confirmando y etiquetando la versión');
run('git', ['add', pkgPath, lockPath]);
run('git', ['commit', '--quiet', '-m', pkgKey === 'docx' ? `chore(release): ${version}` : `chore(release): ${pkg.name} ${version}`]);
run('git', ['tag', '-a', tag(version), '-m', `${pkg.name} ${version}`]);
if (!run('git', ['push', '--follow-tags', 'origin', BRANCH], { allowFail: true })) {
  console.error(`\n⚠ Publicado en npm, pero no se pudo subir a GitHub. Repite: git push --follow-tags origin ${BRANCH}`);
  process.exit(1);
}
console.log(`\n✔ ${pkg.name}@${version} publicado: https://www.npmjs.com/package/${pkg.name}/v/${version}`);
