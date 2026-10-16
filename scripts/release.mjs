#!/usr/bin/env node
// Publishes a new version of @dimarborda/docx-editor to npm.
//
//   npm run release -- <patch|minor|major|x.y.z> [--dry-run] [--otp=123456]
//
// 1. Checks: on master, no pending changes (except .claude/), up to date with origin, logged
//    in to npm and the version not published yet.
// 2. Rust tests, wasm build, JS tests and npm audit, as CI does.
// 3. Checks the package contents (engine .wasm, sources, types).
// 4. Bumps the version (package.json + package-lock.json) and publishes. If publishing fails
//    the version bump is undone.
// 5. Commits "chore(release): x.y.z", tags vx.y.z and pushes both.
//
// --dry-run runs steps 1–3 and `npm publish --dry-run`, changing nothing (pending changes are
// only a warning there, to check a release before committing).
import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkgDir = 'packages/docx-editor';
const pkgPath = path.join(root, pkgDir, 'package.json');
const BRANCH = 'master';
/** Files the published package cannot work without */
const REQUIRED_FILES = ['wasm/docx_engine_bg.wasm', 'wasm/docx_engine.js', 'src/index.js', 'src/index.d.ts', 'src/style.css', 'README.md', 'LICENSE'];

const args = process.argv.slice(2);
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
    default: fail(`Indica patch, minor, major o una versión x.y.z (recibido: ${kind ?? 'nada'}).\n  Uso: npm run release -- minor [--dry-run] [--otp=123456]`);
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
if (read('git', ['tag', '--list', `v${version}`])) fail(`La etiqueta v${version} ya existe.`);

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
run('cargo', ['test', '--quiet']);
step('Compilando el motor WebAssembly');
run('npm', ['run', 'build:wasm']);
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
run('git', ['commit', '--quiet', '-m', `chore(release): ${version}`]);
run('git', ['tag', '-a', `v${version}`, '-m', `${pkg.name} ${version}`]);
if (!run('git', ['push', '--follow-tags', 'origin', BRANCH], { allowFail: true })) {
  console.error(`\n⚠ Publicado en npm, pero no se pudo subir a GitHub. Repite: git push --follow-tags origin ${BRANCH}`);
  process.exit(1);
}
console.log(`\n✔ ${pkg.name}@${version} publicado: https://www.npmjs.com/package/${pkg.name}/v/${version}`);
