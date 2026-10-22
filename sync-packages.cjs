const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execSync } = require('child_process');

const projectRoot = 'C:\\Users\\Xeon\\.gemini\\antigravity\\scratch\\echo-powerpoint-theme';
const contentDir = path.join(projectRoot, 'content');
const nodeExe = 'E:\\Steam\\steamapps\\common\\ECHO\\ShinawaseLoader\\node.exe';

// 1. Ensure root mod.js and content/mod.js are 100% identical before computing manifest hashes
fs.copyFileSync(path.join(projectRoot, 'mod.js'), path.join(contentDir, 'mod.js'));

execSync(`"${nodeExe}" .echo-sdk/bin/echo-workshop-sdk.mjs sync .`, { cwd: projectRoot, stdio: 'inherit' });

function copyDirRecursive(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const s = path.join(src, entry.name);
    const d = path.join(dest, entry.name);
    if (entry.isDirectory()) copyDirRecursive(s, d);
    else fs.copyFileSync(s, d);
  }
}

const manifestBuf = fs.readFileSync(path.join(contentDir, 'echo.workshop.json'));
const theme = JSON.parse(fs.readFileSync(path.join(contentDir, 'theme.json'), 'utf8'));
const manifestSha256 = crypto.createHash('sha256').update(manifestBuf).digest('hex');

const storageKey = crypto.createHash('sha256').update('local\0echo.powerpoint-theme', 'utf8').digest('hex').slice(0, 32);
const installedRoot = 'C:\\Users\\Xeon\\AppData\\Roaming\\ECHO Steam\\workshop';
const itemVersionDir = path.join(installedRoot, 'installed', storageKey, 'echo.powerpoint-theme', '1.0.0');
const destDir = path.join(itemVersionDir, manifestSha256);

copyDirRecursive(contentDir, destDir);
if (fs.existsSync(itemVersionDir)) {
  for (const entry of fs.readdirSync(itemVersionDir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      copyDirRecursive(contentDir, path.join(itemVersionDir, entry.name));
    }
  }
}

const hostModsDir = 'E:\\Steam\\steamapps\\common\\ECHO\\Mods\\installed\\echo.powerpoint-theme';
if (fs.existsSync(hostModsDir)) {
  fs.rmSync(hostModsDir, { recursive: true, force: true });
}
fs.mkdirSync(hostModsDir, { recursive: true });
fs.copyFileSync(path.join(projectRoot, 'echo.mod.json'), path.join(hostModsDir, 'echo.mod.json'));
fs.copyFileSync(path.join(projectRoot, 'mod.js'), path.join(hostModsDir, 'mod.js'));
copyDirRecursive(contentDir, path.join(hostModsDir, 'content'));

// Clean temporary debug/test files in projectRoot before syncing to Desktop
const allowedRootEntries = new Set([
  '.echo-sdk',
  '.git',
  '.gitignore',
  'CHANGELOG.md',
  'LICENSE',
  'MEMO.md',
  'content',
  'dist',
  'echo.mod.json',
  'echo.workshop.project.json',
  'mod.js',
  'package.json',
  'preview.png',
  'README.md',
  'scan-emoji.cjs',
  'sync-packages.cjs'
]);
for (const entry of fs.readdirSync(projectRoot, { withFileTypes: true })) {
  if (!allowedRootEntries.has(entry.name)) {
    fs.rmSync(path.join(projectRoot, entry.name), { recursive: true, force: true });
  }
}

const desktopFolder = 'C:\\Users\\Xeon\\Desktop\\echo.powerpoint-theme';
if (fs.existsSync(desktopFolder)) {
  fs.rmSync(desktopFolder, { recursive: true, force: true });
}
copyDirRecursive(projectRoot, desktopFolder);

const desktopPackageFolder = 'C:\\Users\\Xeon\\Desktop\\echo.powerpoint-theme-package';
if (fs.existsSync(desktopPackageFolder)) fs.rmSync(desktopPackageFolder, { recursive: true, force: true });
copyDirRecursive(contentDir, desktopPackageFolder);

const desktopZip = 'C:\\Users\\Xeon\\Desktop\\echo.powerpoint-theme-1.4.1.zip';
const desktopZipOld = 'C:\\Users\\Xeon\\Desktop\\echo.powerpoint-theme-1.0.0.zip';
if (fs.existsSync(desktopZip)) fs.unlinkSync(desktopZip);
execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${contentDir}\\*' -DestinationPath '${desktopZip}' -Force"`);
if (fs.existsSync(desktopZipOld)) fs.unlinkSync(desktopZipOld);
fs.copyFileSync(desktopZip, desktopZipOld);

const nowIso = new Date().toISOString();
const registryPath = path.join(installedRoot, 'registry.json');
const registry = JSON.parse(fs.readFileSync(registryPath, 'utf8'));
const revision = { contentId: 'echo.powerpoint-theme', contentKind: 'theme', version: '1.0.0', manifestSha256, directory: destDir, installedAt: nowIso };
const recIdx = registry.records.findIndex(r => r.sourceId === 'local' && r.itemId === 'echo.powerpoint-theme');
const rec = { sourceId: 'local', itemId: 'echo.powerpoint-theme', state: 'enabled', candidateRevision: revision, activeRevision: revision, lastKnownGoodRevision: null, approvedCapabilities: theme.runtime?.capabilities || ['playback:read'], error: null, createdAt: nowIso, updatedAt: nowIso };
if (recIdx >= 0) registry.records[recIdx] = rec; else registry.records.push(rec);
registry.revision = (registry.revision || 1) + 1;
fs.writeFileSync(registryPath, JSON.stringify(registry, null, 2));

const catalogPath = path.join(installedRoot, 'data-catalog.json');
const catalog = JSON.parse(fs.readFileSync(catalogPath, 'utf8'));
const catIdx = catalog.records.findIndex(r => r.sourceId === 'local' && r.itemId === 'echo.powerpoint-theme');
const catRec = { sourceId: 'local', itemId: 'echo.powerpoint-theme', contentId: 'echo.powerpoint-theme', contentKind: 'theme', version: '1.0.0', manifestSha256, entryPath: 'theme.json', contribution: theme, activatedAt: nowIso };
if (catIdx >= 0) catalog.records[catIdx] = catRec; else catalog.records.push(catRec);
catalog.revision = (catalog.revision || 1) + 1;
fs.writeFileSync(catalogPath, JSON.stringify(catalog, null, 2));

const activeReceipt = { formatVersion: 1, receipt: { sourceId: 'local', itemId: 'echo.powerpoint-theme', contentId: 'echo.powerpoint-theme', version: '1.0.0', manifestSha256, registryUpdatedAt: nowIso } };
fs.writeFileSync(path.join(installedRoot, 'active-theme-background.json'), JSON.stringify(activeReceipt, null, 2));
fs.writeFileSync(path.join(installedRoot, 'active-theme-background.json.lyrics.json'), JSON.stringify(activeReceipt, null, 2));

const settingsPath = 'C:\\Users\\Xeon\\AppData\\Roaming\\ECHO Steam\\echo-settings.json';
if (fs.existsSync(settingsPath)) {
  try {
    const settings = JSON.parse(fs.readFileSync(settingsPath, 'utf8'));
    settings.appearanceThemeCustomId = 'workshop:714422596da03d89810e';
    settings.lyricsWorkshopBackgroundEnabled = true;
    settings.lyricsBackgroundMode = 'theme';
    fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2));
  } catch (err) {
    console.error('Failed to update echo-settings.json:', err);
  }
}
console.log('Synced all targets with new SHA256:', manifestSha256);
