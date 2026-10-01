// Convertit les EMF/WMF d'un dossier media en PNG « <fichier>.png » via convert-metafiles.ps1 (.NET/WIC).
// Windows uniquement ; sur autre OS ou en cas d'échec, on ignore (les EMF resteront non rendus).
import { existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const HERE = dirname(fileURLToPath(import.meta.url));

export function convertMetafiles(mediaDir) {
  if (process.platform !== 'win32' || !existsSync(mediaDir)) return;
  try {
    execFileSync('powershell', ['-ExecutionPolicy', 'Bypass', '-File', join(HERE, 'convert-metafiles.ps1'), mediaDir],
      { stdio: 'ignore', timeout: 120000 });
  } catch { /* pas de .NET/Windows -> EMF non convertis, non bloquant */ }
}
