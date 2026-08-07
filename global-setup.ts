import { execSync } from 'child_process';

/**
 * Runs before every test session.
 * Kills ONLY Playwright-launched Chrome instances (identified by --remote-debugging-port
 * in their command line) so user's own Chrome windows are never touched.
 */
export default async function globalSetup() {
  console.log('[globalSetup] Checking for leftover Playwright Chrome processes...');

  try {
    // WMIC finds only Chrome processes whose command line contains --remote-debugging-port
    // (Playwright always passes this flag; regular user Chrome windows do not)
    const result = execSync(
      'wmic process where "name=\'chrome.exe\' and commandline like \'%--remote-debugging-port%\'" get processid /format:value 2>nul',
      { encoding: 'utf8' }
    );

    const pids = result
      .split('\n')
      .map(l => l.trim())
      .filter(l => l.startsWith('ProcessId='))
      .map(l => l.replace('ProcessId=', '').trim())
      .filter(p => p.length > 0);

    if (pids.length === 0) {
      console.log('[globalSetup] No leftover Playwright Chrome found — clean start.');
    } else {
      console.log(`[globalSetup] Killing ${pids.length} leftover Playwright Chrome process(es): ${pids.join(', ')}`);
      for (const pid of pids) {
        try {
          execSync(`taskkill /F /PID ${pid} /T 2>nul`, { stdio: 'ignore' });
        } catch { /* already gone */ }
      }
      console.log('[globalSetup] Done — user Chrome windows are untouched.');
    }
  } catch {
    // WMIC returned nothing or failed — no Playwright Chrome running
    console.log('[globalSetup] No leftover Playwright Chrome found — clean start.');
  }
}
