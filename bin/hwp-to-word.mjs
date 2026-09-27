#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const metadata = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const help = `Usage: hwp-to-word <source.hwp> <output.docx> [options]

Convert an unprotected HWP 3 or HWP 5 document to editable Word.
HWP 5 supports multiple sections with per-section page sizes and margins.
Web/email/FTP links and section-initial headers/footers are supported subsets.
HWP 3 support is an experimental subset: ASCII/modern Hangul, basic styles, tables.
Requires Node.js 18+ and Python 3.9+ on macOS, Linux, or Windows.
First conversion downloads Python dependencies into a private user cache.
Document conversion runs locally. The parser is not sandboxed.
Windows runtime setup is implemented; actual Windows conversion is unverified.

Options:
  --font-map <path>          JSON map of original to installed font families
  --preserve-source-pages   Infer page breaks for single-column HWP 5 documents
  --recover-text            Replace invalid HWP 5 UTF-16 with U+FFFD and report it
  --max-input-mb <MiB>       Input size budget (default: 50; HWP 3 ceiling: 50)
  --max-stream-mb <MiB>      HWP 5 stored/expanded stream budget (default: 64)
  --max-expanded-mb <MiB>    HWP 5 total expanded data budget (default: 128)
  --max-records <count>      HWP 5 records per stream (default: 200000)
  --reader-timeout <seconds> Reader time budget (default: 60)
  --help                    Show this help without installing dependencies
  --version                 Show package version

HWPX, protected files, WMF images, and exact layout fidelity are unsupported.
HWP 3 images, legacy Hanja/symbols/old Hangul, and complex controls are unsupported.
Uses this project's native HWP parser. Project license: AGPL-3.0-or-later.
본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
`;

export function invoke(command, args, options = {}) {
  const result = spawnSync(command, args, { stdio: 'inherit', ...options });
  if (result.error) throw new Error(`Cannot run ${command}: ${result.error.message}`);
  if (result.signal) throw new Error(`${command} stopped by ${result.signal}`);
  return result.status ?? 1;
}

export function findPython(run = spawnSync, environment = process.env) {
  const candidates = environment.HWP_TO_WORD_PYTHON
    ? [environment.HWP_TO_WORD_PYTHON] : ['python3', 'python'];
  for (const command of candidates) {
    const result = run(command, ['-c', 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)'], {
      stdio: 'ignore',
    });
    if (!result.error && result.status === 0) return command;
  }
  throw new Error('Python 3.9+ is required. Install Python or set HWP_TO_WORD_PYTHON to its executable path.');
}

export function ensureRuntime({ cacheRoot, python, run = invoke, platform = process.platform }) {
  const requirements = join(root, 'requirements.txt');
  const digest = createHash('sha256').update(readFileSync(requirements)).digest('hex').slice(0, 16);
  const directory = join(cacheRoot, `${metadata.version}-${platform}-${process.arch}-${digest}`);
  const executable = platform === 'win32'
    ? join(directory, 'venv', 'Scripts', 'python.exe')
    : join(directory, 'venv', 'bin', 'python');
  const marker = join(directory, 'ready');
  if (existsSync(marker) && existsSync(executable)) return executable;
  mkdirSync(directory, { recursive: true });
  const lock = join(directory, 'install.lock');
  try {
    mkdirSync(lock);
  } catch (error) {
    if (error.code === 'EEXIST') {
      throw new Error(`Runtime setup is locked at ${lock}. Wait for the other setup to finish. If it was interrupted, remove this lock directory after confirming no setup is running.`);
    }
    throw error;
  }
  try {
    console.error('Preparing the local Python runtime (first use requires internet access)...');
    if (run(python, ['-m', 'venv', join(directory, 'venv')]) !== 0) {
      throw new Error('Could not create a Python virtual environment. Check that Python venv support is installed.');
    }
    if (run(executable, ['-m', 'pip', 'install', '--disable-pip-version-check', '-r', requirements]) !== 0) {
      throw new Error('Python dependency installation failed. Check network access, then retry.');
    }
    writeFileSync(marker, `${metadata.version}\n`);
    return executable;
  } finally {
    rmSync(lock, { recursive: true, force: true });
  }
}

export function main(args = process.argv.slice(2)) {
  if (args.includes('--help') || args.includes('-h') || args.length === 0) {
    console.log(help);
    return 0;
  }
  if (args.length === 1 && args[0] === '--version') {
    console.log(metadata.version);
    return 0;
  }
  const cacheRoot = process.env.HWP_TO_WORD_CACHE
    || join(process.env.XDG_CACHE_HOME || join(homedir(), '.cache'), 'hwp-to-word');
  const executable = ensureRuntime({ cacheRoot, python: findPython() });
  return invoke(executable, [join(root, 'convert_hwp.py'), ...args]);
}

if (process.argv[1] && import.meta.url === pathToFileURL(realpathSync(process.argv[1])).href) {
  try {
    process.exitCode = main();
  } catch (error) {
    console.error(`hwp-to-word: ${error.message}`);
    process.exitCode = 1;
  }
}
