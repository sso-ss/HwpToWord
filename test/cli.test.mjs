import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, existsSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';
import { ensureRuntime, findPython, invoke } from '../bin/hwp-to-word.mjs';

const { version } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));

test('help and version do not require Python', () => {
  for (const flag of ['--help', '--version']) {
    const result = spawnSync(process.execPath, ['bin/hwp-to-word.mjs', flag], {
      encoding: 'utf8', env: { ...process.env, HWP_TO_WORD_PYTHON: '/missing/python' },
    });
    assert.equal(result.status, 0, result.stderr);
    if (flag === '--help') {
      assert.match(result.stdout, /Usage: hwp-to-word/);
      assert.match(result.stdout, /multiple sections/);
      for (const option of ['--max-input-mb', '--max-stream-mb', '--max-expanded-mb', '--max-records', '--reader-timeout']) {
        assert.ok(result.stdout.includes(option), `${option} should be documented`);
      }
    }
    else assert.equal(result.stdout.trim(), version);
  }
});

test('Python discovery falls back and honors explicit executable paths', () => {
  const commands = [];
  const run = command => {
    commands.push(command);
    return { status: command === 'python3' ? 1 : 0 };
  };
  assert.equal(findPython(run, {}), 'python');
  assert.deepEqual(commands, ['python3', 'python']);
  assert.equal(findPython(run, { HWP_TO_WORD_PYTHON: '/path with spaces/python' }), '/path with spaces/python');
  assert.throws(() => findPython(() => ({ status: 1 }), {}), /Python 3.9/);
});

test('CLI runs through npm-style executable symlinks', { skip: process.platform === 'win32' }, () => {
  const directory = mkdtempSync(join(tmpdir(), 'hwp-cli-link-'));
  try {
    const link = join(directory, 'hwp-to-word');
    symlinkSync(resolve('bin/hwp-to-word.mjs'), link);
    const result = spawnSync(process.execPath, [link, '--version'], { encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(result.stdout.trim(), version);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

for (const platform of ['darwin', 'linux', 'win32']) test(`runtime is cached and cleans up installation locks on ${platform}`, () => {
  const cacheRoot = mkdtempSync(join(tmpdir(), 'hwp-cli-test-'));
  try {
    assert.throws(() => ensureRuntime({ cacheRoot, python: 'python3', platform, run: () => 1 }), /virtual environment/);
    let calls = 0;
    const executable = ensureRuntime({ cacheRoot, python: 'python3', platform, run: (command, args) => {
      calls += 1;
      if (args[1] === 'venv') {
        const binary = platform === 'win32' ? join(args[2], 'Scripts', 'python.exe') : join(args[2], 'bin', 'python');
        mkdirSync(dirname(binary), { recursive: true });
        writeFileSync(binary, '');
      }
      return 0;
    } });
    assert.equal(calls, 2);
    assert.ok(existsSync(executable));
    assert.equal(ensureRuntime({ cacheRoot, python: 'python3', platform, run: () => assert.fail('cached runtime should not install') }), executable);
  } finally {
    rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('process invocation preserves arguments and exit codes without a shell', () => {
  assert.equal(invoke(process.execPath, ['-e', 'process.exit(process.argv[1] === "file with spaces.hwp" ? 7 : 9)', 'file with spaces.hwp'], { stdio: 'ignore' }), 7);
});
