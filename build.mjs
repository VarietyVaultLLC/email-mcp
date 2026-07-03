import { build, context } from 'esbuild';
import { chmod } from 'node:fs/promises';

const shared = {
  bundle: true,
  platform: 'node',
  target: 'node20',
  format: 'esm',
  outdir: 'dist',
  outbase: 'src',
  sourcemap: true,
  external: [
    'inquirer',
    'google-auth-library',
    'googleapis',
    '@azure/msal-node',
    '@microsoft/microsoft-graph-client',
    'imapflow',
    'mailparser',
    'nodemailer',
  ],
};

if (process.argv.includes('--watch')) {
  const ctx = await context({
    ...shared,
    entryPoints: ['src/index.ts', 'src/setup/wizard.ts'],
  });
  await ctx.watch();
  console.log('Watching for changes...');
} else {
  // Build wizard (no shebang)
  await build({
    ...shared,
    entryPoints: ['src/setup/wizard.ts'],
  });

  // Build headless watcher-fetch CLI (BCC inbox-watcher Phase B consumer).
  // esbuild bundles ONLY the entry points listed here — a new src/cli file
  // that isn't added compiles to nothing, silently.
  await build({
    ...shared,
    entryPoints: ['src/cli/watcher-fetch.ts'],
  });

  // Build CLI entry with shebang
  await build({
    ...shared,
    entryPoints: ['src/index.ts'],
    banner: { js: '#!/usr/bin/env node\n' },
  });

  // Ensure CLI entry is executable
  await chmod('dist/index.js', 0o755);

  console.log('Build complete.');
}
