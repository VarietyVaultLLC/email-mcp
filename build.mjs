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
    'google-auth-library',
    'googleapis',
    '@azure/msal-node',
    '@microsoft/microsoft-graph-client',
    'imapflow',
    'mailparser',
    'nodemailer',
    'html-to-text',
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
  // Build wizard with shebang (registered as bin: email-mcp-setup)
  await build({
    ...shared,
    entryPoints: ['src/setup/wizard.ts'],
    banner: { js: '#!/usr/bin/env node\n' },
  });

  // Headless CLIs for scripts and schedulers (no MCP server). esbuild builds
  // only the entry points listed here: a new src/cli file that is not added
  // compiles to nothing, silently.
  await build({
    ...shared,
    entryPoints: [
      'src/cli/watcher-fetch.ts',
      'src/cli/inbound-scan.ts',
      'src/cli/jev-fetch.ts',
      'src/cli/jev-send.ts',
    ],
  });

  // Build CLI entry with shebang
  await build({
    ...shared,
    entryPoints: ['src/index.ts'],
    banner: { js: '#!/usr/bin/env node\n' },
  });

  // Ensure both CLI entries are executable
  await chmod('dist/index.js', 0o755);
  await chmod('dist/setup/wizard.js', 0o755);

  console.log('Build complete.');
}
