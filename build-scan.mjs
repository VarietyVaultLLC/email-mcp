// build-scan.mjs — builds ONLY dist/cli/inbound-scan.js.
//
// Deliberately separate from build.mjs: the live email-mcp MCP server runs
// from dist/index.js, and any rebuild of that file requires a FULL Claude
// Code restart (see CLAUDE.md). This script never touches dist/index.js,
// so the scanner can be iterated without disturbing the running server.
//
// Run: node build-scan.mjs
import { build } from 'esbuild';

await build({
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
  entryPoints: ['src/cli/inbound-scan.ts'],
});

console.log('inbound-scan build complete (dist/index.js untouched).');
