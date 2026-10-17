import fs from 'fs';
import { spawn } from 'child_process';

const log = fs.createWriteStream('C:\\Users\\<user>\\Documents\\email-mcp\\mcp-debug.log', { flags: 'a' });
const ts = () => new Date().toISOString();

log.write(`\n\n=== ${ts()} SERVER START ===\n`);
log.write(`argv: ${JSON.stringify(process.argv)}\n`);
log.write(`cwd: ${process.cwd()}\n`);
log.write(`stdin isTTY: ${process.stdin.isTTY}\n`);
log.write(`NODE_PATH: ${process.env.NODE_PATH}\n`);
log.write(`PATH: ${process.env.PATH?.substring(0, 200)}\n`);

const child = spawn(process.execPath, ['C:\\Users\\<user>\\Documents\\email-mcp\\dist\\index.js'], {
  stdio: ['pipe', 'pipe', 'pipe']
});

child.stderr.on('data', (data) => {
  log.write(`${ts()} CHILD STDERR: ${data.toString()}\n`);
});

child.stdout.on('data', (data) => {
  log.write(`${ts()} CHILD STDOUT: ${data.toString().substring(0, 200)}\n`);
  process.stdout.write(data);
});

child.on('error', (err) => {
  log.write(`${ts()} CHILD ERROR: ${err.message}\n`);
});

child.on('close', (code) => {
  log.write(`${ts()} CHILD EXIT: code=${code}\n`);
  process.exit(code || 0);
});

process.stdin.on('data', (data) => {
  log.write(`${ts()} STDIN: ${data.toString().substring(0, 200)}\n`);
  child.stdin.write(data);
});

process.stdin.on('end', () => {
  log.write(`${ts()} STDIN EOF\n`);
  child.stdin.end();
});

process.on('SIGTERM', () => {
  log.write(`${ts()} SIGTERM received\n`);
  child.kill();
});
