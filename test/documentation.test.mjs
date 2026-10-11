import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';

const root = new URL('../', import.meta.url);
const read = file => fs.readFileSync(new URL(file, root), 'utf8');

test('Claude Code onboarding uses a valid HTTP MCP configuration', () => {
  const readme = read('README.md');
  const snippets = [...readme.matchAll(/```json\s*\n([\s\S]*?)\n```/g)]
    .map(match => JSON.parse(match[1]));
  const config = snippets.find(snippet => snippet.mcpServers?.['ms365-21v']);
  assert.ok(config, 'README must include the Claude Code MCP configuration');
  assert.deepEqual(config.mcpServers['ms365-21v'], {
    type: 'http',
    url: 'https://your-mcp-host.example.cn/mcp'
  });
});

test('Docker defaults and onboarding documents agree on the release version', () => {
  const { version } = JSON.parse(read('package.json'));
  const tag = `v${version}`;
  const image = 'ghcr.io/mdwsk88/ms-365-21v-mcp-server';
  const imageLine = read('docker-compose.release.yml')
    .split('\n')
    .map(line => line.trim())
    .find(line => line.startsWith('image: '));
  assert.equal(imageLine, `image: ${image}:\${MCP_IMAGE_TAG:-${tag}}`);

  for (const file of ['README.md', 'README.en.md', 'docs/CONTAINER.md']) {
    const contents = read(file);
    assert.ok(contents.includes(`${image}:${tag}`), `${file} must show the current image`);
    assert.ok(contents.includes(`MCP_IMAGE_TAG=${tag}`), `${file} must explain the version override`);
  }
  assert.ok(read('docs/CONTAINER.md').includes(`git clone --branch ${tag} `));
});
