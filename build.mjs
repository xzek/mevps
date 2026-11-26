/**
 * mevps 前端构建脚本（node build.mjs）
 *
 * 做三件事：
 *   1. 用 esbuild 压缩 src/admin.css 与 src/admin.js
 *   2. 按内容哈希重命名（admin.<hash>.css / admin.<hash>.js），供 immutable 长缓存
 *   3. 把带哈希的引用注入 src/index.html，连同 _headers 一起输出到 public/
 *
 * 产物 public/ 确定性生成：相同源码 → 相同文件名，可放心提交进仓库。
 */
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, mkdirSync, rmSync } from 'node:fs';
import esbuild from 'esbuild';

const cssRaw = readFileSync('src/admin.css', 'utf8');
const jsRaw = readFileSync('src/admin.js', 'utf8');
const tpl = readFileSync('src/index.html', 'utf8');
const headers = readFileSync('src/_headers', 'utf8');

const css = (await esbuild.transform(cssRaw, { loader: 'css', minify: true })).code.trim();
const js = (
  await esbuild.transform(jsRaw, {
    loader: 'js',
    minify: true,
    target: 'es2019',
    charset: 'utf8', // 保留中文字面量，避免 \uXXXX 转义膨胀
  })
).code.trim();

const hash8 = (s) => createHash('sha256').update(s).digest('hex').slice(0, 8);
const cssName = `admin.${hash8(css)}.css`;
const jsName = `admin.${hash8(js)}.js`;

rmSync('public', { recursive: true, force: true });
mkdirSync('public/assets', { recursive: true });
writeFileSync(`public/assets/${cssName}`, css);
writeFileSync(`public/assets/${jsName}`, js);
writeFileSync(
  'public/index.html',
  tpl.replace('{{CSS}}', `/assets/${cssName}`).replace('{{JS}}', `/assets/${jsName}`)
);
writeFileSync('public/_headers', headers);

console.log(`build ok → public/assets/${cssName} (${css.length}B), public/assets/${jsName} (${js.length}B)`);
