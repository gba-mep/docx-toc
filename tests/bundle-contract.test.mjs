// bundle 合约测试：守住「这个仓库能被 dsh plugin add 安装」的机械前提。
//
// 这些断言对应官方收录要求（awesome-dsh-plugin contributing.md）：仓库的 package.json
// 必须声明 `dsh.bundle` manifest，并在仓库根放一个 cordis.patch.yml，其 insert 行的
// `name` 必须等于包名——只声明 `dsh.client` 是**不可安装**的，也是最常见的被拒原因。
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const pkg = JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8'))

test('package.json 声明了 dsh.bundle manifest', () => {
  assert.ok(pkg.dsh, 'package.json 需要 dsh 字段')
  assert.ok(pkg.dsh.bundle, '只声明 dsh.client 不可安装；必须有 dsh.bundle')
  assert.equal(pkg.dsh.bundle.patch, './cordis.patch.yml')
  assert.ok(existsSync(join(packageRoot, pkg.dsh.bundle.patch)), 'patch 文件必须真实存在')
})

test('cordis.patch.yml 以 insert 行引用本包包名', () => {
  const patch = readFileSync(join(packageRoot, 'cordis.patch.yml'), 'utf8')
  assert.match(patch, /^-\s*insert:\s*$/m, '新增行必须用 - insert: 形式')
  assert.ok(patch.includes(`name: ${pkg.name}`), `insert 行的 name 必须等于包名 ${pkg.name}`)
  assert.ok(patch.includes('id: '), 'insert 行应有 id，便于其他层按 id 覆盖配置')
})

test('包名合法，入口文件存在', () => {
  assert.match(pkg.name, /^[a-z0-9]+(?:-[a-z0-9]+)*$/, 'npm / DSH 包名应为 kebab-case')
  assert.ok(existsSync(join(packageRoot, pkg.main)), 'main 指向的入口必须存在')
})

test('入口不依赖任何宿主包（只允许 node: 内建）', () => {
  assert.equal(pkg.dependencies, undefined, '入口设计为零依赖：不得有 dependencies')
  assert.equal(pkg.peerDependencies, undefined, '宿主包若用 peerDependencies 声明，需带显式预发布分支；本包选择完全不依赖')
  const source = readFileSync(join(packageRoot, pkg.main), 'utf8')
  const imports = [...source.matchAll(/^import\s+(?:[^'"]*from\s+)?['"]([^'"]+)['"]/gm)].map((m) => m[1])
  assert.ok(imports.length > 0, '入口应有 import')
  for (const specifier of imports) {
    assert.ok(specifier.startsWith('node:'), `入口只能 import node: 内建模块，发现 ${specifier}`)
  }
})

test('files 白名单内的每一条都真实存在（npm 不会静默少发文件）', () => {
  assert.ok(Array.isArray(pkg.files) && pkg.files.length > 0)
  for (const entry of pkg.files) {
    assert.ok(!entry.startsWith('!'), '白名单里不应出现取反项')
    assert.ok(existsSync(join(packageRoot, entry)), `files 里的 ${entry} 不存在`)
  }
  for (const required of ['lib', 'SKILL.md', 'scripts', 'references', 'cordis.patch.yml']) {
    assert.ok(pkg.files.includes(required), `files 必须包含 ${required}`)
  }
})

test('npm 的 repository 指回被收录的 GitHub 仓（官方关联规则）', () => {
  // awesome-dsh-plugin 的规则：已发布 npm 包的 repository 必须指回被收录的那个仓，
  // 否则两者不会关联（这是刻意设计，防止包挂到并未认领它的仓上）。
  assert.ok(pkg.repository, 'package.json 需要 repository 栏位')
  assert.equal(pkg.repository.type, 'git')
  assert.match(pkg.repository.url, /^git\+https:\/\/github\.com\/[^/]+\/[^/]+\.git$/, 'repository.url 应为 GitHub https 形式')
  const slug = pkg.repository.url.replace(/^git\+https:\/\/github\.com\//, '').replace(/\.git$/, '')
  assert.equal(slug, 'gba-mep/docx-toc', 'repository 必须指向清单要收录的那个仓')
  assert.equal(pkg.homepage, `https://github.com/${slug}#readme`)
  assert.equal(pkg.bugs?.url, `https://github.com/${slug}/issues`)
})

test('截图声明指向仓库内的 GitHub 托管路径', () => {
  const file = join(packageRoot, 'screenshots.json')
  if (!existsSync(file)) return // 截图是选用声明
  const list = JSON.parse(readFileSync(file, 'utf8'))
  const shots = Array.isArray(list) ? list : list.screenshots
  assert.ok(Array.isArray(shots) && shots.length >= 1 && shots.length <= 8, '截图 1-8 张')
  for (const shot of shots) {
    assert.ok(!shot.startsWith('/') && !shot.includes('..'), `截图路径不得跳出插件目录：${shot}`)
    assert.ok(existsSync(join(packageRoot, shot)), `截图 ${shot} 必须随仓库存在`)
  }
})
