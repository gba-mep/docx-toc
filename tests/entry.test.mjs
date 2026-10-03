// 入口行为测试：用最小 ctx 真执行 apply()，断言注册字段与技能资源齐备。
//
// 为什么不用真实宿主跑：本测试要能在**任何**机器上离线运行，而 DSH 宿主需要已安装的
// profile 与运行中的 Electron 运行时。契约面（`skills.register` 的入参形状）已按
// DSH 0.2.0-rc.2 的 `skills` 服务契约核对；这里守的是「入口确实把该注册的东西注册了」。
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const entryUrl = pathToFileURL(join(packageRoot, 'lib', 'index.js')).href
const entry = await import(entryUrl)

/** 造一个只带本测试所需能力的最小 ctx。 */
function makeCtx({ withSkills = true } = {}) {
  const registered = []
  const logs = []
  const ctx = {
    logger: {
      info: (m) => logs.push(['info', m]),
      warn: (m) => logs.push(['warn', m]),
    },
    effect: (fn) => {
      const disposer = fn()
      return () => {
        if (typeof disposer === 'function') disposer()
      }
    },
  }
  if (withSkills) {
    ctx.skills = {
      register: (skill) => {
        registered.push(skill)
        return () => {}
      },
    }
  }
  return { ctx, registered, logs }
}

test('导出名与硬依赖声明符合包契约', () => {
  assert.equal(entry.name, 'gba-mep-docx-toc')
  assert.deepEqual(entry.inject, ['skills'])
  assert.equal(typeof entry.apply, 'function')
})

test('apply() 注册恰好一个技能，字段来自 SKILL.md frontmatter', () => {
  const { ctx, registered, logs } = makeCtx()
  entry.apply(ctx)

  assert.equal(registered.length, 1)
  const skill = registered[0]
  assert.equal(skill.name, 'gba-mep-docx-toc')
  assert.equal(skill.source, 'runtime')
  assert.ok(skill.description.length > 20, 'description 是模型侧路由依据，不应过短')
  assert.ok(skill.description.includes('Word 目录'), 'description 应说明技能做什么')
  assert.ok(skill.content.length > 500, 'body 应是完整技能正文而非摘要')
  assert.ok(skill.content.includes('## 边界声明'), 'body 必须包含边界声明章节')
  assert.ok(!skill.content.includes('slug:'), 'frontmatter 不应混进 body')
  assert.equal(skill.resourceBase?.kind, 'directory')
  assert.equal(skill.resourceBase?.path, packageRoot)
  assert.equal(logs.filter(([level]) => level === 'warn').length, 0, '正常路径不应有告警')
})

test('resourceBase 下的相对引用文件确实存在', () => {
  const { ctx, registered } = makeCtx()
  entry.apply(ctx)
  const base = registered[0].resourceBase.path
  for (const rel of ['SKILL.md', 'scripts/docx_toc.py', 'references/OOXML-TOC.md']) {
    assert.ok(existsSync(join(base, rel)), `技能引用 ${rel} 必须随包存在`)
  }
})

test('宿主缺 skills 服务时跳过注册并响亮告警，不抛错', () => {
  const { ctx, registered, logs } = makeCtx({ withSkills: false })
  assert.doesNotThrow(() => entry.apply(ctx))
  assert.equal(registered.length, 0)
  assert.ok(
    logs.some(([level, message]) => level === 'warn' && message.includes('skills')),
    '缺硬依赖必须有一条点名 skills 的告警',
  )
})

test('CRLF 行尾的 SKILL.md 仍能解析 frontmatter（回归：曾静默退化为兜底描述）', async () => {
  // 把入口与 SKILL.md 复制到临时目录，把 SKILL.md 转成 CRLF 后再加载。
  const dir = mkdtempSync(join(tmpdir(), 'gba-mep-entry-'))
  mkdirSync(join(dir, 'lib'))
  copyFileSync(join(packageRoot, 'lib', 'index.js'), join(dir, 'lib', 'index.js'))
  const crlf = readFileSync(join(packageRoot, 'SKILL.md'), 'utf8').replace(/\r?\n/g, '\r\n')
  writeFileSync(join(dir, 'SKILL.md'), crlf, 'utf8')

  const { ctx, registered, logs } = makeCtx()
  const tempEntry = await import(pathToFileURL(join(dir, 'lib', 'index.js')).href)
  tempEntry.apply(ctx)

  assert.equal(registered.length, 1)
  assert.ok(registered[0].description.includes('Word 目录'), 'CRLF 不应破坏 frontmatter 解析')
  assert.ok(registered[0].content.includes('## 边界声明'))
  assert.equal(logs.filter(([, m]) => m.includes('frontmatter 未解析')).length, 0)
})
