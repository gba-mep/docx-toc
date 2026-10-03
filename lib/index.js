// gba-mep-docx-toc — DeepSeek Harness bundle entry point.
//
// 本入口在 bundle 被挂载时，把随包的 Word 目录技能注册进宿主的 `skills` 服务。
//
// 设计约束（两条，都会影响可安装性）：
//   ① **入口不 import 任何宿主包**——只用 `node:` 内建模块。宿主用 pnpm 安装插件时只为
//      已声明的依赖建链接；入口一旦静态 import `@deepseek-ai/*`，在解析不到该包的布局下
//      入口 import 就会抛错，技能连带注册失败（失败形态：装上了、但技能不存在）。
//      因此 package.json 里没有 `peerDependencies`，宿主版本兼容由「只依赖 skills 服务契约」保证。
//   ② **技能目录 = 包根**。SKILL.md / scripts/ / references/ 与 package.json 同级，
//      `resourceBase` 指向包根，技能内 `scripts/docx_toc.py`、`references/OOXML-TOC.md`
//      这类相对引用在任意 cwd 下都能解析。
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'gba-mep-docx-toc'
// `skills` 是硬依赖：合规宿主在 apply 之前就已就绪（服务契约见 DSH 0.2.0 的 `skills` 服务）。
export const inject = ['skills']

/** 包根：本文件位于 lib/，故根目录是它的上一级。 */
const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const skillFile = join(packageRoot, 'SKILL.md')

/**
 * frontmatter 缺失或损坏时的兜底描述。
 * 模型侧靠 description 决定是否加载技能，所以宁可给一句保守的自述，也不要空描述。
 */
const FALLBACK_DESCRIPTION =
  '为已有 .docx 插入真正的 Word 目录域（TOC）：分级缩进、点前导线、页码右对齐，并为标题打 _Toc 书签。纯 Python 标准库实现。'

/**
 * 切出 SKILL.md 的 YAML frontmatter。
 *
 * 为什么要先归一化行尾与 BOM：`readFileSync(..., 'utf8')` 不做行尾归一。SKILL.md 一旦是 CRLF
 * （Windows 检出或编辑器另存为 CRLF），首行就是 `---\r\n`，用 `startsWith('---\n')` 判定会为假，
 * frontmatter 整块**静默**退化成兜底描述——description / whenToUse 正是模型侧路由的唯一依据，
 * 静态检查全绿也照样漏过。这里统一归一，并且要求闭合 `---` 必须成行（缺闭合 = 未解析，不猜）。
 *
 * @param {string} text - SKILL.md 原始内容。
 * @returns {{ fields: Record<string, string | undefined>, body: string, parsed: boolean }}
 */
function splitFrontmatter(text) {
  const norm = String(text).replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n')
  const match = /^---[ \t]*\n([\s\S]*?)\n---[ \t]*(?:\n|$)/.exec(norm)
  if (!match) return { fields: {}, body: norm, parsed: false }
  const meta = match[1]
  const body = norm.slice(match[0].length).replace(/^\n+/, '')
  const read = (key) => {
    const hit = new RegExp(`^${key}:\\s*(.+)$`, 'm').exec(meta)
    return hit?.[1]?.trim().replace(/^["']|["']$/g, '')
  }
  return {
    fields: { name: read('name'), description: read('description'), whenToUse: read('whenToUse') },
    body,
    parsed: true,
  }
}

/**
 * 状态行出口：宿主有 `ctx.logger` 就走日志面，否则退回 stderr（宁可见勿静默）。
 * @param {unknown} ctx - 插件上下文（最小宿主可能没有 logger）。
 * @returns {(level: 'info' | 'warn', message: string) => void}
 */
function makeReporter(ctx) {
  const logger = /** @type {{ logger?: Record<string, (m: string) => void> }} */ (ctx)?.logger
  return (level, message) => {
    if (logger && typeof logger[level] === 'function') logger[level](message)
    else console.error(message)
  }
}

/**
 * 注册随包技能。注册本身是一个 effect：返回的 disposer 在卸载时移除该贡献。
 *
 * @param {unknown} ctx - Cordis 上下文；`skills` 服务来自 `inject` 声明。
 * @returns {void}
 */
export function apply(ctx) {
  const say = makeReporter(ctx)

  let raw
  try {
    raw = readFileSync(skillFile, 'utf8')
  } catch (error) {
    // 入口与 SKILL.md 同包，读不到就是包本身坏了——响亮失败，不假装注册成功。
    say('warn', `· gba-mep-docx-toc 技能未注册：读不到 ${skillFile}（${String(error?.message ?? error)}）`)
    return
  }

  const { fields, body, parsed } = splitFrontmatter(raw)
  const skillName = fields.name ?? name
  if (fields.name && fields.name !== name) {
    say('warn', `· gba-mep-docx-toc 技能名不一致：SKILL.md frontmatter "${fields.name}" ≠ 包名 "${name}"——以 frontmatter 为准注册。`)
  }
  if (!parsed) {
    say(
      'warn',
      `· gba-mep-docx-toc 技能 frontmatter 未解析（首行不是成行 \`---\` 或缺闭合行）：${skillFile}` +
        ' —— 已退回兜底 description，whenToUse 未注册；请检查该文件是否被 CRLF / BOM / 前置空行破坏。',
    )
  }

  const skills = /** @type {{ skills?: { register?: (s: unknown) => () => void } }} */ (ctx)?.skills
  if (typeof skills?.register !== 'function') {
    // 合规宿主在 apply 前就已提供（inject 硬依赖）。这不是降级路径而是宿主违约路径：
    // 点名 + 跳过，绝不以 `Cannot read properties of undefined` 的形式出现。
    say('warn', '· gba-mep-docx-toc 技能未注册：宿主 ctx 缺 `skills` 服务（inject 声明的硬依赖未满足）。')
    return
  }

  ctx.effect(() =>
    skills.register({
      name: skillName,
      source: 'runtime',
      description: fields.description ?? FALLBACK_DESCRIPTION,
      ...(fields.whenToUse ? { whenToUse: fields.whenToUse } : {}),
      content: body,
      // 技能目录 = 包根：SKILL.md 里的 scripts/、references/ 相对路径据此解析。
      resourceBase: { kind: 'directory', path: packageRoot },
    }),
  )

  say('info', `· gba-mep-docx-toc 技能已注册（${skillName}；resourceBase=${packageRoot}）`)
}
