# gba-mep-docx-toc

一个 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（`dsh`）bundle 插件，注册一个技能：**为已有的 `.docx` 插入真正的 Word 目录**。

技能直接改写 `word/document.xml`，写入真正的 `TOC` 域与 `_Toc` 书签——按标题级别分级缩进、点前导线、页码右对齐——只用 Python 标准库。不需要 Microsoft Word、不需要 LibreOffice、不需要 COM、不需要任何第三方包。

## 安装

```sh
dsh plugin --profile web add gba-mep-docx-toc
```

装完重启一次 `dsh web`，让 bundle 层重新合成。本仓库也支持直接从源码安装，并收录于 [awesome-dsh-plugin](https://awesome-dsh-plugin.com) 目录，因此 [dsh-market](https://dshmarket.com) 可一键安装。

## 它做什么

给一份标题使用了 `Heading 1`–`Heading 9` 样式（或带 `w:outlineLvl`）的 `.docx`：

1. 按文档顺序找出标题；
2. 为每个标题插入 `_Toc…` 书签，使目录项可跳转；
3. 构造 `TOC \o "1-3" \h \z \u` 域，每个条目 = 超链接锚点 → 标题文字 → 右对齐点前导 tab → `PAGEREF` 域；
4. 二级缩进 `100`、三级缩进 `200` 字符单位；
5. 只重写 `word/document.xml`，包内其余部件逐字节复制；
6. 用 `verify` 自检（条目数 = 书签数 = `PAGEREF` 数）。

## 环境要求

- 提供 `skills` 服务的 DSH 版本——入口不 import 任何宿主包，只依赖该服务契约（已按 DSH `0.2.0-rc.2` 核对）。
- 跑脚本的机器上有 Python 3.9+，无需安装任何包。
- 作为 dsh 插件自然需要 Node.js；入口本身只用 `node:fs`、`node:path`、`node:url`。

## 用法

```sh
python scripts/docx_toc.py build 输入.docx -o 输出.docx --title "目　录" --page-break both
python scripts/docx_toc.py verify 输出.docx
python scripts/docx_toc.py selftest
```

在会话里也可以直接描述任务——模型按技能的 `description` 自行路由。

## 边界声明

- **页码是占位值**。Python 不分页，真实页码要等 Word 更新域（`Ctrl+A` → `F9` → 仅更新页码）后才出现；在此之前所有条目都显示 `1`。
- **不处理分节页码**。文档分节且各节重新编页（前置罗马数字、正文阿拉伯数字）时，`PAGEREF` 结果取决于 Word 的节设置，本脚本不做保证。
- **不修改已有目录**。文档里已有一个 `TOC` 域时会生成第二份，请先删除旧的。
- **不保证与 Word 完全一致的分页**。渲染验证用的是 LibreOffice，其分页规则与 Microsoft Word 存在差异。
- **标题必须使用标题样式**。手动加粗放大的文字，脚本找不到。

## 目录结构

```
lib/index.js            bundle 入口——经 ctx.skills.register() 注册技能
cordis.patch.yml        loader 补丁：一行 insert，按包名引用本包
SKILL.md                技能正文（frontmatter：name / description / whenToUse）
scripts/docx_toc.py     实现（纯标准库）
references/OOXML-TOC.md OOXML 目录域的笔记
tests/                  离线测试：入口行为 + 可安装性合约
screenshots.json        插件市场展示用截图声明
```

## 测试

```sh
node --test tests/
```

`tests/entry.test.mjs` 用最小上下文真执行 `apply()`，断言注册字段、`resourceBase` 与随包资源齐备，并含一条 CRLF 回归；`tests/bundle-contract.test.mjs` 守可安装性合约：声明了 `dsh.bundle`、patch 行按包名引用、`files` 白名单完整、入口只 import `node:` 内建。

## 许可证

MIT——全文见 [LICENSE.md](LICENSE.md)。

---

由 **GBA 工程自动化** 维护 ｜ 交流：david_1999cn@hotmail.com
