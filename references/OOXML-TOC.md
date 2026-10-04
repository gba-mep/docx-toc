# OOXML 目录（TOC）原理与坑位

本文件是 `docx_toc.py` 的技术背景。只讲通用 OOXML/Word 知识，与任何具体项目无关。

## 1. Word 目录本质是一个「域」

目录不是普通文字，而是 Word 的 **field**。在 `word/document.xml` 里长这样：

```xml
<w:p>                                  <!-- 第一个目录条目段落 -->
  <w:pPr>…</w:pPr>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:hyperlink w:anchor="_Toc90000001">
    <w:r><w:t>第一章 总则</w:t></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc90000001 \h </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>            <!-- 域结果：页码 -->
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:hyperlink>
</w:p>
…
<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>   <!-- 整个 TOC 域结束 -->
```

- `TOC \o "1-3"` = 收集 1–3 级大纲；`\h` = 条目做成超链接；`\z` = Web 版视图隐藏页码；`\u` = 用大纲级别。
- 域结果（`separate` 与 `end` 之间）是**缓存**。按 `F9` 时 Word 重新扫描标题、重建条目、刷新页码与书签。

### 为什么要自己打书签

`HYPERLINK \l "_Toc…"` 与 `PAGEREF _Toc…` 都指向一个书签。书签不存在时，点击不跳转、页码域报错。Word 首次更新域时会自己补书签；但若希望**打开就有可点目录**，就必须自己把 `bookmarkStart` / `bookmarkEnd` 打进每个标题段落。

书签 id 必须唯一，名字建议 `_Toc` + 8~9 位数字（Word 惯例）。

### 占位页码

不排版就无法知道页码。工程做法是填占位值（如 `1`），由用户在 Word 里更新域。想更省事，也可以在 `settings.xml` 里设置打开时更新域（`<w:updateFields w:val="true"/>`），但 Word 会弹窗询问，体验因版本而异，不如直接提示按 F9。

## 2. XML 子元素顺序是硬约束

OOXML 是 schema 校验的，**子元素顺序错了 Word 会报「内容有问题」**（甚至直接修不好）。两个必须记住的顺序：

**`w:pPr` 常用子元素顺序**（截取常用段）：

```
pStyle → keepNext → keepLines → pageBreakBefore → framePr → widowControl
→ numPr → pBdr → shd → tabs → spacing → ind → jc → outlineLvl → rPr → sectPr
```

注意 **`tabs` 在 `spacing` 之前，`spacing` 在 `ind` 之前，`ind` 在 `jc` 之前**。

**`w:rPr` 常用子元素顺序**：

```
rStyle → rFonts → b → bCs → i → iCs → caps → strike → color → spacing
→ sz → szCs → highlight → u → shd → vertAlign → lang
```

注意 **`rFonts` 最前，`b` 紧跟其后，`sz` / `szCs` 在很后面**。

> 实践中「加了粗体却整段报错」几乎都是 `b` 插到了 `sz` 后面。

## 3. 中文字体要显式声明 eastAsia

```xml
<w:rFonts w:hint="eastAsia" w:ascii="微软雅黑" w:hAnsi="微软雅黑" w:eastAsia="微软雅黑"/>
```

只写 `w:ascii` 对中文无效——中文字符走 `w:eastAsia` 分支。这是中文文档最常见的「字体没生效」原因。

## 4. 缩进用字符单位，不要写死 twips

```xml
<w:ind w:leftChars="100" w:left="240"/>   <!-- 二级：1 字符 -->
```

- `w:leftChars` 单位是 **1/100 字符**，Word 按当前字号换算，换字号不会错位。
- `w:left` 单位是 twips，写死了就与字号绑定（12pt 时 1 字符 ≈ 240 twips，10.5pt 时 ≈ 210）。
- 两个都写：`leftChars` 给 Word 用，`left` 给不认 `leftChars` 的渲染器兜底。

## 5. ⚠️ 已生成的目录不能用 run 级格式工具改

这是最容易造成**不可逆损坏**的一条。

- 目录条目是**域**，其文字被拆成大量 run（超链接、tab、PAGEREF、页码各自成 run）。
- 任何**跨越整个条目的 run 级格式操作**（例如「选中这段改成某字体」）会重写这些 run，把 `fldChar` / `instrText` 一并清掉 → 目录变成一堆死文字，页码消失、点击不跳转。
- **安全**的做法只有两种：
  1. 段落级属性（`w:pPr` 里的缩进、间距、对齐）——不动 run，安全；
  2. 直接改 `word/document.xml` 文本，精确替换目标属性。

改目录字体时，用正则在 XML 里替换 `w:rFonts` 属性串，而不是通过编辑器的「设置字体」接口。

## 6. 为什么 Word 默认目录是「平铺」的

若文档的 `styles.xml` 里没有 `TOC1` / `TOC2` / `TOC3` 样式定义，Word 生成目录时会退回 `Normal` 样式：

- 所有层级**同一缩进、一律不加粗**；
- 字体常常是 `SimSun`（宋体）。

要分级，两条路：

1. **自己补** `w:ind` / `w:b`（本脚本走这条，不依赖样式存在）；
2. 在 `styles.xml` 里定义 TOC1–3 样式（更规范，但已生成的域结果要按 F9 才会用上新样式）。

## 7. 怎么认出哪些段落是「目录项」

| 办法 | 说明 | 可靠度 |
|---|---|---|
| 找 `PAGEREF` | 逐段找含 `PAGEREF` 的段落 | ★★★ |
| 找 `SimSun` | Word 内建目录默认用宋体；先数全文件 `w:eastAsia="SimSun"` 的段落数，**等于标题数**才可安全批量处理 | ★★★（前提是文档里没有别的宋体段落） |
| 解析导出的 PDF 文字 | 按版式猜 | ★（曾实测多认出多余项，不要用） |

## 8. 条目层级从哪来

```
目录项（按文档顺序） ≡ 标题节点（按文档顺序），一一对应
```

因此只要按文档顺序取出所有标题的层级序列，依次套到目录项上即可，不必解析域内部结构。

⚠️ 若用文档结构接口拿标题列表，注意**分页参数**：限制了条数就只拿到前几个，层级序列会与目录项错位，断言直接失败。

## 9. 验证清单

跑完必查（`docx_toc.py verify` 已内置）：

- `TOC \o "1-` 域指令存在
- 条目数 == `PAGEREF` 数 == `bookmarkStart` 数 == 标题数
- 点前导线 `w:leader="dot"` 存在
- 该文档的所有 XML 部件都能被 XML 解析器解析（zip 完整 + well-formed）

## 10. 刷新与撤销

- **刷新**：`Ctrl+A` → `F9` → 选「只更新页码」。（选「更新整个目录」会重建条目，覆盖你手改的缩进——所以本脚本把缩进写进条目属性，并优先建议「只更新页码」。）
- **撤销**：改坏了就回退文件。若走 git，直接 checkout 原档最稳。

## 11. 渲染器差异

- LibreOffice 会**给域加灰底**显示，这是它的域可视化，不是文档里的真实底纹；Microsoft Word 默认只在选中时显示域底纹。
- 两大引擎的**分页算法不同**，同一份文档页数可能差 1 页。目录页码以 Word 更新域后的结果为准。
