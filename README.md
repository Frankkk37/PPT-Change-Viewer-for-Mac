# PPT Diff Tool v0.9 GUI【Mac版本】

本版本继续优先解决：

1. 人工选择文件 A 和文件 B。
2. 准确比较两个 PPT。
3. 用简单 UI 降低使用成本。
4. 提升文字差异报告的可读性。

## v0.9 关键更新

v0.7 已经把 PowerPoint XML 的碎片 text run 合并成段落/句子块，避免报告过碎。

但段落/句子过长时，使用者仍然需要读完整个旧句和新句，才能找到改动点。

v0.9 改为：

```text
先合并同一段落里的 text run
再按逗号、顿号、分号、冒号、句号等标点切成短句/分句
最后只展示发生变化的短句/分句
```

报告会更接近：

```text
句内新增：
  - 旧：热门靶点内卷倒逼差异化，能开发
  - 新：热门靶点内卷倒逼差异化，能持续开发
  - 变化：新增 `持续`
```

而不是展示整段很长的句子。

## 文件说明

```text
ppt_diff_tool.py          主程序：支持 GUI 和命令行
打开PPTDiff.command       macOS 双击启动 GUI
README.md                使用说明
```

## 推荐使用方式：双击打开

在 macOS 上：

1. 解压文件夹。
2. 双击 `打开PPTDiff.command`。
3. 选择旧版 PPT。
4. 选择新版 PPT。
5. 选择输出文件夹。
6. 点击“比较两个 PPT”。

如果 macOS 提示没有权限，在终端中运行：

```bash
chmod +x 打开PPTDiff.command
```

然后再双击。

## 命令行方式

```bash
python3 ppt_diff_tool.py "旧版.pptx" "新版.pptx" -o diff_output
```

如需检测格式/布局变化：

```bash
python3 ppt_diff_tool.py "旧版.pptx" "新版.pptx" -o diff_output --detect-format
```

## 输出文件

每次比较会生成：

```text
xxx.diff.md
xxx.diff.json
```

`.md` 是人看的报告，`.json` 是机器可读数据。

## 当前检测能力

- 新增页面
- 删除页面
- 移动页面
- 修改页面
- 修改页文字差异
- 短句/分句级 old/new 展示
- 图片变化
- 形状数量变化
- 表格/图表容器数量变化
- 页码/自动编号变化过滤
- 可选：格式/布局变化检测

## 当前限制

短句切分基于标点规则，不是自然语言模型。因此当一行文字没有明显标点时，仍可能显示较长片段。后续可以增加“变化点前后 N 个字”的上下文模式。

v0.9 默认关闭“格式/布局变化检测”。如果你只关心内容变化，保持默认关闭。如果你要检查格式调整版，再打开该选项。


## v0.9 新增

- 新增 `xxx.diff.html`，作为默认主要报告。
- HTML 报告包含摘要卡片、增删移动页清单、可折叠的修改页详情。
- 文字变化在 HTML 中用高亮显示：删除/替换旧内容、插入/新增新内容。
- UI 中新增“打开 HTML 报告”按钮。

输出文件现在包括：

```text
xxx.diff.html
xxx.diff.md
xxx.diff.json
```
Copyright © 2026 Frank Fan. All rights reserved.

本项目仅用于个人作品展示和内部测试。未经许可，不得复制、修改、分发或商业使用。
