# New-Exam-Tool · 试题转换台

把考试题（上传 `.docx` 或粘贴文字）一键转成**考试系统导入模板**格式的 Excel。
单个 HTML 文件，**完全离线**运行：没有「AI 出题」页，没有任何联网代码，题目内容不会离开你的设备。

当前版本：见 `src/version.json`（页面右下角也会显示版本号和更新日期，新版带 **NEW** 标签）。

## 功能

- 上传 `.docx` 或粘贴文字，自动识别单选、多选、判断题和答案，导出 7 个工作表的导入模板（填写前 4 张：试卷信息、单选题、多选题、是非题）。
- 支持三种答案写法：文末「参考答案」区块、题干括号里直接写答案（`……（B）`）、独立一行的 `【答案】B` / `答案：ABD`。
- 页面上直接修改题干、选项、答案，可删除、可新增题目。
- 导出前自动检查：选项不足、单选多答案、多选少于 2 个答案、题数不对（既对照标准 10/5/5，也对照试卷自己写的题数）、总分不是 100。有问题时第一次点击只提示，**再点一次才导出**。
- 默认每题 5 分；10 单选 + 5 多选 + 5 判断 = 20 题 = 100 分。
- 选到 `.doc` / `.wps` / `.pdf` 等不支持的文件时，明确提示「请另存为 .docx」。

使用方法见 [使用说明.md](使用说明.md)。

## 用 GitHub Pages 在线打开（手机、平板也能用）

一次性设置：仓库 → **Settings → Pages → Source 选 `Deploy from a branch` → Branch 选 `main`、目录 `/ (root)` → Save**。
1～2 分钟后访问：

    https://xiaokaishuibuxing.github.io/New-Exam-Tool/

之后每次把新的 `index.html` 推送到 `main`，页面会自动更新。
如果打开后右下角版本号不是最新，是浏览器缓存：强制刷新（电脑 Ctrl+F5 / Cmd+Shift+R；手机可用浏览器的「无痕模式」打开，或清除该网站数据）。

## 本地构建与测试

源码在 `src/`，`index.html` 是构建出来的成品：

    python3 build.py                       # 生成离线版 index.html
    python3 build.py --edition online      # 生成带「AI 出题」的在线版 dist/online.html（需联网，非必需）
    python3 tests/run_tests.py             # 构建并跑全部回归测试（上传前必跑）

回归测试依赖：`pip install playwright python-docx openpyxl`，并安装 Chromium（`playwright install chromium`）。
测试内容：误吸、胸腔闭式引流、胰岛素三套固定试题（粘贴和 `.docx` 两种输入）、题干内嵌答案、答案缺失时的二次确认、页面改题、题数检查、不支持文件的提示、版本号、离线保证（整个流程无网络请求）、`index.html` 与源码一致。
全部通过退出码为 0。**每次改动后、上传前先跑一遍。**

## 目录结构

| 路径 | 作用 |
|---|---|
| `index.html` | 成品（离线版），GitHub Pages 部署的就是它 |
| `src/template.html` | 页面结构和样式 |
| `src/app.js` | 全部业务逻辑（识别、编辑、检查、导出） |
| `src/version.json` | 版本号、更新日期、更新内容 |
| `src/jszip.min.js` | 第三方压缩库（JSZip，MIT） |
| `src/template_data.js` | 内嵌的考试系统导入模板（base64） |
| `build.py` | 把 `src/` 拼成 `index.html` |
| `tests/run_tests.py` | 回归测试 |
| `tests/samples/`、`tests/expected.json` | 固定测试题与期望答案 |
| `proxy/worker.js` | 仅在线版（AI 出题）可选的中转，离线版用不到 |

## 发新版的步骤

1. 改 `src/` 里的源码；
2. 改 `src/version.json` 的 `version`、`date` 和更新内容；
3. `python3 tests/run_tests.py`，确认全部通过；
4. 提交并推送 `index.html` 与 `src/`。
