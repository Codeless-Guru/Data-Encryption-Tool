# v0.1.0 发布准备清单

> 生成日期：2026-11-11 · 基于提交 `c7005c2`（main）的实测结果

## 0. 发布判定

**结论：可以发 v0.1.0，但发布说明里不得宣传 AES-256-CBC 与 AGE 公钥模式。**

依据（本机实测，非文档转述）：

| 检查项 | 命令 | 结果 |
|---|---|---|
| 自检套件 | `python tests/test_skill.py` | 19 项通过 / 1 项跳过（SM4 需 gmssl）/ 退出码 0 |
| Python 版本 | `python --version` | 3.13.16（要求 >= 3.10，满足） |
| 依赖 | `pip install "cryptography>=42" openpyxl` | cryptography 50.0.2 / openpyxl 3.1.5 |
| 算法可用性 | 逐个调用 `encrypt(..., alg)` | `aesgcm` ✅ `chacha20` ✅ `fernet` ✅ `sm4gcm` ⚠需 gmssl `aesgcm-cbc` ❌ `age` ❌ |
| 许可 | `LICENSE` | MIT |
| 密钥泄露扫描 | token/私钥/AKIA 模式 | 零命中 |

## 1. Tag

```
v0.1.0
```

建议用 annotated tag（带说明与作者，`git describe` 能取到）：

```bash
git tag -a v0.1.0 -m "v0.1.0: cell-level reversible encryption for Excel/CSV; self-checks 19/19 (SM4 skipped)"
git push origin v0.1.0
```

## 2. Release 标题

```
v0.1.0 — 单元格级可逆加密，未选中的格子一个字节都不改
```

## 3. 发布说明（可直接粘贴到 GitHub Release 正文）

```markdown
首个可发布版本。按「列 / 行」加密 Excel(.xlsx/.xlsm)、CSV/TSV 与粘贴数据，未选中的单元格逐格零改动，格式原样保留。

## 亮点

- **只碰选中的格子**：逐格改写 `cell.value` 后保存，从不重建工作簿，字体/填充/边框/对齐/数字格式/列宽行高/合并单元格/冻结窗格/隐藏行列/数据验证/条件格式/批注/日期类型/公式字符串全部保留。
- **绝不覆盖源文件**：一律输出 `原名.encrypted.<ext>`，同名自动加 `_2`。
- **口令不落盘**：只从当次环境变量或交互式输入取，不写入代码、YAML、日志与输出文件；密文不含任何可推导口令的信息。
- **三道中止护栏**：列名在全表都找不到 / 选列写法非法 / 命中 0 格 → 直接报错，不产出"看着成功其实没加密"的文件。
- **公式护栏**：以 `=` 开头的格子自动跳过并计入 `formulas_skipped`，避免把整表算废。
- **CSV 保真**：BOM 只写一份、换行符沿用原文、引号与内嵌换行交给 csv 模块，`"张,三"` 不会错行。
- **密文自描述**：`enc:v1:<alg>:<b64 salt>:<b64 nonce>:<b64 密文+tag>`，每格独立 salt + nonce，同一明文两次加密结果不同，杜绝等值比对；KDF 为 Scrypt(n=2^14,r=8,p=1)。

## 当前功能

**输入**：`.xlsx` / `.xlsm`（`keep_vba` 保宏）/ `.csv` / `.tsv` / 粘贴文本（TSV、`｜` 分隔、Markdown 表格）/ 一行一值纯文本。

**选范围**：
- 列：列名（表头精确匹配）、列字母 `C`、区间 `D:F`、列号 `3`
- 行：`--rows 2-100`（可多段叠加）
- 表头：默认保留不加密，`--no-header` 可连表头一起加密
- 多工作表语义：按列名指定 → 无此表头的 sheet 整页跳过并记入 `skipped_sheets`；按列字母/列号 → 视为显式跨表指令

**模式**：`aesgcm`（默认）/ `chacha20` / `fernet` / `sm4gcm`（需 gmssl）/ `hmac`（单向，可 JOIN）/ `mask`（单向，仅展示）

**工作流**：认输入形态 → 定范围 → 定算法与口令 → dry-run 报数确认 → 执行并跑自检 → 交付时报告算法、加密格数、跳过格数、`skipped_sheets`、`formulas_skipped`。

**负向拦截**：错口令、密文篡改 1 bit、把明文当密文、垃圾串、未知算法、AAD 域不匹配 —— 全部抛异常，不返回半截数据。

## 已知限制（v0.1.0）

- **AES-256-CBC 未实现**：`MODES` 中保留了名字，调用会抛 `ValueError: 未知算法: aesgcm-cbc`。README/SKILL 中的 CBC 表述为规划项，请以本行说明为准。
- **AGE 公钥模式未实现**：需 `age`/`rage` 二进制，当前无对应代码分支，调用抛 `未知算法: age`。
- **SM4-GCM 依赖 gmssl**：未安装时明确报错，不会静默降级为 AES。
- **样式保真为人工核对**：自检套件覆盖 sheet 顺序、合并区域、未选中格的值与回环；字体/颜色/列宽/数字格式的自动断言尚未落地（计划 v0.1.1 补齐）。
- **不支持的输入**：`.xls`、`.ods`、被密码保护的 workbook（请先解密另存）。

## 安装与验证

```bash
pip install "cryptography>=42" openpyxl    # Python >= 3.10
python tests/test_skill.py                 # 期望：通过 19 / 19，退出码 0
```

作为 Agent Skill 使用时，把整个目录放进 Skills 目录即可，无需构建或注册。

## 已知会丢的格式（交付时须告知）

图表 / 透视表 / 切片器 / 迷你图、VBA 窗体与 ActiveX 控件不被 openpyxl 保留；被加密列若被公式引用，该公式会变成 `#VALUE!`（dry-run 阶段应点名）。

## 许可

MIT
```

## 4. 当前功能（速查表）

| 维度 | 内容 |
|---|---|
| 输入 | `.xlsx` `.xlsm` `.csv` `.tsv` 粘贴文本（TSV / `｜` / Markdown 表格）、一行一值 |
| 选列 | 列名 / 列字母 / 区间 / 列号 |
| 选行 | `1`、`2-100`、多段叠加 |
| 表头 | 默认保留，可关闭 |
| 可逆加密 | AES-256-GCM（默认）、ChaCha20-Poly1305、Fernet、SM4-GCM（需 gmssl） |
| 不可逆脱敏 | HMAC-SHA256（同口令同摘要，可 JOIN）、掩码（手机/身份证/邮箱/文本） |
| 格式保全 | 字体、填充、边框、对齐、数字格式、列宽行高、合并区、冻结窗格、隐藏行列、数据验证、条件格式、批注、日期类型、公式、VBA 宏、sheet 顺序 |
| 安全机制 | 每格独立 salt+nonce、AEAD（AAD=`DETv1`）、幂等跳过、Scrypt KDF、源文件不覆盖、口令不落盘 |
| 护栏 | 列名拼错中止、选列语法非法中止、零命中空转中止、公式跳过 |
| 自检 | 19 项通过 / 1 项跳过（SM4），退出码 0 |
| 许可 | MIT |

## 5. 发布前建议修正（按优先级）

1. **收回未实现算法的宣传**（阻断级）：README 第 4 行的「AES-256-CBC / AGE 公钥」，SKILL.md 的 `MODES["aesgcm-cbc"]`、算法表里的 `age` 行、E3 示例的 `mode=age`，均无对应实现。两种处理：标注「规划中（v0.2）」，或在 `encrypt()` 里补 CBC 分支（AGE 需外部二进制，建议留到 v0.2）。
2. **修正断言口径**（阻断级）：README 写「22 条断言 / 7 项样式保真 / 逐格断言含字体、颜色、数字格式、列宽」，但 `tests/test_skill.py` 的编号从 3 直接跳到 8，字体/颜色/列宽/数字格式**没有任何自动断言**。要么补 4–7 号用例，要么把 README 口径改成「19 项自动 + 样式人工核对」。
3. **清理 SKILL.md 中的恒真断言**（阻断级）：第 4 节示例 `assert ... and sa.value is not vb if False else True` 恒为真（等于没断言），`k = _u(...) if False else base64...` 为死代码。照抄这套代码会得到假的「全绿」。
4. **加版本元数据**（建议）：`SKILL.md` frontmatter 增加 `version: 0.1.0`；新增 `CHANGELOG.md` 记录本次发布。
5. **加 `.gitignore`**（建议）：忽略 `*.encrypted.xlsx`、`*.encrypted.csv`、`*.encrypted.tsv`、`*.DECRYPTED.*`、`__pycache__/`，避免样例密文与测试产物误提交。
6. **加 CI**（建议）：`.github/workflows/test.yml` 跑 `python tests/test_skill.py`，把「19/19 全绿」变成仓库自动保证；注意公开仓库首次推送 workflow 文件需要令牌带 `workflow` 权限（当前令牌缺该 scope）。

## 6. 发布前修正：已落实项（2026-11-11）

| # | 修正内容 | 文件 | 证据 |
|---|---|---|---|
| 1 | 首屏算法宣传改为「已实现 4 种可逆 + 2 种单向」，CBC/AGE 明确标注规划项 | `README.md` | 首屏第 5 行、§2.3 表新增 `⏳` 行 |
| 2 | 断言口径由「22 条 / 7 项样式」改为「19 项自动 + 样式人工核对」，并注明 4–7 号为空号 | `README.md` §1、§2.5、§4.2；`SKILL.md` §4 | 复跑 `tests/test_skill.py` → 19/19、退出码 0 |
| 3 | 清理两处恒真断言与死代码：`if False else True` 改为真实断言（`number_format` 相等 + `data_type` 相等）；`k = _u(...) if False else ...` 改为直接 urlsafe_b64 编码 | `SKILL.md` §2 代码块、§4 `assert_roundtrip` | 代码块语法检查通过 |
| 4 | `MODES` 移除 `aesgcm-cbc` 与 `age`，改为注释形式的「规划中」说明 | `SKILL.md` §2 | 同上 |
| 5 | 算法表、E3 示例、`references/algorithm-params.md` 的 age/CBC 行标注未实现 | `SKILL.md`、`references/` | — |
| 6 | frontmatter 增加 `version: 0.1.0` | `SKILL.md` | `yaml` 解析通过 |
| 7 | `agents/openai.yaml`：`optional_bins` 清空（age 属 v0.2）、护栏改指向真实自检套件 | `agents/openai.yaml` | PyYAML 解析通过，`optional_bins= []` |
| 8 | 新增 `.gitignore`（忽略 `*.encrypted.*`、`*.DECRYPTED.*`、`__pycache__/` 等） | `.gitignore` | 新增文件 |
| 9 | 新增 `CHANGELOG.md`（v0.1.0 条目） | `CHANGELOG.md` | 新增文件 |

未落实项：CI 工作流（需要令牌 `workflow` scope，当前链路无法完成 scope 刷新）。自动样式断言（v0.1.1 计划）。

## 7. 打标签与建 Release 的命令
```bash
cd /d E:\Demo\Data-Encryption-Tool

# 1) 提交本次发布准备（如决定把本文件入库）
git add RELEASE-NOTES-v0.1.0.md
git commit -m "docs: prepare v0.1.0 release notes"

# 2) 打 annotated tag 并推送
git tag -a v0.1.0 -m "v0.1.0: cell-level reversible encryption for Excel/CSV; self-checks 19/19 (SM4 skipped)"
git push origin main
git push origin v0.1.0

# 3) 创建 GitHub Release（正文用第 3 节，标题用第 2 节）
gh release create v0.1.0 --title "v0.1.0 — 单元格级可逆加密，未选中的格子一个字节都不改" --notes-file RELEASE-NOTES-v0.1.0.md
```

若只想先发草稿、不对外可见，把第 3 步换成：

```bash
gh release create v0.1.0 --draft --title "..." --notes-file <(sed -n '/^## 3\./,/^## 4\./p' RELEASE-NOTES-v0.1.0.md)
```
