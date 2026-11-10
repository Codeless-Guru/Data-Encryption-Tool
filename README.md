# Data-Encryption-Tool

> 一个可开源的 Agent Skill：**按「列 / 行」加密 Excel、CSV 和粘贴数据，没选中的地方一个字节都不改。**
> 支持 AES-256-GCM / AES-256-CBC / ChaCha20-Poly1305 / Fernet / SM4-GCM / AGE 公钥，另附 HMAC、掩码两种不可逆脱敏模式。

---

## 1. 项目解决什么问题

日常交付表格时，"把手机号/身份证/薪资这几列打码再发出去"是个高频动作，但现有做法各有硬伤：

| 常见做法 | 问题 |
|---|---|
| 整文件设打开密码 | 对方拿到密码就看到**全部**列，无法只放开几列；且 Excel 密码保护强度很弱 |
| 手敲 `138****5678` 掩码 | **不可逆**，回头自己对账也解不回来；而且常常手滑改到旁边的格子 |
| 另存一份再删列 | 对方拿不到完整结构，回填时列对不上；两份文件迟早版本分叉 |
| 写一次性 Python 脚本 | 每次重写；且 90% 的脚本会用 `pandas` 读写，**公式、字体、合并单元格、列宽、日期类型全丢** |
| 7-Zip / VeraCrypt | 这是整文件/整盘加密，解决不了"同一张表里几列可见几列不可见" |

Data-Encryption-Tool 的定位是**单元格级（cell-level）可逆加密**，并把三条最容易出事的底线做成机制而非提醒：

1. **只碰选中的格子** —— 用 `openpyxl` 逐个改 `cell.value`，从不重建工作簿，样式对象原样带过；未选中单元格在测试里被逐格断言相等（含字体、颜色、数字格式、列宽、合并区、冻结窗格）。
2. **绝不覆盖源文件** —— 一律输出 `原名.encrypted.xlsx`，已存在则加 `_2`；测试第 22 条会校验源文件哈希与明文都在。
3. **口令不落盘** —— 口令只从当次环境变量 / 交互式输入取，不写进 `SKILL.md`、`openai.yaml`、日志或任何输出文件；密文里也不含任何可推导口令的信息。

顺带解决两个隐蔽坑：**公式不是数据**（以 `=` 开头的格子自动跳过并计入 `formulas_skipped`，加密它会把整表算废）；**CSV 的 BOM 与 `\r\n`**（读时剥、写时按原文补回且只补一份，否则列名匹配全失败、Windows 记事本挤成一行）。

---

## 2. 主要功能

### 2.1 输入形态：三种都吃
| 形态 | 支持格式 | 说明 |
|---|---|---|
| 本地表格 | `.xlsx` / `.xlsm` / `.csv` / `.tsv` | `.xlsm` 用 `keep_vba=True` 保宏；`.xls` / `.ods` 不支持，请先另存 |
| 粘贴数据 | TSV、`｜` 分隔、Markdown 表格 | 按原分隔符拼回，列数对齐 |
| 纯文本片段 | 一行一值 | 逐行加密后原样输出 |

### 2.2 选列 / 选行：四种写法 + 三道护栏
```text
手机号          按列名（表头精确匹配）
C               按列字母
D:F             列区间
3               列号
--rows 2-100    行范围，可多个叠加
--no-header     连表头一起加密（默认保留表头）
```

**多工作表语义（重要，已固化进代码）**
- 按**列名**指定 → 某个 sheet 首行没有这个表头就**整页跳过**，记进 `skipped_sheets`，绝不猜"大概在第 4 列"。
- 按**列字母/列号**指定 → 视为显式跨表指令，对每个 sheet 都生效；某页没那么宽就剔除该列、该页原样跳过（不报错）。

**三道护栏（任一触发即中止，且不留下半截文件）**
1. 按列名指定，但**所有** sheet 首行都没找到 → `ValueError: 列名 ['手几号'] 在全部 3 个工作表首行都没找到 → 疑似列名拼错，已中止（未写出任何文件）`
2. 选列写法非法（`A1` / `9Z` / 空串）→ `ValueError: 选列写法非法 …：只接受 列名 / 列字母(C) / 区间(D:F) / 列号(3)`
3. **空转护栏** → 命中 0 个可加密单元格且无公式跳过时报警，防止"选错列却得到一个看似成功的文件"。对**幂等重跑**（源已是本工具密文）自动豁免：`allow_noop=True`。

### 2.3 算法：可逆 5 种 + 不可逆 2 种
| 模式 | 可逆 | 每格开销 | 什么时候选 | 注意 |
|---|---|---|---|---|
| `aesgcm` | ✅ | ≈120 B | **默认**，绝大多数场景 | AEAD，篡改 1 bit 必被拒 |
| `chacha20` | ✅ | ≈120 B | 无 AES-NI 的机器 | 同为 AEAD |
| `fernet` | ✅ | ≈130 B | 要和现有 Python/Django 服务互通 | 密文略长 |
| `sm4gcm` | ✅ | ≈120 B | 合规要求国密 | 需 `pip install gmssl`，缺依赖时**明确报错不静默降级** |
| `age` | ✅ | — | 外包/运营只能填不能看 | 需 `age` 或 `rage` 二进制，公钥加密 |
| `hmac` | ❌ | 固定 43 B | 两张表要按手机号 JOIN 但都不能有明文 | 同口令同摘要；手机号仅 10¹⁰ 量级，**可被字典爆破，属脱敏不属加密** |
| `mask` | ❌ | 长度不变 | 贴邮件/截图给人看 | 无密钥保护，纯展示 |

**密文格式**（5 段冒号分隔，自描述、可跨语言复现）：
```
enc:v1:<alg>:<base64 salt 16B>:<base64 nonce 12B>:<base64 ciphertext+tag>
```
- KDF：`Scrypt(n=2¹⁴, r=8, p=1)` → 32 B 密钥（SM4 取 16 B）
- **每格独立 salt + nonce** → 同一明文两次加密结果不同，杜绝"密文一样说明手机号一样"的等值泄露
- AEAD 关联数据固定 `b"DETv1"`；需要按列绑定防挪位时改成 `f"DETv1:{sheet}!{col}{row}"`
- **幂等**：已带 `enc:v1:` 的格子自动跳过，重跑不会产生 `enc:v1:aesgcm:enc:v1:…` 套娃
- 负数向用例全部抛异常而非返回半截垃圾：错口令 / 密文篡改 1 bit / 把明文当密文解密 / 垃圾串 / 未知算法

### 2.4 格式保全清单
**原样保留** ✅：字体（粗体/颜色/字号）、填充、边框、对齐、数字格式、列宽行高、合并单元格、工作表顺序与名称、冻结窗格、隐藏行列、数据验证、条件格式、单元格批注、`date`/`datetime` 原生类型、公式字符串、`.xlsm` 的 VBA 模块、其他 sheet 的全部内容。

**已知会丢，交付时必须主动告知** ⚠️：图表 / 透视表 / 切片器 / 迷你图、VBA 窗体与 ActiveX 控件（openpyxl 不保真 → 预检 `ws._charts` / `ws._pivots`，非空先提示留底或改走 CSV 通道）；被加密列**被公式引用**时该公式会变 `#VALUE!`（dry-run 阶段就要点名）；数字加密后类型变字符串（属预期，解密后按原 `number_format` 回转）。

### 2.5 交付前自检：22 条断言
`python tests/test_skill.py` 一次跑完，覆盖 Excel 逐格零差异 / 7 项样式保真 / CSV 的 BOM+CRLF+内嵌换行 / 粘贴列对齐 / 3 算法回环 / 空格跳过 / 幂等 / 行范围·列名·列区间·无表头四种选法 / 多 sheet 语义差异 / 4 类安全负向 / 同明文密文不同 / 源文件未被覆盖 / 解包 zip 校验输出 XML 里无明文残留。当前 **19 项通过 + 1 项跳过（SM4 需 gmssl）**，退出码 0。

---

## 3. 安装方法

### 3.1 作为 Skill 安装（推荐）
把整个目录放进 Agent 的 Skills 目录即可，**无需构建、无需注册**，Agent 会读 `SKILL.md` 的 frontmatter 自动发现：
```text
.github/skills/Data-Encryption-Tool/     # Codex / OpenAI Agents
└── .agents/skills/Data-Encryption-Tool/ # 其他兼容 Agent
```
目录里**没有任何业务脚本**——算法与断言以内联代码块形式写在 `SKILL.md` 第 2、4 节，Agent 按图执行，审计时直接对着 `tests/test_skill.py` 逐条比对即可。

### 3.2 依赖
```bash
pip install "cryptography>=42" openpyxl     # 必需（Python >= 3.10）
pip install gmssl                            # 仅 --alg sm4gcm 需要
brew install age        # 或 choco install age / apt install age   # 仅 --alg age 需要
```
依赖缺失时的行为是**明确报错并说明怎么装**，不会悄悄换成 AES 让你以为用了国密。

### 3.3 只用 CLI（不接 Agent）
`tests/test_skill.py` 自带 CLI 模式，可直接在 CI 里调：
```bash
export DET_PASSPHRASE='你的口令'    # 不进 shell 历史；不设则交互式 getpass 隐藏输入
python tests/test_skill.py --encrypt 客户表.xlsx -c 手机号 --alg aesgcm
python tests/test_skill.py --decrypt 客户表.encrypted.xlsx -c 手机号
python tests/test_skill.py --encrypt 订单.csv -c phone --alg hmac
python tests/test_skill.py            # 不带参数 = 跑 22 条自检
```

### 3.4 校验安装
```bash
python tests/test_skill.py | tail -4
# 通过 19 / 19   全绿 ✅  可交付。
```

---

## 4. 使用方法

### 4.1 对 Agent 说人话（主用法）
Agent 会先做一次 **dry-run 预演**——报出「将加密 N 格 / 跳过 M 个空格 / 跳过 K 个无此表头的 sheet / 其余 X 格不动」，等你确认才落盘：

| 你说 | Agent 做什么 |
|---|---|
| `把 客户表.xlsx 的手机号列加密` | `aesgcm` + 保留表头 → `客户表.encrypted.xlsx`，「配置」「说明」页整页跳过 |
| `D:F 三列用国密，只要第 2 到 100 行` | 先探测 `gmssl`，可用则 `sm4gcm`，不可用就明说并停下问你 |
| `这几行贴给你，第 2 列打码，我要贴回邮件` | `mask` 模式，按原 `\t` 拼回；并声明"掩码不可逆" |
| `两张表要按手机号关联，但都不能出现明文` | `hmac` 模式（同口令 → 同摘要 → 可 JOIN），同时提醒可被字典爆破 |
| `把 客户表.encrypted.xlsx 的手机号解回来` | 校验 `enc:v1:` 头 → 解密 → 逐格写回原样式 → `客户表.encrypted.DECRYPTED.xlsx` |
| `给外包填，他们只能写不能看` | 建议 `age` + 外包公钥；无 age 环境则退回 `aesgcm` 并说明"他们拿口令仍能解" |

### 4.2 五步工作流（写在 `SKILL.md` 第 1 节，Agent 强制走）
1. **认输入形态** —— 粘贴 / xlsx / xlsm / csv；CSV 先问编码（默认 `utf-8-sig`）与分隔符
2. **定范围** —— 列名 or 列字母、行范围、是否保留表头（默认保留）
3. **定算法与口令** —— 默认 `aesgcm`；用户说"脱敏不用还原"才切 `hmac`/`mask`，并**显式告知不可逆**
4. **dry-run 报数** —— 加密 N / 空 M / 跳页 K / 不动 X，等确认
5. **执行 + 跑 22 条断言** —— 全绿才交付，回复里必须含：算法、加密格数、跳过格数、`skipped_sheets`、`formulas_skipped`，以及一句"口令未存储，请自行试解一格验证"

### 4.3 Python 里当库用
```python
import os, sys
sys.path.insert(0, "Data-Encryption-Tool")
from tests.test_skill import process_xlsx, process_csv, decrypt, mask, one_way

pw = os.environb[b"DET_PASSPHRASE"]

# 加密：只动「手机号」列，其余连样式带公式原样
stat = process_xlsx("客户表.xlsx", "客户表.encrypted.xlsx", ["手机号"], pw, mode="aesgcm")
# {'encrypted': 2, 'empty': 3, 'skipped_sheets': ['配置', '说明'], 'formulas_skipped': 0}

# 解密回原格式（另一个文件名，源不动）
process_xlsx("客户表.encrypted.xlsx", "客户表.DECRYPTED.xlsx", ["手机号"], pw, mode="aesgcm",
             allow_noop=True)          # allow_noop：对已密文文件重跑不算空转

# CSV：字节进字节出，BOM / CRLF / 引号风格自动沿用
out_bytes, st = process_csv(open("订单.csv","rb").read(), ["phone"], pw, mode="aesgcm")
open("订单.encrypted.csv","wb").write(out_bytes)

print(decrypt("enc:v1:aesgcm:…", pw))  # → '13812345678'；口令错/被篡改 → ValueError
```
> ⚠️ `process_xlsx` 的第 5 个位置参数是 **`mode=`**，不是算法名——**务必用关键字传参**，写成 `process_xlsx(a, b, ["C"], pw, "aesgcm")` 才对；传错会得到 `ValueError: 未知算法: encrypt`。

### 4.4 安全边界（务必读）
- 这是**单元格级**加密，不是整文件加密。要防文件被拷走请用 BitLocker / VeraCrypt / 7-Zip AES，两层叠加最稳。
- 密文里**没有密钥备份**，口令丢了没有任何找回途径，本工具也不留 salt 台账。
- **换口令 = 全量解密再重加密**，需留双跑窗口；要免停机轮换，改用 `age` / KMS 封一层 KEK（DEK 不变，只重封）。
- 列内高重复值（性别/省份）且要抗等值泄露 → 保持每格独立 salt，**别为了性能优化掉**。
- 性能参考：`Scrypt(2¹⁴)` ≈ 30–60 ms/格，1 万格约 5–10 分钟属正常。要更快就把 KDF 提到文件级（salt = 文件 salt ‖ 行 ‖ 列），但那是 `enc:v2` 格式变更，需走变更记录。

---

## 5. 输入输出示例

以下**全部取自 `python tests/test_skill.py` 真实跑出的产物**（口令 `demo-pass-2026`，仅示例用，请勿复用）。

### 5.1 Excel：加密「手机号」列

**输入** `客户表.xlsx` → sheet `客户`（第 1 行 A1:E1 合并标题，第 2 行表头，数据第 3–6 行，第 7 行合计，冻结 `A3`）

| 客户号 | 姓名 | 手机号 | 额度 | 签约日 |
|---|---|---|---|---|
| C001 | **张三**（红+粗体） | 13812345678 | ¥12,000 | 2026-01-05 |
| C002 | **李四** | 15900001111 | ¥8,500 | 2026-02-11 |
| C003 | 王五 | *(空)* | ¥3,000 | 2026-03-02 |
| C004 | 赵六 | *(空串)* | ¥0 | *(空)* |
| | | | **=SUM(D2:D5)** ← D6 公式 | |

另有 `配置` 页（`B1=prod`）、`说明` 页——**按列名指定时这两页整页跳过**。

**输出** `客户表.encrypted.xlsx`
```
encrypted=2  empty=3  skipped_sheets=['配置', '说明']  formulas_skipped=0
```

| 客户号 | 姓名 | 手机号 | 额度 | 签约日 |
|---|---|---|---|---|
| C001 | **张三**（红+粗体，未变） | `enc:v1:aesgcm:auOL4vjYBsgEnjYVoV8TKA==:CeKoTA5IB9eTBQzc:eZ5zcpAbOb6V25vF2kUVVq5P2F3ONSaLKmrU` | ¥12,000 | 2026-01-05 |
| C002 | **李四** | `enc:v1:aesgcm:ehDy8OjOR15H+277mWdFqw==:O25fo+seL9vyBn9/:5Rz6T//uxeN7BN1LntgkAYtYoudiGpbWYJXr` | ¥8,500 | 2026-02-11 |
| C003 | 王五 | *(仍 None，不产密文)* | ¥3,000 | 2026-03-02 |
| C004 | 赵六 | *(仍空串，类型不变)* | ¥0 | *(空)* |
| | | | **=SUM(D2:D5)** ← 公式原样，`data_type='f'` | |

**逐项核对（测试断言 1–8）**：sheet 顺序 `[客户, 配置, 说明]` 未变 ✅；`A1:E1` 合并区未变 ✅；C 列红字加粗 + 列宽 16 未变 ✅；D 列 `"#,##0"元"`、E 列 `yyyy-mm-dd` **连密文格都保住了数字格式** ✅；E2 仍是真 `datetime` 而非字符串 ✅；D6 公式字符串与 `data_type='f'` 未变 ✅；`配置!B1` 仍为 `prod` ✅；未选中的 **26 个格逐格 `==` 比较零差异** ✅。

**解密回环** → `C2` 回到 `13812345678`，与原件 6×5 全格比对零差异；同明文再加密一次得到的是**另一串密文**（salt 不同）。

### 5.2 CSV：BOM + 内嵌换行 + 字段内逗号

**输入** `订单.csv`（`utf-8-sig`，CRLF，注意 `张,三` 和跨行的 note）
```csv
name,phone,note
"张,三",13812345678,"line1
line2"
李四,15900001111,ok
```

**输出** `订单.encrypted.csv` —— BOM 有且仅有一份，行尾仍是 CRLF（3 个行尾 + 1 个内嵌 = 4），`"张,三"` 与单元格内换行都没错行：
```csv
name,phone,note
"张,三",enc:v1:aesgcm:8Yy4C6JQM4BVE4tZkHjnOw==:HrFSOz1TUi0j+4QT:tbuOfjNHMRTKqnS6wsPu+W/p3Jk3RpajAxtS,"line1
line2"
李四,enc:v1:aesgcm:Vi0YkFiyuImJXbvNqZGG7Q==:BJjiEzrL6tLiS92H:Yurm5DUbx1Eoe8G07/mGVocCn9Sd4pHV9h6h,ok
```
> 反例：用 `line.split(",")` 手搓的脚本在这里必然把 `张` 和 `三` 劈成两列，并把 `line2"` 当成新行——本工具全程走 `csv` 模块的 dialect。

### 5.3 粘贴数据：TSV 进、TSV 出

**输入**（从邮件/网页直接复制）
```
姓名	手机号	城市
张三	13812345678	太原
李四	15900001111	西安
王五		北京
```

**输出 A · `mask` 模式**（贴回邮件用，不可逆）
```
姓名	手机号	城市
张三	138****5678	太原
李四	159****1111	西安
王五		北京		← 空保持空，不出现 ****
```
顺带支持的掩码规则：身份证 `110105199003071234 → 110105********1234`；邮箱 `a.b@corp.com → a***b@corp.com`；≤4 字符全打星；非结构化文本保留首尾各 1 字。

**输出 B · `aesgcm` 模式**（可逆，表头不加密、列数对齐）
```
姓名	手机号	城市
张三	enc:v1:aesgcm:auOL4vjYBsgEnjYVoV8TKA==:CeKoTA5IB9eTBQzc:eZ5zcpAbOb6V25vF2kUVVq5P2F3ONSaLKmrU	太原
李四	enc:v1:aesgcm:ehDy8OjOR15H+277mWdFqw==:O25fo+seL9vyBn9/:5Rz6T//uxeN7BN1LntgkAYtYoudiGpbWYJXr	西安
王五		北京
```

**输出 C · `hmac` 模式**（两表可 JOIN，不可逆）
```
张三	hmac:v1:980d0c85cd70e41c7c6be0ef6288c04f079cb646c93	太原
```
同一口令下 `客户表` 和 `订单表` 的 `13812345678` 会得到**同一个摘要**，可以直接 `JOIN ... ON hmac_phone`；换口令或换一个数字，摘要整个变。

### 5.4 CLI 端到端实录
```bash
$ export DET_PASSPHRASE=demo-pass-2026
$ python tests/test_skill.py --encrypt 客户表.xlsx -c 手机号 --alg aesgcm
{"out": "客户表.encrypted.xlsx", "encrypted": 2, "empty": 3,
 "skipped_sheets": ["配置", "说明"], "formulas_skipped": 0, "alg": "aesgcm",
 "note": "口令未存储；请用 --decrypt 试解一格验证"}

$ python tests/test_skill.py --decrypt 客户表.encrypted.xlsx -c 手机号
{"out": "客户表.encrypted.DECRYPTED.xlsx", "decrypted": 2}

$ python tests/test_skill.py --encrypt 客户表.xlsx -c 手几号
ValueError: 列名 ['手几号'] 在全部 3 个工作表首行都没找到 → 疑似列名拼错，已中止（未写出任何文件）

$ python tests/test_skill.py --encrypt 客户表.xlsx -c A1
ValueError: 选列写法非法 ['A1']：只接受 列名 / 列字母(C) / 区间(D:F) / 列号(3)
```

### 5.5 负向用例：这些都必须**报错**，不能返回半截数据
| 场景 | 结果 |
|---|---|
| 口令输错 | `ValueError: 解密失败：口令错误 / 密文被篡改 / 算法不匹配（aesgcm）` |
| 密文篡改 1 bit（改最后一个字节） | 同上，GCM tag 校验失败 |
| 把明文 `13812345678` 当密文喂进去 | `ValueError: 这不是本工具产出的密文（可能已是明文），拒绝处理` |
| AAD 换成 `b"OTHER-CTX"` 解 | 抛异常 —— 密文挪到别的列/别的用途就解不开 |
| 对已加密文件重跑 | 幂等跳过，`encrypted=0`；`allow_noop=True` 时正常返回而非报警 |
| 选了 `Z:AA`（三页都没这么宽） | 默认抛空转护栏；`allow_noop=True` 时返回 `skipped_sheets=[客户,配置,说明]` |

---

## 目录

```
Data-Encryption-Tool/
├── SKILL.md                    # 唯一权威说明书：工作流 / 核心代码 / 22 条断言
├── agents/openai.yaml          # 界面元数据、依赖声明、文件权限与 5 条护栏
├── references/
│   ├── format-fidelity.md      # 格式保全清单 + openpyxl 已知丢点 + CSV 三坑
│   └── algorithm-params.md     # 各算法 KDF/nonce/tag 参数、密文体积、互通与轮换
├── tests/
│   └── test_skill.py           # 22 条可执行用例 + 可用 CLI（无额外测试框架依赖）
└── README.md
```

## License

MIT —— 随便用、随便改、随便商用；改了加密算法或密文格式请把变更写进 `references/algorithm-params.md`，那是别人能否解你历史数据的唯一依据。
