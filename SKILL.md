---
name: Data-Encryption-Tool
version: 0.1.0
description: >-
  对用户粘贴的表格数据，或本地 Excel（.xlsx/.xlsm）、CSV/TSV 文件，按「列」或「行」
  做可逆加密与解密。v0.1.0 已实现的算法：AES-256-GCM（默认）、ChaCha20-Poly1305、
  Fernet、国密 SM4-GCM（需本地有 gmssl）；另有单向 HMAC-SHA256 脱敏与掩码两种
  「不可解密」模式。AES-256-CBC 与 AGE 公钥模式属规划项，本版未实现（调用抛未知算法）。核心保证：未选中的单元格逐格零改动，字体/颜色/合并/
  列宽/数字格式/日期类型/公式/工作表顺序/CSV 的 BOM 与换行符全部原样保留；结果另存为
  原名.encrypted.xlsx，绝不覆盖原件；口令不写入文件、日志与对话记录。
  触发词：加密这一列、把手机号脱敏、密文导出、解密回去、列加密、excel 加密某几列。
---

# Data-Encryption-Tool

把「哪几列 / 哪几行要加密」这件事，做成一个不破坏原表格式的可逆操作。

---

## 0. 三条不可协商的底线

| # | 规则 | 违反时的表现 |
|---|---|---|
| 1 | **只碰选中的格子**。其余单元格一律 `cell.value` 原样写回，不重设任何 style | 用户对比原件发现格式漂了 = 交付失败 |
| 2 | **绝不覆盖源文件**。输出 `原名.encrypted.<ext>`，同名则加 `_2`、`_3` | 覆盖 = 数据事故 |
| 3 | **口令只从用户当次输入取**，不落盘、不写进 YAML、不写进日志、不复述在回复里 | 出现在任何文件里 = 泄密 |

> 密文里**不含口令**，也不含可推导口令的信息。密文格式：`enc:v1:<alg>:<b64salt>:<b64nonce>:<b64ct>`。每格独立 salt+nonce，同一明文两次加密结果不同（防等值比对）。

---

## 1. 工作流（每步都有验收门）

1. **确认输入形态**：粘贴文本 / `.xlsx` / `.xlsm` / `.csv` / `.tsv`。CSV 先问编码（默认 `utf-8-sig`）与分隔符。
2. **确认范围**：让用户指出列名或列字母（`C`、`手机号`）、行范围（`2-100`）、是否**保留表头不加密**（默认保留）。
3. **确认算法与口令**：默认 `aesgcm`；用户只说「脱敏不用还原」时改用 `hmac` 或 `mask`（并明确告知**这两种解不回来**）。
4. **预演（dry-run）**：先只统计「将加密 N 格，跳过 M 个空格，未选中 K 格」，把数字报给用户确认，再动手。
5. **执行 + 自检**：跑第 4 节的断言套件，全绿才交付。
6. **交付**：给出输出文件路径、算法名、加密格数、跳过格数；**附一句「口令请自行保管，本工具不存储」**；并提示用户自己试解密一格验证。

---

## 2. 核心代码（唯一代码片段，Excel / CSV / 粘贴三用）

```python
# ---- 算法层：依赖仅 cryptography（Python>=3.10）----
import os, re, json, base64, hashlib, hmac as _hmac

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM, ChaCha20Poly1305
    from cryptography.hazmat.primitives.kdf.scrypt import Scrypt
except ImportError:  # pragma: no cover
    raise SystemExit("需要: pip install cryptography")

HEAD = re.compile(r"^enc:v1:(?P<alg>[a-z0-9]+):(?P<salt>[^:]+):(?P<nonce>[^:]+):(?P<ct>[^:]+)$")

def _key(pw: bytes, salt: bytes, n=32) -> bytes:
    return Scrypt(salt=salt, length=n, n=2**14, r=8, p=1).derive(pw)

def _b(b): return base64.b64encode(b).decode()
def _u(s): return base64.b64decode(s)

def encrypt(plain: str, pw: bytes, alg="aesgcm") -> str:
    salt, nonce = os.urandom(16), os.urandom(12)
    if alg == "aesgcm":
        ct = AESGCM(_key(pw, salt)).encrypt(nonce, plain.encode(), b"DETv1")
    elif alg == "chacha20":
        ct = ChaCha20Poly1305(_key(pw, salt)).encrypt(nonce, plain.encode(), b"DETv1")
    elif alg == "fernet":
        from cryptography.fernet import Fernet
        k = base64.urlsafe_b64encode(_key(pw, salt, 32))   # Fernet 要 urlsafe base64 的 32B 密钥
        ct = Fernet(k).encrypt(plain.encode())
        return f"enc:v1:fernet:{_b(salt)}:{_b(nonce)}:{_b(ct)}"   # nonce 位占位，保持 5 段同构
    elif alg == "sm4gcm":                                        # 国密，需 pip install gmssl
        from gmssl.sm4 import CryptSM4, SM4_GCM
        c = CryptSM4(); c.set_key(_key(pw, salt, 16), SM4_GCM)
        ct, tag = c.encrypt_gcm(plain.encode(), nonce, b"DETv1")
        ct = ct + tag
    else:
        raise ValueError(f"未知算法: {alg}")
    return f"enc:v1:{alg}:{_b(salt)}:{_b(nonce)}:{_b(ct)}"

def decrypt(tok: str, pw: bytes) -> str:
    m = HEAD.match(tok.strip())
    if not m:
        raise ValueError("这不是本工具产出的密文（可能已是明文），拒绝处理")
    alg, salt, nonce, ct = m["alg"], _u(m["salt"]), _u(m["nonce"]), _u(m["ct"])
    try:
        if alg == "aesgcm":   return AESGCM(_key(pw, salt)).decrypt(nonce, ct, b"DETv1").decode()
        if alg == "chacha20": return ChaCha20Poly1305(_key(pw, salt)).decrypt(nonce, ct, b"DETv1").decode()
        if alg == "fernet":
            from cryptography.fernet import Fernet
            return Fernet(base64.urlsafe_b64encode(_key(pw, salt, 32))).decrypt(ct).decode()
        if alg == "sm4gcm":
            from gmssl.sm4 import CryptSM4, SM4_GCM
            c = CryptSM4(); c.set_key(_key(pw, salt, 16), SM4_GCM)
            return c.decrypt_gcm(ct[:-16], nonce, b"DETv1", ct[-16:]).decode()
    except Exception as e:
        raise ValueError(f"解密失败：口令错误 / 密文被篡改 / 算法不匹配（{alg}）") from e
    raise ValueError(f"未知算法: {alg}")

def one_way(plain: str, pw: bytes) -> str:      # 不可逆：同一明文同一口令 → 同一摘要，可用于 JOIN
    d = _hmac.new(_key(pw, b"det-hmac-pepper", 32), plain.encode(), hashlib.sha256).hexdigest()
    return f"hmac:v1:{d[:43]}"

def mask(plain: str) -> str:                    # 不可逆：仅展示用
    s = str(plain)
    if len(s) <= 4: return "*" * len(s)
    if re.fullmatch(r"\d{11}", s):  return s[:3] + "****" + s[-4:]
    if re.fullmatch(r"\d{15,18}", s): return s[:6] + "********" + s[-4:]
    if "@" in s:
        u, d = s.split("@", 1); return (u[0] + "***" + u[-1:] + "@" + d) if len(u) > 2 else "***@" + d
    return s[0] + "*" * (len(s) - 2) + s[-1:]

MODES = {"aesgcm": "AES-256-GCM（默认，推荐）",
         "chacha20": "ChaCha20-Poly1305", "fernet": "Fernet",
         "sm4gcm": "SM4-GCM（国密，需 gmssl）", "hmac": "HMAC-SHA256（单向，不可解密）",
         "mask": "掩码（单向，不可解密）"}
# 规划中、v0.1.0 未实现（调用会抛 ValueError 未知算法）：
#   aesgcm-cbc → AES-256-CBC；age → AGE/X25519 公钥（需 age 二进制）

# ---- Excel 层：openpyxl，只改 value，不碰 style ----
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string as col2i

def _is_named(cols):   # 列名（非列字母/列号）→ 按表头匹配；列字母 → 显式跨 sheet 指令
    return any(not (re.fullmatch(r"[A-Z]{1,3}(:[A-Z]{1,3})?", str(c)) or str(c).isdigit()) for c in cols)

def resolve_cols(ws, spec, header_row=1):      # spec: 单个 "C" / "手机号" / "D:F" / 3
    out = []
    for s in spec:
        if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", s):
            a, b = s.split(":"); out += list(range(col2i(a), col2i(b) + 1))
        elif re.fullmatch(r"[A-Z]{1,3}", s):
            out.append(col2i(s))
        else:
            hit = [c for c in range(1, ws.max_column + 1)
                   if str(ws.cell(header_row, c).value).strip() == s.strip()]
            if not hit: missing.append((ws.title, s))   # 该页无此表头 → 整页跳过，绝不猜列
            out += hit
    out = sorted(set(out))
    if _is_named([spec[0]]) and not out:
        raise ValueError(f"sheet「{ws.title}」第 {header_row} 行找不到列名 {spec}")
    return [c for c in out if c <= ws.max_column]   # 列字母超出该页宽度 → 空列表 → 该页原样跳过

COL_SPEC = re.compile(r"^[A-Z]{1,3}(:[A-Z]{1,3})?$|^\d+$")   # 列字母 / 区间 / 列号

def process_xlsx(src, dst, cols, pw, mode="encrypt", rows=None, header_row=1, keep_header=True,
                 allow_noop=False):   # allow_noop=True 用于「重跑/校验」场景，允许 0 加密
    cols = [str(c).strip() for c in cols]
    bad = [c for c in cols if not c or (not COL_SPEC.match(c) and c.isascii())]
    if bad:
        raise ValueError(f"选列写法非法 {bad}：只接受 列名 / 列字母(C) / 区间(D:F) / 列号(3)，"
                         f"不要写单元格坐标(A1) 或带数字的列字母(9Z)")   # ★先拦语法，再谈匹配
    wb = load_workbook(src, keep_vba=str(src).endswith("xlsm"), data_only=False)
    stat = {"encrypted": 0, "empty": 0, "skipped_sheets": [], "formulas_skipped": 0}
    missing = []
    for ws in wb.worksheets:
        try:
            idx = resolve_cols(ws, cols, header_row)
        except ValueError:
            # 该 sheet 首行没有这些列名（如"配置"页）→ 整页跳过，绝不猜列
            if all(_is_named(c) for c in cols):
                missing.append(ws.title); stat["skipped_sheets"].append(ws.title); continue
            raise                       # 按列字母/列号却解析失败 = 真错误
        if not idx:
            stat["skipped_sheets"].append(ws.title); continue   # 列字母超出该页宽度
        for r in range((header_row + 1) if keep_header else 1, ws.max_row + 1):
            if not _in(r, rows): continue
            for c in idx:
                cell = ws.cell(r, c); v = cell.value
                if v is None or v == "": stat["empty"] += 1; continue
                if isinstance(v, str) and HEAD.match(v): stat["empty"] += 1; continue   # 幂等
                if isinstance(v, str) and v.startswith("="):
                    stat["formulas_skipped"] += 1; continue      # ★公式护栏：公式不是数据
                v = str(v)
                if mode == "mask":   cell.value = mask(v)
                elif mode == "hmac": cell.value = one_way(v, pw)
                else:                cell.value = encrypt(v, pw, mode)
                stat["encrypted"] += 1
    # ★硬护栏：按列名指定却在所有 sheet 都未命中 → 多半拼错，宁可中止也不交"啥也没加密"的文件
    if missing and len(missing) == len(wb.worksheets):
        raise ValueError(f"列名 {cols} 在全部 {len(missing)} 个工作表首行都没找到 → 疑似列名拼错，已中止（未写出任何文件）")
    # ★空转护栏：宁可报错，也不交出"看着成功、其实一格没锁"的文件
    if not allow_noop and stat["encrypted"] == 0 and stat["formulas_skipped"] == 0:
        raise ValueError(f"选列 {cols} / 行范围 {rows} 命中 0 个可加密单元格（空 {stat['empty']}、"
                         f"跳过页 {stat['skipped_sheets']}）→ 疑似选错，已中止（未写出任何文件）")
    wb.save(dst); return stat


def _in(r, specs):
    for s in specs:
        if isinstance(s, tuple):  r0, r1 = s
        elif "-" in str(s):       r0, r1 = map(int, s.split("-"))
        else:                     r0 = r1 = int(s)
        if r0 <= r <= r1: return True
    return False
```

CSV / 粘贴数据的差异只有**读写方式**，算法层完全复用：

```python
import csv, io, codecs

def process_csv(raw_bytes, cols, pw, mode="encrypt", keep_header=True, dialect=None):
    """raw_bytes: 文件原始字节。返回 (输出字节, stat)。BOM/换行符/引号风格全部沿用。"""
    bom = raw_bytes[:3] == codecs.BOM_UTF8
    text = raw_bytes.decode("utf-8-sig")                 # 读时剥 BOM，否则 BOM 会粘进第一个列名
    d = dialect or csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    rows = list(csv.reader(io.StringIO(text, newline=""), d))
    head = rows[0] if keep_header and rows else []
    idx = [i for i, name in enumerate(head) if name.strip() in cols] if keep_header else \
          [col2i(c) - 1 for c in cols if re.fullmatch(r"[A-Z]{1,3}", c)]
    stat = {"encrypted": 0, "empty": 0}
    for r in rows[int(keep_header):]:
        for i in idx:
            if i >= len(r) or r[i] == "": stat["empty"] += 1; continue
            if HEAD.match(r[i]):           stat["empty"] += 1; continue
            r[i] = encrypt(r[i], pw, mode) if mode != "hmac" else one_way(r[i], pw)
            stat["encrypted"] += 1
    out = io.StringIO(newline=""); w = csv.writer(out, d, lineterminator=d.lineterminator)
    w.writerows(rows)
    body = out.getvalue().encode("utf-8")
    return ((codecs.BOM_UTF8 + body) if bom else body), stat   # 有 BOM 就原样写回，且只写一份

# 粘贴数据：先 csv.reader(io.StringIO(粘贴文本)) 解析成 rows，跑同一套 idx 逻辑，
# 再用 "\t".join / " | ".join 按原分隔符拼回，首行原样当表头输出。
```

### 算法怎么选（给用户看的那张表）

| 模式 | 可逆 | 何时选 | 注意 |
|---|---|---|---|
| `aesgcm` | ✅ | **默认**。单文件、单人口令 | 带 GCM 认证标签，篡改必被拒 |
| `chacha20` | ✅ | 无 AES-NI 的机器 / 偏好流加密 | 同样带认证 |
| `fernet` | ✅ | 要和现有 Python 服务互通 | 密文略长 |
| `sm4gcm` | ✅ | 合规要求国密 | 先 `python -c "import gmssl"` 探测，没有就明说并回退 aesgcm，**不要静默换** |
| `age` | ⏳ | 多人只加密不解密（给运营/外包填表） | **v0.1.0 未实现**（规划 v0.2）；需 `age` CLI，公钥加密、本人私钥解 |
| `hmac` | ❌ | 要拿脱敏列去做 JOIN / 去重 | 同明文同摘要，可被字典猜，别当加密用 |
| `mask` | ❌ | 只给人眼看 | 纯展示，无任何密钥保护 |

---

## 3. 格式保全：为什么「只改 value」就够

openpyxl 的 `load_workbook` → 改 `cell.value` → `save` 会保留：字体/填充/边框/对齐/数字格式/列宽行高/合并单元格/工作表顺序/冻结窗格/数据验证/条件格式/命名样式/其他 sheet 内容。

**已知会丢的东西，必须在交付时主动告知用户**：
- Excel **图表、透视表、切片器、VBA 里的窗体控件**（openpyxl 不保真）→ 检测到 `ws._charts` / `ws._pivots` 非空时，先提醒「建议先把该 sheet 复制一份留底」，或建议改走 CSV 通道。
- `.xlsm` 宏：用 `keep_vba=True` 可保住宏模块，但仍提示用户在副本上点一次「启用宏」验证。
- 公式：`data_only=False` 读回的是公式字符串本身（不是缓存值），所以 `=SUM(D2:D4)` 原样保留 ✅；但被加密列若被公式引用，公式结果会变成 `#VALUE!` —— **预演时必须检查并提示**。
- `.xls`（老格式）不支持 → 让用户先另存为 .xlsx。

CSV 侧三个坑（都在代码里处理了）：**BOM 只能有一份**、**`lineterminator` 必须沿用原文的 `\r\n` 或 `\n`**、**引号风格交给 `csv` 模块**，手写 `split(",")` 一定会在 `"张,三"` 上错行。

---

## 4. 交付前自检套件（必须全绿）

```python
def assert_roundtrip(src, dst, ws_name, cols, pw, mode="aesgcm", header_row=1):
    a, b = load_workbook(src), load_workbook(dst)
    aw, bw = a[ws_name], b[ws_name]
    assert a.sheetnames == b.sheetnames, "工作表顺序/数量变了"
    assert {str(x) for x in aw.merged_cells.ranges} == {str(x) for x in bw.merged_cells.ranges}, "合并区域变了"
    idx = resolve_cols(bw, cols, header_row)
    n = 0
    for r in range(1, aw.max_row + 1):
        for c in range(1, max(aw.max_column, bw.max_column) + 1):
            va, vb = aw.cell(r, c).value, bw.cell(r, c).value
            if c in idx and r > header_row and va not in (None, ""):
                assert HEAD.match(str(vb)), f"({r},{c}) 没变成密文"
                if mode not in ("hmac", "mask"): assert decrypt(str(vb), pw) == str(va)
                n += 1
            else:
                assert va == vb, f"({r},{c}) 未选中单元格被改动: {va!r} -> {vb!r}"
                sa, sb = aw.cell(r, c), bw.cell(r, c)
                assert sa.font.b == sb.font.b and sa.font.color.rgb == sb.font.color.rgb
                assert sa.number_format == sb.number_format, f"({r},{c}) 数字格式变了"
                assert aw.cell(r, c).data_type == bw.cell(r, c).data_type, f"({r},{c}) 单元格类型变了"
    for col in "ABCDEFGHIJ":
        assert aw.column_dimensions[col].width == bw.column_dimensions[col].width, f"{col} 列宽变了"
    # 负向用例：错口令 / 篡改 1 bit / 明文当密文，三者都必须抛异常
    sample = [bw.cell(r, idx[0]).value for r in range(header_row + 1, bw.max_row + 1)
              if bw.cell(r, idx[0]).value][0]
    for label, fn in [("错口令", lambda: decrypt(sample, b"wrong")),
                      ("明文当密文", lambda: decrypt("13812345678", pw))]:
        try: fn(); raise AssertionError(f"负向用例未拦截: {label}")
        except ValueError: pass
    import base64 as _b64
    p = sample.split(":"); buf = bytearray(_b64.b64decode(p[-1])); buf[-1] ^= 1
    tampered = ":".join(p[:-1] + [_b64.b64encode(bytes(buf)).decode()])
    try: decrypt(tampered, pw); raise AssertionError("篡改 1bit 未拦截")
    except ValueError: pass
    assert encrypt("同一明文", pw) != encrypt("同一明文", pw), "同明文密文相同，存在等值泄露"
    return n   # 返回加密格数，写进交付说明
```

**8 条断言速查**：① sheet 顺序 ② 合并区域 ③ 密文都带 `enc:v1:` ④ 未选中格逐格相等 ⑤ 字体/颜色 ⑥ 列宽 ⑦ 数字格式与日期/公式类型 ⑧ 三个负向用例全被拒 + 同明文密文不同。CSV 再补 3 条：BOM 有且仅一份、行尾数 = 数据行 + 内嵌换行数、`"张,三"` 不错行。

> ⚠️ 本节是**人工核对模板**：`tests/test_skill.py` 的自动断言目前覆盖 ① ② ③ ④（值维度）与 ⑧，**未覆盖 ⑤ ⑥ ⑦**（编号 4–7 为空号，计划 v0.1.1 补齐），交付前须按本段手工比对字体/列宽/数字格式。

---

## 5. Examples

**E1** 用户：「把 `客户表.xlsx` 的 C 列手机号加密，其他别动。」
→ dry-run 报「将加密 2 格、跳过 4 个空格、其余 26 格不动」→ 确认 → 输出 `客户表.encrypted.xlsx` → 跑 8 条断言 → 回复：`aesgcm / 加密 2 格 / 空格 4 个原样保留 / 红字与合并单元格已核对 / 口令未存储，请用 decrypt 试解一格`。

**E2** 用户：「这几行贴给你，第 2 列打码，我要贴回邮件。」
→ 走粘贴通道，`mode=mask`，按原 ` | ` 对齐输出，`13812345678 → 138****5678`，并声明「掩码不可逆，需要还原请改用 aesgcm」。

**E3** 用户：「导给外包填，他们只能写不能看。」
→ v0.1.0 未实现 `mode=age`，只能退回 `aesgcm` 并明确说「这版他们拿口令仍能解；`age` 公钥模式计划 v0.2」。

**E4** 用户：「两张表要按手机号关联，但都不能出现明文。」
→ `mode=hmac`，两张表同一口令 → 摘要一致可 JOIN；同时提醒「HMAC 可被手机号字典爆破，属于脱敏不是加密」。

---

## 6. Scope（不做什么）

- 不做**整文件/整盘**加密（那是 VeraCrypt、BitLocker、7-Zip AES 的活）；本工具只做「单元格级」。
- 不做**密钥管理**：不管 KMS、不管口令轮换、不留 salt 台账。要审计就改用 `age`/KMS 外包密钥。
- 不处理 `.xls`、`.ods`、密码保护的 workbook（先让用户解密另存）。
- 不加密**表头**（默认），不加密**已是 `enc:v1:` 的格子**（幂等跳过），**不加密公式**（`=` 开头，默认跳过并计入 `formulas_skipped`；要连公式一起锁请显式说，会先把公式算成值）。
- 按**列名**指定时若所有 sheet 都没这个表头 → **直接报错不产出文件**（防「手机号」写成「手几号」后静默什么都不做）。
- 不用 `mask`/`hmac` 冒充加密；用户说「脱敏」时必须追问一句「以后还要不要还原」。
- 不初始化 Git、不推送远端，除非用户明确要求。
