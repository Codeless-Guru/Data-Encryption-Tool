#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Data-Encryption-Tool 自检套件：22 条用例。
用法:  python tests/test_skill.py                 # 跑全部断言（自建 fixture）
       python tests/test_skill.py --encrypt F.xlsx -c 3 --alg aesgcm   # 当 CLI 用
       python tests/test_skill.py --decrypt F.encrypted.xlsx -c 3
口令取自环境变量 DET_PASSPHRASE，未设置时用测试专用口令（仅 fixture 模式）。
"""
import os, re, sys, csv, io, json, base64, shutil, hashlib, datetime, argparse, tempfile
import hmac as _hmac
from pathlib import Path

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM, ChaCha20Poly1305
    from cryptography.hazmat.primitives.kdf.scrypt import Scrypt
    from cryptography.fernet import Fernet
except ImportError:
    sys.exit("pip install cryptography")

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import column_index_from_string as col2i
except ImportError:
    sys.exit("pip install openpyxl")

# ==================== 以下 3 个函数与 SKILL.md「核心代码」逐字一致 ====================
HEAD = re.compile(r"^enc:v1:(?P<alg>[a-z0-9]+):(?P<salt>[^:]+):(?P<nonce>[^:]+):(?P<ct>[^:]+)$")
_b = lambda x: base64.b64encode(x).decode()
_u = lambda s: base64.b64decode(s)

def _key(pw, salt, n=32):
    return Scrypt(salt=salt, length=n, n=2**14, r=8, p=1).derive(pw)

def encrypt(plain, pw, alg="aesgcm"):
    salt, nonce = os.urandom(16), os.urandom(12)
    if alg == "aesgcm":
        ct = AESGCM(_key(pw, salt)).encrypt(nonce, plain.encode(), b"DETv1")
    elif alg == "chacha20":
        ct = ChaCha20Poly1305(_key(pw, salt)).encrypt(nonce, plain.encode(), b"DETv1")
    elif alg == "fernet":
        k = base64.urlsafe_b64encode(_key(pw, salt, 32))
        ct = Fernet(k).encrypt(plain.encode())
        return f"enc:v1:fernet:{_b(salt)}:{_b(nonce)}:{_b(ct)}"
    elif alg == "sm4gcm":
        from gmssl.sm4 import CryptSM4, SM4_GCM
        c = CryptSM4(); c.set_key(_key(pw, salt, 16), SM4_GCM)
        ct, tag = c.encrypt_gcm(plain.encode(), nonce, b"DETv1")
        ct = ct + tag
    else:
        raise ValueError(f"未知算法: {alg}")
    return f"enc:v1:{alg}:{_b(salt)}:{_b(nonce)}:{_b(ct)}"

def decrypt(tok, pw):
    m = HEAD.match(tok.strip())
    if not m:
        raise ValueError("这不是本工具产出的密文（可能已是明文），拒绝处理")
    alg, salt, nonce, ct = m["alg"], _u(m["salt"]), _u(m["nonce"]), _u(m["ct"])
    try:
        if alg == "aesgcm":
            return AESGCM(_key(pw, salt)).decrypt(nonce, ct, b"DETv1").decode()
        if alg == "chacha20":
            return ChaCha20Poly1305(_key(pw, salt)).decrypt(nonce, ct, b"DETv1").decode()
        if alg == "fernet":
            return Fernet(base64.urlsafe_b64encode(_key(pw, salt, 32))).decrypt(ct).decode()
        if alg == "sm4gcm":
            from gmssl.sm4 import CryptSM4, SM4_GCM
            c = CryptSM4(); c.set_key(_key(pw, salt, 16), SM4_GCM)
            return c.decrypt_gcm(ct[:-16], nonce, b"DETv1", ct[-16:]).decode()
    except Exception as e:
        raise ValueError(f"解密失败：口令错误 / 密文被篡改 / 算法不匹配（{alg}）") from e
    raise ValueError(f"未知算法: {alg}")

def one_way(plain, pw):
    d = _hmac.new(_key(pw, b"det-hmac-pepper", 32), plain.encode(), hashlib.sha256).hexdigest()
    return f"hmac:v1:{d[:43]}"

def mask(plain):
    s = str(plain)
    if len(s) <= 4: return "*" * len(s)
    if re.fullmatch(r"\d{11}", s): return s[:3] + "****" + s[-4:]
    if re.fullmatch(r"\d{15,18}", s): return s[:6] + "********" + s[-4:]
    if "@" in s:
        u, d = s.split("@", 1); return (u[0] + "***" + u[-1:] + "@" + d) if len(u) > 2 else "***@" + d
    return s[0] + "*" * (len(s) - 2) + s[-1:]
# ============================ 说明书代码结束 ============================

def resolve_cols(ws, spec, header_row=1):  # noqa: F811  (与 SKILL.md 内联版一致)
    out = []
    for s in spec:
        s = str(s)
        if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", s):
            a, b = s.split(":"); out += list(range(col2i(a), col2i(b) + 1))
        elif re.fullmatch(r"[A-Z]{1,3}", s):
            out.append(col2i(s))
        elif s.isdigit():
            out.append(int(s))
        else:
            hit = [c for c in range(1, ws.max_column + 1)
                   if str(ws.cell(header_row, c).value).strip() == s.strip()]
            if not hit: raise ValueError(f"sheet「{ws.title}」第 {header_row} 行找不到列名「{s}」")
            out += hit
    out = sorted({c for c in out if c <= ws.max_column})   # 列字母超出该页宽度 → 剔除 → 该页原样跳过
    if not out and any(not (re.fullmatch(r"[A-Z]{1,3}(:[A-Z]{1,3})?", str(x)) or str(x).isdigit()) for x in spec):
        raise ValueError(f"sheet「{ws.title}」第 {header_row} 行找不到列名 {spec}")
    return out

def _in(r, specs):
    if not specs: return True
    for s in specs:
        s = str(s)
        if "-" in s:
            a, b = map(int, s.split("-"))
        else:
            a = b = int(s)
        if a <= r <= b: return True
    return False

def _is_named(col): return not re.fullmatch(r"[A-Z]{1,3}(:[A-Z]{1,3})?|\d+", str(col))

COL_SPEC = re.compile(r"^[A-Z]{1,3}(:[A-Z]{1,3})?$|^\d+$")

def process_xlsx(src, dst, cols, pw, mode="encrypt", rows=None, header_row=1, keep_header=True,
                 allow_noop=False):   # allow_noop=True 用于「重跑/校验」场景，允许 0 加密
    cols = [str(c).strip() for c in cols]
    bad = [c for c in cols if not c or (not COL_SPEC.match(c) and c.isascii())]
    if bad:
        raise ValueError(f"选列写法非法 {bad}：只接受 列名 / 列字母(C) / 区间(D:F) / 列号(3)，"
                         f"不要写单元格坐标（A1）或带数字的列字母（9Z）")
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
    # ★空转护栏：一格都没加密且不全是空格原因 → 多半选列/行范围写错，拒绝交出"看着成功其实啥也没干"的文件
    if not allow_noop and stat["encrypted"] == 0 and stat["formulas_skipped"] == 0:
        raise ValueError(f"选列 {cols} / 行范围 {rows} 命中 0 个可加密单元格"
                         f"（空 {stat['empty']}、跳过页 {stat['skipped_sheets']}）→ 疑似选错，已中止（未写出任何文件）")
    wb.save(dst); return stat


def process_csv(raw_bytes, cols, pw, mode="encrypt", keep_header=True):
    import codecs
    bom = raw_bytes[:3] == codecs.BOM_UTF8
    text = raw_bytes.decode("utf-8-sig")
    d = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    rows = list(csv.reader(io.StringIO(text, newline=""), d))
    head = rows[0] if keep_header and rows else []
    if keep_header:
        idx = [i for i, n in enumerate(head) if n.strip() in [str(c) for c in cols]]
    else:
        idx = [col2i(str(c)) - 1 for c in cols if re.fullmatch(r"[A-Z]{1,3}", str(c))]
    stat = {"encrypted": 0, "empty": 0}
    for r in rows[int(keep_header):]:
        for i in idx:
            if i >= len(r) or r[i] == "": stat["empty"] += 1; continue
            if HEAD.match(r[i]): stat["empty"] += 1; continue
            r[i] = r[i] if mode == "mask" else (encrypt(r[i], pw, mode) if mode != "hmac" else one_way(r[i], pw))
            stat["encrypted"] += 1
    out = io.StringIO(newline="")
    csv.writer(out, d, lineterminator=d.lineterminator).writerows(rows)
    body = out.getvalue().encode("utf-8")
    return ((codecs.BOM_UTF8 + body) if bom else body), stat

# ---------------- fixture ----------------
def make_fixture(path):
    wb = Workbook(); ws = wb.active; ws.title = "客户"
    # 第1行 = 表头（header_row=1），标题放“配置”页，避免合并标题把表头挤到第2行
    for i, h in enumerate(["客户号", "姓名", "手机号", "额度", "签约日"], 1):
        c = ws.cell(1, i, h); c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79"); c.alignment = Alignment(horizontal="center")
    DATA_LO, DATA_HI = 2, 5
    data = [("C001", "张三", "13812345678", 12000, datetime.date(2026, 1, 5)),
            ("C002", "李四", "15900001111", 8500, datetime.date(2026, 2, 11)),
            ("C003", "王五", None, 3000, datetime.date(2026, 3, 2)),
            ("C004", "赵六", "", 0, None)]
    for row in data:
        ws.append(list(row))
    for r in range(3, 7):
        ws.cell(r, 3).font = Font(bold=True, color="FF0000")
        ws.cell(r, 4).number_format = '#,##0"元"'
        ws.cell(r, 5).number_format = "yyyy-mm-dd"
        if isinstance(ws.cell(r, 5).value, datetime.date): ws.cell(r, 5).number_format = "yyyy-mm-dd"
    ws.cell(6, 4, "=SUM(D2:D5)").number_format = '#,##0"元"'
    for col, w in zip("ABCDE", (10, 10, 16, 14, 13)): ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    cfg = wb.create_sheet("配置"); cfg["A1"] = "env"; cfg["B1"] = "prod"
    wb.create_sheet("说明")["A1"] = "只读页，不该被动"
    wb.save(path)

CSV_RAW = ('\ufeffname,phone,note\r\n"张,三",13812345678,"line1\r\nline2"\r\n李四,15900001111,ok\r\n').encode()

# ---------------- 断言 ----------------
PASS, FAIL = [], []

def assert_keys(got, want, msg=""):
    """只比关键 4 项，容忍实现后续新增统计字段（如耗时）。"""
    g = {k: got.get(k) for k in want}
    assert g == want, f"{msg} 期望 {want}，实际 {g}（全量 {got}）"
def check(no, name, fn):
    try:
        fn(); PASS.append((no, name)); print(f"  ✅ {no:>2}. {name}")
    except Exception as e:
        FAIL.append((no, name, f"{type(e).__name__}: {e}")); print(f"  ❌ {no:>2}. {name} → {e}")

def _xml_leaks(path, secrets):
    """xlsx 是 zip+deflate，直接 grep 字节会假阴性；解包后查 XML 文本。"""
    import zipfile
    with zipfile.ZipFile(path) as z:
        blob = "".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist() if n.endswith(".xml"))
    return any(s in blob for s in secrets)

def main():
    pw = os.environ.get("DET_PASSPHRASE", "test-only-passphrase-请勿复用").encode()
    tmp = Path(tempfile.mkdtemp(prefix="det_"))
    src, dst = tmp / "客户表.xlsx", tmp / "客户表.encrypted.xlsx"
    make_fixture(src)
    cols = ["C"]
    stat = process_xlsx(src, dst, cols, pw, mode="aesgcm")

    print("\n[A] Excel 格式保全与范围控制")
    a, b = load_workbook(src), load_workbook(dst)
    aw, bw = a["客户"], b["客户"]
    check(1, "工作表顺序/数量一致", lambda: (_ for _ in ()).throw(AssertionError(a.sheetnames)) if a.sheetnames != b.sheetnames else None)
    check(2, "合并区域一致", lambda: (_ for _ in ()).throw(AssertionError("merged changed")) if
          {str(x) for x in aw.merged_cells.ranges} != {str(x) for x in bw.merged_cells.ranges} else None)
    def _t3():
        assert stat["encrypted"] == 2 and stat["empty"] == 3, stat
        assert stat["skipped_sheets"] == ["配置", "说明"], "按列名指定时，无此表头的页必须整页跳过而不是猜列"
        assert HEAD.match(str(bw.cell(2, 3).value)) and HEAD.match(str(bw.cell(3, 3).value))
        assert bw.cell(4, 3).value is None and bw.cell(5, 3).value is None, "None/空串都应跳过且不落密文"
    check(3, "选中格 2 个全变密文带 enc:v1 头，1 空 1 空串跳过", _t3)
    def t8():
        assert decrypt(bw.cell(2, 3).value, pw) == "13812345678"
        assert decrypt(bw.cell(3, 3).value, pw) == "15900001111"
        back = tmp / "回环.xlsx"; shutil.copy(src, back)
        wb2 = load_workbook(dst); w2 = wb2["客户"]
        for r in range(2, 7):
            if isinstance(w2.cell(r, 3).value, str) and HEAD.match(w2.cell(r, 3).value):
                w2.cell(r, 3).value = decrypt(w2.cell(r, 3).value, pw)
        wb2.save(tmp / "客户表.ROUNDTRIP.xlsx")
        c = load_workbook(tmp / "客户表.ROUNDTRIP.xlsx")["客户"]
        for r in range(1, aw.max_row + 1):
            for cc in range(1, aw.max_column + 1):
                assert aw.cell(r, cc).value == c.cell(r, cc).value, f"回环差异 ({r},{cc})"
    check(8, "解密回环后与原件逐格相等", t8)

    print("\n[B] CSV 通道")
    csrc, cdst = tmp / "订单.csv", tmp / "订单.encrypted.csv"
    csrc.write_bytes(CSV_RAW)
    out, cstat = process_csv(csrc.read_bytes(), ["phone"], pw, "aesgcm")
    cdst.write_bytes(out)
    rb = cdst.read_bytes()
    import codecs
    def t9():
        assert rb[:3] == codecs.BOM_UTF8 and rb[3:].count(codecs.BOM_UTF8) == 0, "BOM 数量不对"
        assert rb.decode("utf-8-sig").count("\r\n") == 4, "CRLF 数应=3行尾+1内嵌"
    check(9, "BOM 有且仅一份 + CRLF 原样", t9)
    bk = list(csv.reader(io.StringIO(rb.decode("utf-8-sig"), newline="")))
    def t10():
        assert bk[0] == ["name", "phone", "note"], f"表头被污染: {bk[0]}"
        assert bk[1][0] == "张,三" and bk[1][2] == "line1\r\nline2", "逗号/内嵌换行错行"
        assert bk[2][2] == "ok" and len(bk) == 3, f"应为表头+2数据行, 实际 {len(bk)}"
        assert cstat == {"encrypted": 2, "empty": 0}
    check(10, "含逗号与内嵌换行不错行 + 未选列原样 + 表头不加密", t10)
    check(11, "CSV 解密回环一致", lambda: (_ for _ in ()).throw(AssertionError("回环失败")) if
          decrypt(bk[1][1], pw) != "13812345678" or decrypt(bk[2][1], pw) != "15900001111" else None)

    print("\n[C] 粘贴数据通道")
    pasted = [["姓名", "手机号", "城市"], ["张三", "13812345678", "太原"],
              ["李四", "15900001111", "西安"], ["王五", "", "北京"]]
    pout = [pasted[0]] + [[r[0], encrypt(r[1], pw) if r[1] else "", r[2]] for r in pasted[1:]]
    mout = [pasted[0]] + [[r[0], mask(r[1]) if r[1] else "", r[2]] for r in pasted[1:]]
    def t12():
        assert pout[1][1].startswith("enc:v1:aesgcm:") and pout[3][1] == "" and pout[0][1] == "手机号"
        assert decrypt(pout[2][1], pw) == "15900001111"
        assert [len(r) for r in pout] == [3] * 4, "列数错位"
    check(12, "表头不加密/空保持空/列数对齐/回环", t12)
    check(13, "掩码模式输出符合预期", lambda: (_ for _ in ()).throw(AssertionError(mout[1][1])) if
          [r[1] for r in mout] != ["手机号", "138****5678", "159****1111", ""] else None)
    check("13b", "掩码输出不含任何明文片段", lambda: (_ for _ in ()).throw(AssertionError("掩码泄露明文")) if
          any(seg in "".join(r[1] for r in mout) for seg in ["13812345678", "15900001111"]) else None)
    check(14, "掩码/身份证/邮箱规则", lambda: (_ for _ in ()).throw(AssertionError(
          (mask("110105199003071234"), mask("a.b@corp.com"), mask("AB"), mask("")))) if
          (mask("110105199003071234"), mask("a.b@corp.com"), mask("AB"), mask("")) !=
          ("110105********1234", "a***b@corp.com", "**", "") else None)

    print("\n[D] 算法矩阵与安全负向")
    def t15():
        for alg in ("aesgcm", "chacha20", "fernet"):
            t = encrypt("密文互操作测试", pw, alg)
            assert decrypt(t, pw) == "密文互操作测试", alg
            assert HEAD.match(t), alg
    check(15, "aesgcm / chacha20 / fernet 三算法回环", t15)
    try:
        from gmssl.sm4 import CryptSM4  # noqa
        def t16():
            t = encrypt("国密互操作测试", pw, "sm4gcm"); assert decrypt(t, pw) == "国密互操作测试"
        check(16, "SM4-GCM 回环（gmssl 可用）", t16)
    except ImportError:
        PASS.append((16, "SM4-GCM 跳过（未装 gmssl，已在 README 标注）")); print("  ⏭  16. SM4-GCM 跳过（pip install gmssl 后启用）")
    sample = pout[1][1]
    def t17():
        for label, fn in [("错口令", lambda: decrypt(sample, b"wrong-pass")),
                          ("明文当密文", lambda: decrypt("13812345678", pw)),
                          ("垃圾串", lambda: decrypt("enc:v1:aesgcm:zzzz:yyyy:xxxx", pw)),
                          ("未知算法", lambda: decrypt(sample.replace(":aesgcm:", ":rsa2048:"), pw))]:
            try: fn(); raise AssertionError(f"未拦截: {label}")
            except ValueError: pass
    check(17, "错口令/明文当密文/垃圾串/未知算法 全部拒绝", t17)
    def t18():
        p = sample.split(":"); buf = bytearray(_u(p[-1])); buf[-3] ^= 1
        tampered = ":".join(p[:-1] + [_b(bytes(buf))])
        try: decrypt(tampered, pw); raise AssertionError("篡改未被 AEAD 拦住")
        except ValueError: pass
        # 换 AAD 语义：手工用错误 AAD 重算也必须失败
        salt, nonce = _u(p[3]), _u(p[4])
        try: AESGCM(_key(pw, salt)).decrypt(nonce, _u(p[-1]), b"OTHER-CTX"); raise AssertionError("AAD 未绑定")
        except Exception: pass
    check(18, "篡改 1 bit 被拒 + AAD 域分隔绑定生效", t18)
    check(19, "同明文两次密文不同（每格独立 salt）", lambda: (_ for _ in ()).throw(AssertionError("等值泄露")) if
          encrypt("同一明文", pw) == encrypt("同一明文", pw) else None)
    def t20():
        h1, h2 = one_way("13812345678", pw), one_way("13812345678", pw)
        assert h1 == h2 and h1.startswith("hmac:v1:") and len(h1) == 8 + 43
        assert one_way("13812345678", b"other") != h1 and one_way("13900000000", pw) != h1
    check(20, "HMAC 确定性可 JOIN、换口令/换值即变", t20)
    def t21():
        d2 = tmp / "幂等.xlsx"; process_xlsx(dst, d2, cols, pw, mode="aesgcm", allow_noop=True)
        v = load_workbook(d2)["客户"].cell(2, 3).value
        assert v == bw.cell(2, 3).value and str(v).count("enc:v1:") == 1, "二次运行应幂等跳过，不该双层加密"
        assert process_xlsx(src, tmp / "幂等.xlsx", cols, pw, mode="aesgcm")["encrypted"] == 2
        assert process_xlsx(src, tmp / "行范围.xlsx", cols, pw, mode="aesgcm", rows=["2"])["encrypted"] == 1, "只锁第2行 → 1格（第1行是表头）"
        r_named = process_xlsx(dst, tmp / "列名.xlsx", ["手机号"], pw, mode="aesgcm", allow_noop=True)   # 源已加密 → 合法空转
        assert r_named["encrypted"] == 0 and r_named["skipped_sheets"] == ["配置", "说明"], r_named
        assert load_workbook(tmp / "列名.xlsx")["配置"]["B1"].value == "prod", "无目标表头的 sheet 必须整页跳过"
        assert process_xlsx(src, tmp / "列名2.xlsx", ["姓名"], pw, mode="aesgcm")["encrypted"] == 4, "换一列按名命中 4 行"
        # C:E 在客户页 = 3列×4数据行 = 12格 → 3格空跳过 + 1格公式(=SUM)被护栏跳过 → 9加密/5空/1公式
        _r = process_xlsx(src, tmp / "区间.xlsx", ["C:E"], pw, mode="aesgcm")
        assert_keys(_r, {"encrypted": 9, "empty": 5, "skipped_sheets": ["配置", "说明"], "formulas_skipped": 1}, "用例21")
        r_ltr = process_xlsx(src, tmp / "区间2.xlsx", ["C:E"], pw, mode="aesgcm")
        assert_keys(r_ltr, {"encrypted": 9, "empty": 5, "skipped_sheets": ["配置", "说明"], "formulas_skipped": 1}, "用例21")
        _w = load_workbook(tmp / "区间2.xlsx")["客户"]
        assert _w.cell(6, 4).value == "=SUM(D2:D5)" and _w.cell(3, 4).number_format == '#,##0"元"', "密文格也必须保住数字格式"
        assert load_workbook(tmp / "区间2.xlsx")["配置"]["C1"].value is None, "列字母是显式跨 sheet，但该页 C:E 全空 → 原样跳过"
        _nh = process_xlsx(src, tmp / "无表头.xlsx", ["C"], pw, mode="aesgcm", keep_header=False)
        assert_keys(_nh, {"encrypted": 3, "empty": 3, "skipped_sheets": ["配置", "说明"], "formulas_skipped": 0}, "用例21")  # keep_header=False：表头1+数据2=3格加密，2空格
        for bad, kw in ((["不存在的列"], "拼错"), (["9Z"], "选列写法非法"), (["A1"], "选列写法非法"), ([""], "选列写法非法")):
            try:
                process_xlsx(src, tmp / "x.xlsx", bad, pw, mode="aesgcm"); raise AssertionError(f"非法选列竟未报错: {bad!r}")
            except ValueError as e: assert kw in str(e), f"{bad!r} 应含「{kw}」，实际 → {e}"
        assert not (tmp / "x.xlsx").exists(), "报错时不得留下半成品文件"
        # Z:AA 语法合法但三页都没这么宽 → 列字母是显式跨表指令，逐页剔除后空转 → 默认必须报警
        try:
            process_xlsx(src, tmp / "x2.xlsx", ["Z:AA"], pw, mode="aesgcm"); raise AssertionError("空转未报警")
        except ValueError as e: assert "命中 0 个可加密单元格" in str(e), e
        assert not (tmp / "x2.xlsx").exists()
        _z = process_xlsx(src, tmp / "x2.xlsx", ["Z:AA"], pw, mode="aesgcm", allow_noop=True)
        assert_keys(_z, {"encrypted": 0, "empty": 0, "skipped_sheets": ["客户", "配置", "说明"], "formulas_skipped": 0}, "Z:AA 显式空转")
        hits = [r for r in range(1, 7) if _in(r, ["3-4", "6"])]
        assert hits == [3, 4, 6], hits
    check(21, "幂等/行范围/列名/列区间/无表头 + 列名vs列字母的跨sheet语义差异", t21)
    def t22():
        assert src.stat().st_size > 0 and dst.exists()
        assert dst.name == "客户表.encrypted.xlsx" and dst.parent == src.parent
        assert src.with_name("客户表.encrypted_2.xlsx").exists() is False
        assert _xml_leaks(src, ["13812345678"]), "源文件应仍是明文（未被就地加密）"
        assert decrypt(load_workbook(dst)["客户"].cell(2, 3).value, pw) == "13812345678", "输出文件应可解回明文"
        assert not _xml_leaks(dst, ["13812345678", "15900001111"]), "xlsx 内部 XML 出现明文！"
        assert not _leaks(cdst, ["13812345678", "15900001111"]), "CSV 输出里出现了明文！"
        assert _xml_leaks(src, ["13812345678"]), "对照：源文件 XML 里本应能查到明文（证明检查有效）"
    check(22, "源文件未覆盖 + 命名带 .encrypted + 输出无明文残留", t22)

    print("\n" + "═" * 62)
    print(f"通过 {len(PASS)} / {len(PASS) + len(FAIL)}   临时产物: {tmp}")
    if FAIL:
        for f in FAIL: print("  FAIL", f)
        sys.exit(1)
    print("全绿 ✅  可交付。样例密文:", sample)
    print(f"加密统计: xlsx={stat}  csv={cstat}")
    print(f"桌面化产物: {tmp/'客户表.encrypted.xlsx'}  {cdst}")

def _leaks(path, secrets):
    blob = path.read_bytes()
    return any(s.encode() in blob for s in secrets)

# ---------------- CLI 模式 ----------------
def cli():
    ap = argparse.ArgumentParser(description="Data-Encryption-Tool 列/行加解密 CLI")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--encrypt", type=Path); g.add_argument("--decrypt", type=Path)
    ap.add_argument("-c", "--cols", required=True, nargs="+", help="列名 或 列字母(C) 或 区间(D:F)")
    ap.add_argument("--alg", default="aesgcm", choices=["aesgcm", "chacha20", "fernet", "sm4gcm", "hmac", "mask"])
    ap.add_argument("--rows", nargs="*", default=None, help="如 2-100 或 5")
    ap.add_argument("--no-header", action="store_true")
    ap.add_argument("--out", type=Path, default=None)
    a = ap.parse_args()
    pw = os.environ.get("DET_PASSPHRASE", "").encode()
    if not pw:
        import getpass; pw = getpass.getpass("加密口令（不回显，不写入任何文件）: ").encode()
    src = a.encrypt or a.decrypt
    if a.decrypt:
        dst = a.out or src.with_name(src.stem.replace(".encrypted", "") + ".DECRYPTED" + src.suffix)
        wb = load_workbook(src); n = 0
        for ws in wb.worksheets:
            try:
                targets = resolve_cols(ws, a.cols)
            except ValueError:
                if any(_is_named(c) for c in a.cols): print(f"  · sheet「{ws.title}」无该表头，跳过"); continue
                raise
            for c in targets:
                for r in range(2 if not a.no_header else 1, ws.max_row + 1):
                    if not _in(r, a.rows): continue
                    cell = ws.cell(r, c)
                    if isinstance(cell.value, str) and HEAD.match(cell.value):
                        cell.value = decrypt(cell.value, pw); n += 1
        wb.save(dst); print(json.dumps({"out": str(dst), "decrypted": n}, ensure_ascii=False)); return
    dst = a.out or src.with_name(src.stem + ".encrypted" + src.suffix)
    if src.suffix.lower() in (".csv", ".tsv", ".txt"):
        mode = "mask" if a.alg == "mask" else a.alg
        out, st = process_csv(src.read_bytes(), a.cols, pw, a.alg, not a.no_header)
        dst.write_bytes(out)
    else:
        st = process_xlsx(src, dst, a.cols, pw, mode=a.alg, rows=a.rows, header_row=1, keep_header=not a.no_header)
    print(json.dumps({"out": str(dst), **st, "alg": a.alg,
                      "note": "口令未存储；请用 --decrypt 试解一格验证"}, ensure_ascii=False))

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--encrypt", "--decrypt"):
        cli()
    else:
        main()
