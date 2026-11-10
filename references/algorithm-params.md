# 算法参数与互通性

| 模式 | KDF | 密钥 | Nonce/IV | 认证 | 密文增量（每格） | 互通对象 |
|---|---|---|---|---|---|---|
| `aesgcm`（默认） | Scrypt 2¹⁴/8/1 + salt16B | 32B | 12B 随机 | GCM tag 16B | salt16+nonce12+tag16+头≈44B + b64(4/3) ≈ **120B** | 任何 AES-256-GCM 实现 |
| `aesgcm-cbc` ⏳ | 同上 | 32B | IV16B 随机 + PKCS7 | 外层 HMAC-SHA256(前16B) | ≈140B | OpenSSL `aes-256-cbc`（**v0.1.0 未实现，本表为设计参数，规划 v0.2**） |
| `chacha20` | 同上 | 32B | 12B 随机 | Poly1305 16B | ≈120B | libsodium / `age` 内部原语 |
| `fernet` | 同上 → urlsafe_b64 | 32B | 内嵌 16B | HMAC-SHA256 前10B | ≈130B | `cryptography.fernet`、Django 字段 |
| `sm4gcm` | 同上（截 16B） | 16B | 12B 随机 | GCM tag 16B | ≈120B | `gmssl`、国密套件 |
| `age` ⏳ | X25519 + HKDF（Ruby/Go age 规范） | 文件密钥 | 自动 | ChaCha20-Poly1305 | 按 recipient | `age` CLI、`rage`（**v0.1.0 未实现，本表为设计参数，规划 v0.2**） |
| `hmac` | Scrypt + pepper `det-hmac-pepper` | 32B | 无（确定性） | — | 固定 43B hex | 任何 HMAC-SHA256 |
| `mask` | 无 | — | — | — | 长度不变 | 人眼 |

## 固定项（改了就不互通）

- 头部魔数：`enc:v1:`，5 段冒号分隔；`fernet` 的第 4 段为占位 nonce，解析时忽略
- AEAD AAD：字面量 `b"DETv1"`。**解密时 AAD 不一致会直接失败**，这是防「密文挪到别的列/别的用途」的廉价手段；需要按列绑定时把 AAD 改成 `f"DETv1:{sheet}!{col}{row}"`
- Salt 16B / Nonce 12B 均来自 `os.urandom`，**每格独立**

## 性能

Scrypt n=2¹⁴ 约 30–60ms/格（本机实测量级）。1 万格 ≈ 5–10 分钟，属预期。
- 批量场景把 KDF 提一层：每文件一次 Scrypt 主密钥 + HKDF 派生每格子密钥（salt 用 `文件salt || 行号 || 列号`），可从分钟级降到秒级；代价是**密文格式要升 `enc:v2`**，且失去「每格独立 salt」的强隔离，写进变更记录再改。
- 列内大量重复值（性别/省份）且需要抗等值泄露 → 保持 v1 每格 salt，别优化。

## 迁移与轮换

密文里没有版本号以外的密钥信息 → **换口令 = 全量解密再重加密**，必须留一次双跑窗口。要免轮换停机，改用 `age`/KMS 包一层密钥封装（DEK 不变，只重封 KEK）。
