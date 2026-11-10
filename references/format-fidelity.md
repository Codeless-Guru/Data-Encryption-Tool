# 格式保全清单

## openpyxl 通道：改 `cell.value` 后可原样保留 ✅

字体（含粗体/颜色/字号）、填充、边框、对齐、数字格式 `number_format`、列宽 `column_dimensions[x].width`、行高、合并单元格 `merged_cells.ranges`、工作表顺序与名称、冻结窗格、隐藏行列/sheet、数据验证、条件格式、命名样式、单元格注释、`date`/`datetime` 的原生类型、公式字符串（`data_only=False`）、`.xlsm` 的 VBA 模块（`keep_vba=True`）。

## 已知会丢或变形 ⚠️ —— 必须在交付时主动告知

| 对象 | 后果 | 应对 |
|---|---|---|
| 图表 / 透视表 / 切片器 / 迷你图 | openpyxl 不保真，可能消失 | 预检 `ws._charts` / `ws._pivots`，非空先提醒留底，或改走 CSV |
| VBA 窗体控件 / ActiveX | 丢失 | 同上，建议人工另存副本 |
| 公式引用了被加密列 | 结果变 `#VALUE!`（密文是字符串） | dry-run 时用 `openpyxl.formula` 或正则扫 `=` 引用，命中就点名提示 |
| 数字被加密 | 类型从 `int/float` 变 `str` | 属预期（密文必须是文本）；解密后若要还原数字类型，按原 `number_format` 做一次 `int/float` 回转 |
| 共享工作簿 / 外部链接 | 可能告警 | 提示用户在副本上打开验证一次 |
| `.xls` | 不支持 | 让用户先另存 `.xlsx` |
| 超大文件（>50 万格） | 内存与耗时 | 改用 CSV 通道或 `read_only/write_only` 流式；先分批 |

## CSV 通道三个坑

1. **BOM**：读用 `decode("utf-8-sig")`（否则 BOM 粘进第一个列名，列名匹配全失败），写时仅在原文有 BOM 时补回 `codecs.BOM_UTF8`，且**只能有一份**。
2. **行尾**：必须沿用原文 `dialect.lineterminator`。写成 `\n` 会让 Windows 记事本挤成一行；统计行尾数时记得减去单元格内嵌换行。
3. **引号**：交给 `csv` 模块。手写 `line.split(",")` 在 `"张,三"` 上必然错行，且内嵌换行会直接断行。

## 粘贴数据通道

先 `csv.reader(io.StringIO(text))` 用 `\t` 或 ` | ` dialect 解析，跑同一套列索引逻辑，再按**原分隔符**拼回；连续空格对齐的表格不要猜列宽，直接问用户第几列。
