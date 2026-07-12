# 翻译字典构建方法（领域相关，不随 skill 捆绑）

本 skill 的脚本只负责"把译文嵌入源单元格"的力学，不捆绑任何具体字典。
翻译字典必须按项目/文件现建。下面是稳定、可复用的提取→分类→翻译流程。

注意：表头字符串（如 `Call`、`Cases`）也放进**同一份**字典。脚本用同一个
`--dict` 处理表头与正文，不再有独立的硬编码表头映射（避免跨项目串味）。

## 1. 提取唯一文本串

```python
import openpyxl, re, json
wb = openpyxl.load_workbook(path, data_only=True)
cjk = re.compile(r'[\u4e00-\u9fff]'); lat = re.compile(r'[A-Za-z]')
en, zh, mixed = set(), set(), set()
for sn in wb.sheetnames:
    ws = wb[sn]
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            s = v.strip()
            if not s:
                continue
            if cjk.search(s):
                (mixed if lat.search(s) else zh).add(s)
            elif lat.search(s):
                en.add(s)
```

- `en`：纯英文唯一串（量最大，需译中文）
- `zh`：纯中文唯一串（需译英文）
- `mixed`：中英混合唯一串（中文为主，译英文，保留其中的代码/编号）

## 2. 分类为"可译 / 不可译"

不可译（跳过、不乱写）：
- 测试用例编号：`NT_DSDS_LTE_016`、`GPS_54`、`CLSET_37`、`EL-144`
- 缩写 / 技术代号：`VoLTE`、`5G`、`DSDS`、`WFC`、`USSD`、`OOS`
- 运营商 / 品牌 / 人名：`Airtel`、`JIO`、`VI`、`6759LV`
- Excel 错误值：`#DIV/0!`、`#VALUE!`

可译（进字典）：真正的英文短语、中文描述句、混合描述句。

## 3. 构建字典

最终导出**一份** JSON：`{"source": "translation", ...}`，把英文→中文、
中文→英文、以及表头译文全部合并进去（脚本只认这一个 `--dict`）。

### 英文 → 中文（en2zh）
- 高频短语用一张 `PHRASES` 短语表做"整串优先 + 长词优先替换"的自动翻译：
  ```python
  PHRASES = {"Call log":"通话记录",
             "Incoming call":"来电","Speakerphone":"免提", ...}
  def auto(s):
      if s in PHRASES: return PHRASES[s]
      out = s
      for p in sorted(PHRASES, key=len, reverse=True):
          out = out.replace(p, PHRASES[p])
      return out if out != s else None
  ```
  **注意：判断性 / 状态结论字符（PASS、FAIL、N/A、OK、NG …）不要进字典，
  脚本 `STATUS_SKIP` 会强制跳过它们、绝不翻译——用户明确要求这类字符不修改。**
- 低频 / 领域特有的长句（如测试用例步骤）**必须人工译**，质量优先于覆盖率。
- 不可译项不要进字典（或映射为自身），脚本遇到无译文即跳过。

### 中文 → 英文（zh2en）
- 数量通常很少（几十到几百），**全部人工译**，保证自然准确。
- 混合串：翻译其中的中文部分，保留英文代码/编号（如 `EL-228`、`【phone】`、`6759LV`）不动。
- 例：`"默认免提，不支持听筒"` → `"Default speakerphone; earpiece not supported"`

## 4. 落库与复用

- 把合并后的 `{source: translation}` 字典存为 JSON，脚本用 `--dict` 加载。
- 同一项目的多份文件可复用同一字典；换项目重新提取+人工补译。
- 宁可漏译（跳过）也不要错译——用户明确要求"不要乱写"。
- 幂等/增量：脚本重跑时会用 `is_already_translated()` 识别已嵌好译文的
  单元格（原文\\n译文、首尾语言相反）并跳过，因此**已翻译的文件可安全重复运行，
  只补未译部分**，不会重复翻译或把译文再翻一遍。
