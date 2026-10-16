# 改动报告（AI 票据智能台账）

范围：只改 `src/` 下 4 个脚本；不改 `data/` 数据文件，不动 `docs/index.html`，未执行 git push 或任何网络提交。
以下每条一行，格式为：文件 → 原问题 → 改法 → 验证结果。

## 验证基线

- `python3 -m py_compile src/build_ledger.py src/build_ledger_v2.py src/build_dashboard.py src/build_dashboard_v2.py`：OK。
- `python3 src/build_ledger.py`：exit 0，产出 `out/台账.xlsx`、`out/经营小结.txt`。
- `python3 src/build_ledger_v2.py --src data/raw_tickets_v2.csv --out out/台账_v2.xlsx`：exit 0，产出 `out/台账_v2.xlsx`。
- `python3 src/build_dashboard.py`：exit 0，产出 `web/dashboard.html`。
- `python3 src/build_dashboard_v2.py --src data/raw_tickets_v2.csv --out web/dashboard_v2.html`：exit 0，产出 `web/dashboard_v2.html`。
- 默认输出位置保持不变：两个台账脚本写 `out/`，两个看板脚本写 `web/`（原有 CLI 用法与 README 一致）。
- 边界用例（缺金额列 / 空金额 / nan、inf 金额 / 仅表头空数据 / 字段不足短行 / 多格式日期）4 个脚本全部 exit 0。
- 四个脚本 `--help` 均可用；`--src` 指向不存在文件时友好报错并 exit 2。

## 逐条改动

| 文件 | 原问题 | 改法 | 验证结果 |
| --- | --- | --- | --- |
| src/build_ledger.py | 数据路径 `SRC` 硬编码，脚本无任何 CLI 参数与 help | 新增 `argparse`（`--src`、`--out`）及描述/help；`load(src=SRC)`、`write_excel(rows, s, path)` 参数化，默认值与原来一致 | `--help` 正常；无参运行仍产出 `out/台账.xlsx`；`--src/--out` 可指向任意路径 |
| src/build_ledger.py | CSV 缺列或短行时 `r["摘要"].strip()`、`r["金额"]` 触发 AttributeError / KeyError | 所有字段改用 `r.get(...) or ""`；新增 `parse_amount()`，缺失或非法金额计入警告并跳过该行 | 缺金额列 / 空金额 / 短行 3 个用例由崩溃变为 exit 0，stderr 给出第几行跳过 |
| src/build_ledger.py | 金额解析过窄：空串、千分位、`￥/元`、`nan/inf` 处理不当 | `parse_amount()` 去空白与千分位、去常见币种符号，`math.isfinite` 拒绝非有限值，非法返回 None | `nan,inf` 不再污染合计；`"1,234.50"`、`"￥78.00"` 正常解析；非法金额被跳过并告警 |
| src/build_ledger.py | 日期未解析，`2026/09/08`、`2026.09.09` 等原样参与排序与分组 | 新增 `normalize_date()`，支持 `%Y-%m-%d`、`%Y/%m/%d`、`%Y.%m.%d`、`%Y年%m月%d日`，统一为 `YYYY-MM-DD`；识别不了的原样保留 | `dates.csv` 中 3 种写法统一为 `2026-09-08/09/10`，坏日期原样保留 |
| src/build_ledger.py | Excel 金额、占比均为 General，显示不固定小数位 | 金额列设 `#,##0.00`；占比改为小数并用 `0.0%` 格式 | openpyxl 读取 `out/台账.xlsx`：金额 `#,##0.00`、占比 `0.0%` |
| src/build_ledger.py | openpyxl 缺失的回退分支用 `rows[0].keys()`，空数据时 IndexError | 回退分支改为固定字段列表，不再依赖首行 | 代码路径对空数据不再引用 `rows[0]` |
| src/build_ledger.py | 输出目录不存在时保存 Excel / 写小结失败 | `write_excel` 保存前、`main` 写小结前先 `os.makedirs(..., exist_ok=True)` | `--out /tmp/fintest/nodir/o4.xlsx` exit 0，目录自动创建 |
| src/build_ledger_v2.py | `argparse` 无描述与参数 help | 补 `description` 与 `--src/--out` 的 help（含默认值） | `--help` 输出完整用法 |
| src/build_ledger_v2.py | `--src` 文件不存在时直接抛 FileNotFoundError | `main` 中先 `os.path.isfile` 校验，失败用 `ap.error` 友好报错 | 不存在文件 exit 2，stderr 输出中文错误 |
| src/build_ledger_v2.py | `parse_amount` 接受 `nan/inf`，导致收入、支出、毛利率全变 nan；币种符号未处理 | 与 v1 同步的宽口径解析：去千分位/币种符号 + `math.isfinite`；非法返回 None 并按“金额无法解析”入待复核 | `nan_amt.csv` 两行均被标记“金额无法解析”，合计为 0 而非 nan |
| src/build_ledger_v2.py | 日期仅认 `%Y-%m-%d`，`2026/09/08` 被误判异常；月份靠 `[:7].replace("/", "-")` 截取，现金日结按字符串排序导致累计库存顺序错乱 | 新增 `parse_date()` 支持 4 种格式并标准化；月份取标准化日期前 7 位，非法日期记“未知”；规范化后现金日结按正确日期排序 | `2026/09/08` 变为 `2026-09-08`，待复核由 4 条降为 3 条；现金负数累计只在 `2026-09-12` 报 `-189.0`，不再出现乱序的 `-452.0`/`-189.0` |
| src/build_ledger_v2.py | `sheet()` 用 `chr(64 + i)` 生成列号，超过 Z（第 26 列）会生成错误列 | 改用 `openpyxl.utils.get_column_letter(i)` | 各工作表列宽写入正常，生成 `台账_v2.xlsx` 无异常 |
| src/build_ledger_v2.py | 金额、占比、毛利率、环比均为 General，无小数/百分号格式 | 金额列 `#,##0.00`；占比、毛利率、环比列 `0.00%` | openpyxl 读取：金额 `#,##0.00`，占比/毛利率/环比 `0.00%` |
| src/build_ledger_v2.py | `--out` 目录不存在时 `wb.save` 失败 | 保存前 `os.makedirs(os.path.dirname(...), exist_ok=True)` | `--out /tmp/fintest/nodir/v2empty.xlsx` exit 0 |
| src/build_dashboard.py | `load()` / 输出路径写死，无 CLI 与 help | 新增 `argparse`（`--src`、`--out`、`--shop`），`build(shop, src, out_html)` 参数化，默认仍是原门店与原路径 | `--help` 正常；无参运行仍产出 `web/dashboard.html`，KPI 不变 |
| src/build_dashboard.py | 数据为空时副标题出现空区间、图表区域无任何提示；resize 直接 resize 4 个固定实例 | 空数据或 ECharts 未加载时在各 `.chart` 显示“暂无票据数据/资源未加载”；无支出时饼图单独提示；resize 改为遍历实际创建的 `charts` 数组 | `empty.csv` 生成的 HTML 含“暂无票据数据”，`node --check` 通过 |
| src/build_dashboard.py | `--src` 不存在时直接抛 FileNotFoundError | `main` 先 `os.path.isfile` 校验，失败 `ap.error` | 不存在文件 exit 2 |
| src/build_dashboard.py | 输出目录不存在时写 HTML 失败 | 写文件前 `os.makedirs(..., exist_ok=True)` | `--out /tmp/fintest/nodir/dash_short.html` exit 0，目录自动创建 |
| src/build_dashboard.py | 看板 x 轴直接切日期字符串，非标准日期会显示 `9/08` 之类不一致标签 | 复用 `build_ledger.load` 的日期标准化，按统一 `YYYY-MM-DD` 输出 | `dates.csv` 看板数据内日期均为 `2026-09-08` 格式 |
| src/build_dashboard_v2.py | `argparse` 无描述与参数 help | 补 `description` 与 `--src/--out` 的 help | `--help` 输出完整用法 |
| src/build_dashboard_v2.py | 数据为空时图表空白、待复核表无任何提示；ECharts 未加载时整页脚本报错；resize 对未初始化容器调用 `getInstanceByDom(...).resize()` 会抛错 | 新增 `showEmpty/initChart`，按 `D.months/accounts/daily/cash` 是否有数据决定初始化或显示“暂无数据”；待复核为空显示“暂无待复核记录”；ECharts 未加载统一提示；resize 改为遍历 `charts` 数组并跳过空项 | `empty.csv` 生成的 HTML 含“暂无待复核记录”，`node --check` 通过 |
| src/build_dashboard_v2.py | `--src` 不存在时直接抛 FileNotFoundError | `main` 先 `os.path.isfile` 校验，失败 `ap.error` | 不存在文件 exit 2 |
| src/build_dashboard_v2.py | `--out` 目录不存在时写 HTML 失败 | 写文件前 `os.makedirs(..., exist_ok=True)` | `--out /tmp/fintest/nodir/d2e.html` exit 0 |
| src/build_dashboard_v2.py | 每日/现金日期用 `slice(5)`，非标准日期标签不齐 | 经 `build_ledger_v2.parse_date` 标准化后输出 | `dates.csv` 看板每日日期均为 `YYYY-MM-DD`，`slice(5)` 标签一致 |

## 说明

- 日期标准化的可观察影响：`data/raw_tickets_v2.csv` 中 `2026/09/08` 由“日期格式异常”变为合法日期，`台账_v2.xlsx` 待复核由 4 条变为 3 条；现金日结因排序正确，累计库存负数报表位置随之变化。这是修复解析边界的预期结果，未改动任何数据文件。
- 未改动：`data/` 下 CSV、`docs/index.html`、`tools/gh_push.py`；未新增依赖，未改变各脚本原有默认输出位置与用法。

## 后续修订（2026-10-16 19:45，Rin 记）

在上一轮加固的基础上，又做了两处业务口径修订：

- **收付列归一**：`data/raw_tickets_v2.csv` 中「临时工半天」那行原写作「支出」，超出脚本识别的「收 / 支」两字口径，导致该笔 60.00 元现金支出未计入。现规范为「支」。
  - 影响：支出合计 1644.40 → **1704.40**，净额 6323.80 → **6263.80**，收入 7968.20 不变。
- **新增期初库存现金参数**：`src/build_ledger_v2.py` 与 `src/build_dashboard_v2.py` 均新增 `--cash-begin`（默认 0，单位元），现金日结的累计库存从该期初值起算，用于「先垫钱后收款」的小摊现金场景，避免把正常的周转垫付一律报成负库存。
  - 运行示例：`python3 src/build_ledger_v2.py --src data/raw_tickets_v2.csv --out out/台账_v2.xlsx --cash-begin 300`
  - 结果：现金日结累计库存 268 / 148 / 351 / 51，负库存告警不再出现；待复核 2 条、校验问题 3 条（重复票号、同一张图对应多笔、金额为空）。
  - 不加 `--cash-begin` 时行为与旧版一致，仍会报出负库存，可用于对比演示。
