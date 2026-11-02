# Excel 批量核对助手（UiBot Store 组件）

> 零依赖（纯标准库，不装 openpyxl/pandas）、支持 `.xlsx` 与 `.csv` 的**两表批量核对**组件。
> 按主键列比对两张表，自动找出：① 仅 A 有（B 缺失）② 仅 B 有（A 缺失）③ 主键匹配但值不一致。
> 输出带「核对状态」列的差异报告 CSV。企业财务对账、订单对账、库存盘点等场景通用。

**作者**：支一郎 (Zhiyilang) · **许可证**：MIT · **资助**：https://afdian.com/a/zhiyilang

---

## 为什么值得上架

- **通用痛点**：每个月都有人问"两表怎么比对"，Excel 公式 / VLOOKUP 易错、数据量大就卡。
- **零依赖**：很多 RPA 运行环境没装 openpyxl/pandas；本组件纯标准库，丢进去就能跑，不会"本地能跑线上崩"。
- **可卖点清晰**：财务对账、订单核对、主数据治理——都是付费意愿强的场景。

## 在 UiBot Creator 里使用（Python 模式）

1. 新建「流程」项目，语言选 **Python**，把本目录作为项目根。
2. 用 UiBot 原生 **Excel 读取** 命令把两张表读成数据表（二维数组）。
3. 转成行字典列表后调用核对内核：

```python
from recon_engine import reconcile, write_report, summarize

# rows_a / rows_b 是你从 UiBot Excel 读出的 [{'列名':'值'}, ...]
result = reconcile(rows_a, rows_b, key="订单号")   # key=主键列名
write_report(result, "核对报告.csv")
print(summarize(result))
```

4. 运行 → 在同目录生成 `核对报告.csv`（含「核对状态 / 主键 / A行号 / B行号 / 各比对列」）。
5. 想直接读文件（非 UiBot 场景）：`run("A.xlsx","B.xlsx", key="订单号", xlsx=True)`。

> `project.cfg` 是占位配置，用 Creator 打开后会自动校正为本机格式，无需手动改。

## 独立测试（不依赖 UiBot）

```bash
python main.py          # 生成示例 xlsx 并演示完整核对
python -m unittest discover -s tests   # 跑单元测试（标准库，无需 pytest）
```

## 上架 UiBot Store

1. UiBot Creator 打开本项目 → 菜单「发布到 UiBot Store」（或导出为命令库/流程模板）。
2. 上架文案直接复制 `UIBOT_STORE_LISTING.md`（标题/描述/价格/关键词都写好了）。
3. 截图：跑一次 `python main.py`，截「终端摘要」+「生成的核对报告.csv」即可。

## 变现建议

- 基础版免费/低价（¥9.9）引流，企业定制（多 Sheet、跨工作簿、数据库直连）走爱发电 ¥300/月起。
- 同作者在 Gitee 的 RPA→AI 工具箱：https://gitee.com/ZhiYIlang/rpa2mcp
