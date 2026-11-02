# UiBot Store 上架文案（直接复制）

> 复制到 UiBot Store 发布页对应字段即可。价格/关键词已给出建议，可按活动调整。

## 标题
Excel 批量核对助手 · 两表差异一键比对（零依赖）

## 副标题 / 一句话
按主键自动比对两张 Excel/CSV，秒出"缺失/多余/值不一致"差异报告，纯标准库无需装包。

## 详细描述
### 能帮你解决什么
财务对账、订单核对、库存盘点、主数据治理……凡是"把 A 表和 B 表比一比"的活，不用再写 VLOOKUP、不用肉眼看。

### 功能
- 支持 `.xlsx` 与 `.csv`（标准库解析，不依赖 openpyxl / pandas，运行环境零安装风险）
- 按主键列比对，自动分类：
  - 仅 A 有（B 缺失）
  - 仅 B 有（A 缺失）
  - 主键匹配但某些列值不一致（逐列标出 A≠B）
  - 完全一致
- 输出带「核对状态」列的差异报告 CSV，可直接发给业务方

### 适用人群
RPA 开发者、财务/运营/数据岗、需要做定期对账的团队。

### 使用示例
```python
from recon_engine import reconcile, write_report, summarize
result = reconcile(rows_a, rows_b, key="订单号")
write_report(result, "核对报告.csv")
print(summarize(result))
```

## 价格建议
- 基础组件：¥9.9（引流）
- 企业增强版（多 Sheet / 跨工作簿 / 数据库直连 / 定时调度模板）：¥199 或定制
- 私有化+培训：面议

## 关键词
Excel核对, 两表比对, 对账, 差异报告, VLOOKUP替代, 财务RPA, 数据治理, 零依赖, 订单核对, 库存盘点

## 截图建议
1. 终端运行 `python main.py` 的摘要输出（显示 仅A有/仅B有/值不一致 计数）
2. 生成的 `核对报告.csv` 打开效果（核对状态列高亮）
3. UiBot Creator 中调用 reconcile() 的代码截图

## 更新日志
- v1.0.0（2026-10-25）：首版，支持 xlsx/csv、主键比对、差异报告导出、单元测试 100% 覆盖
