# Excel数据分析RPA控件

这是一个用于RPA系统的Excel数据分析控件，可以根据预定义的规则自动分析Excel数据并生成报告。

## 主要功能

1. 规则查找
   - 支持正则表达式搜索
   - 可在描述和规则类型中查找
   - 返回匹配的规则列表

2. 数据分析
   - KPI分析：支持自定义阈值，返回状态（良好/警告/不达标）
   - 最值分析：支持查找最大/最小N个值
   - 单元格分析：支持格式化输出

3. 报告生成
   - 生成Excel格式的分析报告
   - 包含规则详情和分析结果
   - 支持自定义结果格式化

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

1. 准备规则表（Excel文件）：
   - Sheet名称：规则表
   - 必需列：Sheet, Location, Description, Rule
   - 可选列：Comments（用于KPI阈值和最值数量配置）

2. 代码示例：
```python
from excel_analyzer_control import ExcelAnalyzerControl

# 创建控件实例
analyzer = ExcelAnalyzerControl()

# 加载Excel文件
analyzer.load_excel("your_excel.xlsx")

# 查找特定规则
kpi_rules = analyzer.find_rules_by_pattern("KPI")

# 分析规则
for rule in kpi_rules:
    result = analyzer.analyze_by_rule(rule['id'])
    print(result)

# 生成报告
analyzer.generate_report("analysis_report.xlsx")
```

## 规则类型说明

1. KPI规则
   - 在Comments中定义阈值：`good: >90, warning: >70, bad: <=70`
   - 默认阈值：良好>=90, 警告>=70, 不达标<70

2. 最值规则
   - 在Comments中定义获取数量：`top 5`
   - 默认获取前3个最大值和最小值

3. 单元格规则
   - 直接获取指定位置的值
   - 自动格式化数值（保留2位小数）

## 注意事项

1. Excel文件格式要求：
   - 支持.xlsx格式
   - 第一个sheet必须是规则表
   - 规则表必须包含所有必需列

2. 错误处理：
   - 所有操作都有完善的错误处理
   - 通过日志记录详细信息
   - 返回值包含成功/失败状态 