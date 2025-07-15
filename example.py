from excel_analyzer_control import ExcelAnalyzerControl

def main():
    # 创建控件实例
    analyzer = ExcelAnalyzerControl()
    
    # 加载Excel文件
    success = analyzer.load_excel("example.xlsx")
    if not success:
        print("加载Excel文件失败")
        return
        
    # 查找包含"KPI"的规则
    print("\n查找KPI相关规则：")
    kpi_rules = analyzer.find_rules_by_pattern("KPI")
    for rule in kpi_rules:
        print(f"规则{rule['id']}: {rule['description']}")
        
    # 分析找到的KPI规则
    print("\n分析KPI规则：")
    for rule in kpi_rules:
        result = analyzer.analyze_by_rule(rule['id'])
        print(f"规则{rule['id']}分析结果：")
        print(result)
        
    # 查找包含"最值"的规则
    print("\n查找最值相关规则：")
    extremum_rules = analyzer.find_rules_by_pattern("最值")
    for rule in extremum_rules:
        print(f"规则{rule['id']}: {rule['description']}")
        
    # 分析所有规则并生成报告
    print("\n生成分析报告...")
    analyzer.generate_report("analysis_report.xlsx")
    print("分析完成，请查看analysis_report.xlsx")

if __name__ == "__main__":
    main() 