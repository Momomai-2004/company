# -*- coding: utf-8 -*-
import sys
import json
from excel_analyzer_control import ExcelAnalyzerControl

def run_analysis_for_aa():
    """
    专为 Automation Anywhere 设计的执行入口。
    - 从命令行参数接收输入文件路径。
    - 将分析结果以 JSON 格式打印到标准输出，供 AA 捕获。
    """
    # AA 会把参数传给 sys.argv
    if len(sys.argv) < 2:
        # 如果没有接收到文件路径，打印错误并退出
        print(json.dumps({"success": False, "error": "No input file path provided."}))
        return

    input_excel_path = sys.argv[1]

    try:
        analyzer = ExcelAnalyzerControl()

        # 1. 加载 RPA 从 Google Sheet 下载的 Excel 文件
        if not analyzer.load_excel(input_excel_path):
            raise RuntimeError("Failed to load Excel file.")

        # 2. 执行所有分析规则
        results = analyzer.analyze_all()

        # 3. 将最终结果打包成 JSON 字符串，打印出来
        # AA 会捕获这个 print 输出
        print(json.dumps(results, ensure_ascii=False))

    except Exception as e:
        # 如果过程中出错，也打印 JSON 格式的错误信息
        print(json.dumps({"success": False, "error": str(e)}))

if __name__ == "__main__":
    run_analysis_for_aa() 