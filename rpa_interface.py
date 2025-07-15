from typing import Dict, List, Any
from main import ExcelAnalyzer

class RPAInterface:
    """
    RPA接口类
    用于RPA调用Python分析功能
    """
    def __init__(self):
        self.analyzer = ExcelAnalyzer()
        
    def analyze_data(self, rule_file: str, data_file: str, output_file: str = None) -> Dict[str, Any]:
        """
        分析数据的主要接口
        
        Args:
            rule_file: 规则表文件路径（从Google Sheet下载的Excel）
            data_file: 数据文件路径（从Google Sheet下载的Excel）
            output_file: 输出文件路径（可选，如果需要生成Excel报告）
            
        Returns:
            Dict[str, Any]: {
                'success': bool,  # 是否成功
                'message': str,   # 错误信息（如果失败）
                'results': List[Dict],  # 分析结果
            }
        """
        try:
            # 加载规则表
            if not self.analyzer.load_rule_table(rule_file):
                return {
                    'success': False,
                    'message': '加载规则表失败',
                    'results': []
                }
                
            # 加载数据文件
            if not self.analyzer.load_excel_data(data_file):
                return {
                    'success': False,
                    'message': '加载数据文件失败',
                    'results': []
                }
                
            # 处理规则
            if not self.analyzer.process_rules():
                return {
                    'success': False,
                    'message': '处理规则失败',
                    'results': []
                }
                
            # 如果需要生成Excel报告
            if output_file:
                if not self.analyzer.generate_report(output_file):
                    return {
                        'success': False,
                        'message': '生成报告失败',
                        'results': self.analyzer.get_results()
                    }
                    
            return {
                'success': True,
                'message': '分析完成',
                'results': self.analyzer.get_results()
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'发生错误: {str(e)}',
                'results': []
            }
            
    def get_formatted_results(self) -> List[Dict[str, str]]:
        """
        获取格式化的结果，便于RPA处理
        
        Returns:
            List[Dict[str, str]]: [
                {
                    'description': str,  # 描述
                    'result': str,       # 结果（已格式化为字符串）
                    'comments': str,     # 注释
                    'optimization_plan': str  # 优化计划
                },
                ...
            ]
        """
        results = self.analyzer.get_results()
        formatted_results = []
        
        for result in results:
            formatted_results.append({
                'description': str(result.get('description', '')),
                'result': self.analyzer._format_result(result.get('result')),
                'comments': str(result.get('comments', '')),
                'optimization_plan': str(result.get('optimization_plan', ''))
            })
            
        return formatted_results

# RPA调用示例：
"""
from rpa_interface import RPAInterface

# 创建接口实例
rpa_interface = RPAInterface()

# 分析数据
result = rpa_interface.analyze_data(
    rule_file='path/to/rule_table.xlsx',  # RPA从Google Sheet下载的规则表
    data_file='path/to/data.xlsx',        # RPA从Google Sheet下载的数据文件
    output_file='path/to/report.xlsx'     # 可选，如果需要生成Excel报告
)

if result['success']:
    # 获取格式化的结果
    formatted_results = rpa_interface.get_formatted_results()
    
    # RPA可以直接使用formatted_results更新Google Sheet
    for item in formatted_results:
        description = item['description']
        result_value = item['result']
        comments = item['comments']
        optimization_plan = item['optimization_plan']
        # RPA代码：更新Google Sheet对应的单元格
else:
    print(f"分析失败: {result['message']}")
""" 