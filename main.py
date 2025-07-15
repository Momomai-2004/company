import pandas as pd
from pathlib import Path
from typing import Dict, List, Any
from rule_extractor import extract_rule_data

class ExcelAnalyzer:
    """
    Excel分析器类
    设计成类的形式便于RPA调用和集成
    """
    def __init__(self):
        self.rules = None
        self.data_dict = None
        self.results = None
        
    def load_excel_data(self, file_path: str) -> bool:
        """
        读取Excel文件中的所有工作表数据
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            bool: 是否成功
        """
        try:
            xlsx = pd.ExcelFile(file_path)
            self.data_dict = {}
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name)
                self.data_dict[sheet_name] = df.values.tolist()
                self.data_dict[sheet_name].insert(0, df.columns.tolist())
            return True
        except Exception as e:
            print(f"读取Excel文件失败: {e}")
            return False

    def load_rule_table(self, file_path: str) -> bool:
        """
        读取规则表
        
        Args:
            file_path: 规则表Excel文件路径
            
        Returns:
            bool: 是否成功
        """
        try:
            df = pd.read_excel(file_path, sheet_name=0)
            self.rules = df.to_dict('records')
            return True
        except Exception as e:
            print(f"读取规则表失败: {e}")
            return False

    def process_rules(self) -> bool:
        """
        处理所有规则
        
        Returns:
            bool: 是否成功
        """
        if not self.rules or not self.data_dict:
            print("请先加载规则表和数据文件")
            return False
            
        self.results = []
        for rule in self.rules:
            try:
                result = extract_rule_data(rule, self.data_dict)
                self.results.append(result)
            except Exception as e:
                print(f"处理规则失败: {rule.get('Description', '未知规则')}, 错误: {e}")
                self.results.append({
                    'description': rule.get('Description', '未知规则'),
                    'result': None,
                    'comments': f"处理失败: {str(e)}",
                    'optimization_plan': ''
                })
        return True

    def generate_report(self, output_file: str) -> bool:
        """
        生成报告
        
        Args:
            output_file: 输出文件路径
            
        Returns:
            bool: 是否成功
        """
        if not self.results:
            print("请先处理规则")
            return False
            
        try:
            df = pd.DataFrame(columns=['描述', '结果', '注释', '优化计划'])
            
            for result in self.results:
                df = pd.concat([df, pd.DataFrame([{
                    '描述': result['description'],
                    '结果': self._format_result(result['result']),
                    '注释': result['comments'],
                    '优化计划': result['optimization_plan']
                }])], ignore_index=True)
            
            df.to_excel(output_file, index=False, sheet_name='分析报告')
            print(f"报告已生成: {output_file}")
            return True
        except Exception as e:
            print(f"生成报告失败: {e}")
            return False
    
    def get_results(self) -> List[Dict[str, Any]]:
        """
        获取处理结果
        
        Returns:
            List[Dict[str, Any]]: 处理结果列表
        """
        return self.results if self.results else []
    
    @staticmethod
    def _format_result(result: Any) -> str:
        """
        格式化结果为字符串
        """
        if isinstance(result, list):
            return ', '.join(str(x) for x in result)
        return str(result) if result is not None else ''

def main():
    """
    主函数，用于测试
    实际使用时，应该通过RPA调用ExcelAnalyzer类的方法
    """
    analyzer = ExcelAnalyzer()
    
    # 设置文件路径
    rule_file = 'rule_table.xlsx'
    data_file = 'data.xlsx'
    output_file = 'report.xlsx'
    
    # 检查文件是否存在
    if not Path(rule_file).exists():
        print(f"规则表文件不存在: {rule_file}")
        return
    if not Path(data_file).exists():
        print(f"数据文件不存在: {data_file}")
        return
        
    # 读取数据
    print("正在读取数据...")
    if not analyzer.load_rule_table(rule_file):
        return
    if not analyzer.load_excel_data(data_file):
        return
        
    # 处理规则
    print("正在处理规则...")
    if not analyzer.process_rules():
        return
    
    # 生成报告
    print("正在生成报告...")
    analyzer.generate_report(output_file)

if __name__ == '__main__':
    main()
