import pandas as pd
from pathlib import Path
from typing import Dict, List, Any
from rule_extractor import extract_rule_data

def load_excel_data(file_path: str) -> Dict[str, List[List[str]]]:
    excel_data = {}
    try:
        xlsx = pd.ExcelFile(file_path)
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name)
            excel_data[sheet_name] = df.values.tolist()
            excel_data[sheet_name].insert(0, df.columns.tolist())
    except Exception as e:
        print(f'读取Excel文件失败: {e}')
        return {}
    return excel_data

def load_rule_table(file_path: str) -> List[Dict[str, str]]:
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        rules = df.to_dict('records')
        return rules
    except Exception as e:
        print(f'读取规则表失败: {e}')
        return []

def process_rules(rules: List[Dict[str, str]], data_dict: Dict[str, List[List[str]]]) -> List[Dict[str, Any]]:
    results = []
    for rule in rules:
        try:
            result = extract_rule_data(rule, data_dict)
            results.append(result)
        except Exception as e:
            print(f'处理规则失败: {rule.get("Description", "未知规则")}, 错误: {e}')
            results.append({
                'description': rule.get('Description', '未知规则'),
                'result': None,
                'comments': f'处理失败: {str(e)}',
                'optimization_plan': ''
            })
    return results

def format_result(result: Any) -> str:
    if isinstance(result, list):
        return ', '.join(str(x) for x in result)
    return str(result) if result is not None else ''

def generate_report(results: List[Dict[str, Any]], output_file: str):
    df = pd.DataFrame(columns=['描述', '结果', '注释', '优化计划'])
    
    for result in results:
        df = pd.concat([df, pd.DataFrame([{
            '描述': result['description'],
            '结果': format_result(result['result']),
            '注释': result['comments'],
            '优化计划': result['optimization_plan']
        }])], ignore_index=True)
    
    try:
        df.to_excel(output_file, index=False, sheet_name='分析报告')
        print(f'报告已生成: {output_file}')
    except Exception as e:
        print(f'生成报告失败: {e}')

def main():
    rule_file = 'rule_table.xlsx'
    data_file = 'data.xlsx'
    output_file = 'report.xlsx'
    
    if not Path(rule_file).exists():
        print(f'规则表文件不存在: {rule_file}')
        return
    if not Path(data_file).exists():
        print(f'数据文件不存在: {data_file}')
        return
        
    print('正在读取数据...')
    rules = load_rule_table(rule_file)
    if not rules:
        print('没有找到规则')
        return
        
    data_dict = load_excel_data(data_file)
    if not data_dict:
        print('没有找到数据')
        return
        
    print('正在处理规则...')
    results = process_rules(rules, data_dict)
    
    print('正在生成报告...')
    generate_report(results, output_file)

if __name__ == '__main__':
    main()
