from typing import Dict, List, Union, Tuple
from utils import col_letter_to_index, parse_result_column

def extract_kpi_rule(sheet_data: List[List], location: str, comments: str, optimization_plan: str) -> Dict:
    col = col_letter_to_index(location[0])
    row = int(location[1:]) - 1
    
    try:
        value = float(sheet_data[row][col].strip('%'))
    except (ValueError, IndexError):
        return {
            'description': 'Inventory efficiency',
            'result': None,
            'comments': '数据无效',
            'optimization_plan': ''
        }
    
    comment_list = [c.strip() for c in comments.split('<br>')]
    plan_list = [p.strip() for p in optimization_plan.split('<br>')]
    
    if value < 80:
        idx = 0
    elif value <= 120:
        idx = 1
    else:
        idx = 2
        
    return {
        'description': 'Inventory efficiency',
        'result': f'{value}%',
        'comments': comment_list[idx] if idx < len(comment_list) else '',
        'optimization_plan': plan_list[idx] if idx < len(plan_list) else ''
    }

def extract_top_n_values(sheet_data: List[List], value_col: str, target_col: str, n: int, reverse: bool = True) -> List[str]:
    value_col_idx = col_letter_to_index(value_col)
    target_col_idx = col_letter_to_index(target_col)
    
    values = []
    for i, row in enumerate(sheet_data[1:], 1):
        try:
            val = float(row[value_col_idx].strip('%'))
            values.append((val, i))
        except (ValueError, IndexError):
            continue
            
    values.sort(reverse=reverse)
    top_n = values[:n]
    
    result = []
    for _, row_idx in top_n:
        try:
            val = sheet_data[row_idx][target_col_idx]
            if val:
                result.append(val)
        except IndexError:
            continue
            
    return result

def extract_rule_data(rule_row: Dict, data_dict: Dict[str, List[List]]) -> Dict:
    sheet_name = rule_row['Sheet']
    if sheet_name not in data_dict:
        return {
            'description': rule_row['Description'],
            'result': None,
            'comments': '工作表不存在',
            'optimization_plan': ''
        }
        
    sheet_data = data_dict[sheet_name]
    
    if rule_row['Description'] == 'Inventory efficiency':
        return extract_kpi_rule(
            sheet_data,
            rule_row['Location'],
            rule_row['Comments'],
            rule_row['Optimization plan']
        )
        
    if not rule_row['Rule']:
        try:
            col = col_letter_to_index(rule_row['Location'][0])
            row = int(rule_row['Location'][1:]) - 1
            return {
                'description': rule_row['Description'],
                'result': sheet_data[row][col],
                'comments': rule_row['Comments'],
                'optimization_plan': rule_row['Optimization plan']
            }
        except (ValueError, IndexError):
            return {
                'description': rule_row['Description'],
                'result': None,
                'comments': '位置无效',
                'optimization_plan': ''
            }
    
    target_col = parse_result_column(rule_row['Result'])
    if not target_col:
        return {
            'description': rule_row['Description'],
            'result': None,
            'comments': 'Result格式无效',
            'optimization_plan': ''
        }
    
    rule_text = rule_row['Rule']
    
    if '最低的5个料号' in rule_text:
        result = extract_top_n_values(sheet_data, 'AS', target_col, 5, False)
    elif '最高的5个料号' in rule_text:
        result = extract_top_n_values(sheet_data, 'AS', target_col, 5, True)
    elif '最大的三个料号' in rule_text:
        col = rule_text.split('列')[0].strip()
        result = extract_top_n_values(sheet_data, col, target_col, 3, True)
    elif '天数最多的三条' in rule_text:
        result = extract_top_n_values(sheet_data, 'D', target_col, 3, True)
    elif 'AO列数值最大的三个料号' in rule_text:
        result = extract_top_n_values(sheet_data, 'AO', target_col, 3, True)
    elif '最大' in rule_text and not any(x in rule_text for x in ['三个', '5个']):
        col = rule_text.split('列')[0].strip()
        result = extract_top_n_values(sheet_data, col, target_col, 1, True)
    else:
        return {
            'description': rule_row['Description'],
            'result': None,
            'comments': '未知的规则类型',
            'optimization_plan': ''
        }
        
    return {
        'description': rule_row['Description'],
        'result': result[0] if isinstance(result, list) and len(result) == 1 else result,
        'comments': rule_row['Comments'],
        'optimization_plan': rule_row['Optimization plan']
    }
