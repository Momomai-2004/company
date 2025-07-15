from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
from pathlib import Path
import logging
import re

class ExcelAnalyzerControl:
    """Excel分析RPA控件
    
    主要功能：
    1. 从Excel文件读取规则
    2. 执行数据分析
    3. 返回格式化结果
    """
    
    def __init__(self):
        self.rules_df: Optional[pd.DataFrame] = None
        self.data_df: Optional[pd.DataFrame] = None
        self.logger = self._setup_logger()
        
    def _setup_logger(self) -> logging.Logger:
        """设置日志"""
        logger = logging.getLogger("ExcelAnalyzerControl")
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)
        return logger
    
    def load_excel(self, file_path: str, rules_sheet: str = "规则表") -> bool:
        """加载Excel文件
        
        Args:
            file_path: Excel文件路径
            rules_sheet: 规则表sheet名称
            
        Returns:
            bool: 是否加载成功
        """
        try:
            # 读取所有sheet
            xlsx = pd.ExcelFile(file_path)
            self.data_df = {}
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name)
                self.data_df[sheet_name] = df
                
            # 读取规则表
            self.rules_df = self.data_df[rules_sheet]
            self.logger.info(f"成功加载规则表，共{len(self.rules_df)}条规则")
            return True
        except Exception as e:
            self.logger.error(f"加载Excel文件失败: {str(e)}")
            return False
    
    def find_rules_by_pattern(self, pattern: str) -> List[Dict[str, Any]]:
        """根据模式查找规则
        
        Args:
            pattern: 查找模式（支持正则表达式）
            
        Returns:
            List[Dict]: 匹配的规则列表
        """
        if self.rules_df is None:
            return []
            
        matched_rules = []
        try:
            regex = re.compile(pattern, re.IGNORECASE)
            for idx, row in self.rules_df.iterrows():
                # 在Description和Rule中查找
                if (regex.search(str(row.get('Description', ''))) or 
                    regex.search(str(row.get('Rule', '')))):
                    matched_rules.append({
                        'id': idx,
                        'sheet': row['Sheet'],
                        'location': row['Location'],
                        'description': row['Description'],
                        'rule_type': row['Rule']
                    })
        except Exception as e:
            self.logger.error(f"查找规则失败: {str(e)}")
            
        return matched_rules
            
    def analyze_by_rule(self, rule_id: int) -> Dict[str, Any]:
        """根据规则ID执行分析
        
        Args:
            rule_id: 规则ID（规则表中的行号）
            
        Returns:
            Dict: 分析结果
        """
        try:
            if self.rules_df is None:
                raise ValueError("请先加载Excel文件")
                
            rule = self.rules_df.iloc[rule_id]
            sheet_name = rule["Sheet"]
            location = rule["Location"]
            rule_type = rule["Rule"]
            
            # 获取数据
            if sheet_name not in self.data_df:
                raise ValueError(f"找不到工作表: {sheet_name}")
                
            data = self._get_cell_data(sheet_name, location)
            
            # 根据规则类型执行不同的分析
            if rule_type == "KPI":
                return self._analyze_kpi(rule, data)
            elif rule_type == "最值":
                return self._analyze_extremum(rule, data)
            elif rule_type == "单元格":
                return self._analyze_cell(rule, data)
            else:
                raise ValueError(f"不支持的规则类型: {rule_type}")
                
        except Exception as e:
            self.logger.error(f"规则{rule_id}分析失败: {str(e)}")
            return {"success": False, "error": str(e)}
    
    def _get_cell_data(self, sheet_name: str, location: str) -> Any:
        """获取单元格数据"""
        try:
            # 解析位置（例如：A1, B2等）
            col = ord(location[0].upper()) - ord('A')
            row = int(location[1:]) - 1
            return self.data_df[sheet_name].iloc[row, col]
        except Exception as e:
            raise ValueError(f"无法获取位置{location}的数据: {str(e)}")
            
    def _analyze_kpi(self, rule: pd.Series, data: Any) -> Dict[str, Any]:
        """KPI规则分析"""
        try:
            # 获取KPI阈值（假设在Comments中定义）
            thresholds = self._parse_kpi_thresholds(rule.get('Comments', ''))
            
            # 分析数据
            status = self._evaluate_kpi(data, thresholds)
            
            return {
                "success": True,
                "type": "KPI",
                "value": data,
                "status": status,
                "thresholds": thresholds
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
            
    def _parse_kpi_thresholds(self, comments: str) -> Dict[str, float]:
        """解析KPI阈值
        预期格式：good: >90, warning: >70, bad: <=70
        """
        thresholds = {}
        try:
            parts = comments.split(',')
            for part in parts:
                level, value = part.split(':')
                level = level.strip()
                value = float(re.search(r'[0-9.]+', value.strip()).group())
                thresholds[level] = value
        except:
            # 使用默认值
            thresholds = {
                "good": 90,
                "warning": 70,
                "bad": 0
            }
        return thresholds
            
    def _evaluate_kpi(self, value: float, thresholds: Dict[str, float]) -> str:
        """评估KPI状态"""
        try:
            value = float(value)
            if value >= thresholds.get("good", 90):
                return "良好"
            elif value >= thresholds.get("warning", 70):
                return "警告"
            else:
                return "不达标"
        except:
            return "数据无效"
            
    def _analyze_extremum(self, rule: pd.Series, data: Any) -> Dict[str, Any]:
        """最值分析"""
        try:
            sheet_name = rule["Sheet"]
            df = self.data_df[sheet_name]
            
            # 获取列数据
            col = ord(rule["Location"][0].upper()) - ord('A')
            values = df.iloc[:, col].dropna()
            
            # 获取最值数量（默认3个）
            n = 3
            try:
                n = int(re.search(r'\d+', rule.get('Comments', '3')).group())
            except:
                pass
                
            return {
                "success": True,
                "type": "最值",
                "max_values": values.nlargest(n).tolist(),
                "min_values": values.nsmallest(n).tolist()
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
            
    def _analyze_cell(self, rule: pd.Series, data: Any) -> Dict[str, Any]:
        """单元格分析"""
        try:
            return {
                "success": True,
                "type": "单元格",
                "value": data,
                "formatted_value": self._format_cell_value(data)
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def _format_cell_value(self, value: Any) -> str:
        """格式化单元格值"""
        if pd.isna(value):
            return "空值"
        elif isinstance(value, (int, float)):
            return f"{value:,.2f}"
        return str(value)
            
    def get_all_rules(self) -> List[Dict[str, Any]]:
        """获取所有规则的描述"""
        if self.rules_df is None:
            return []
            
        rules = []
        for idx, row in self.rules_df.iterrows():
            rules.append({
                "id": idx,
                "sheet": row["Sheet"],
                "location": row["Location"],
                "description": row["Description"],
                "rule_type": row["Rule"]
            })
        return rules
        
    def analyze_all(self) -> List[Dict[str, Any]]:
        """分析所有规则"""
        if self.rules_df is None:
            return []
            
        results = []
        for idx in range(len(self.rules_df)):
            result = self.analyze_by_rule(idx)
            results.append(result)
        return results
        
    def generate_report(self, output_file: str, results: Optional[List[Dict[str, Any]]] = None) -> bool:
        """生成分析报告
        
        Args:
            output_file: 输出文件路径
            results: 分析结果（如果为None则使用analyze_all的结果）
            
        Returns:
            bool: 是否成功
        """
        try:
            if results is None:
                results = self.analyze_all()
                
            # 创建报告数据
            report_data = []
            for idx, result in enumerate(results):
                rule = self.rules_df.iloc[idx]
                report_data.append({
                    '规则ID': idx,
                    '描述': rule['Description'],
                    '工作表': rule['Sheet'],
                    '位置': rule['Location'],
                    '规则类型': rule['Rule'],
                    '分析结果': self._format_result(result),
                    '状态': '成功' if result.get('success', False) else '失败',
                    '备注': rule.get('Comments', '')
                })
                
            # 生成Excel报告
            df = pd.DataFrame(report_data)
            df.to_excel(output_file, index=False, sheet_name='分析报告')
            self.logger.info(f"报告已生成: {output_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"生成报告失败: {str(e)}")
            return False
            
    def _format_result(self, result: Dict[str, Any]) -> str:
        """格式化分析结果"""
        if not result.get('success', False):
            return f"错误: {result.get('error', '未知错误')}"
            
        result_type = result.get('type')
        if result_type == 'KPI':
            return f"值: {result['value']}, 状态: {result['status']}"
        elif result_type == '最值':
            return f"最大值: {result['max_values']}, 最小值: {result['min_values']}"
        elif result_type == '单元格':
            return str(result['formatted_value'])
        else:
            return str(result) 