from typing import Dict, List, Any, Optional
import pandas as pd
from pathlib import Path
import logging

class ExcelAnalyzerControl:
    def __init__(self, excel_path: str, entity_name: str, week_number: int):
        """
        初始化Excel分析控制器
        
        Args:
            excel_path: Excel文件路径
            entity_name: 实体名称(如E88)
            week_number: 当前周数
        """
        self.excel_path = Path(excel_path)
        self.entity_name = entity_name
        self.week_number = week_number
        self.rules: Dict[str, Any] = {}
        self.excel: Optional[pd.ExcelFile] = None
        self.logger = logging.getLogger(__name__)
        
    def load_excel(self) -> bool:
        """加载Excel文件"""
        try:
            self.excel = pd.ExcelFile(self.excel_path)
            return True
        except Exception as e:
            self.logger.error(f"加载Excel文件失败: {str(e)}")
            return False
            
    def load_rules(self) -> bool:
        """从第一个Sheet加载规则表"""
        try:
            if not self.excel:
                return False
                
            rules_df = pd.read_excel(self.excel, sheet_name=0)
            
            # 验证必要的列是否存在
            required_columns = ['规则名称', 'Sheet名称', '单元格位置', '计算逻辑']
            missing_columns = [col for col in required_columns if col not in rules_df.columns]
            if missing_columns:
                self.logger.error(f"规则表缺少必要的列: {', '.join(missing_columns)}")
                return False
                
            # 获取所有可用的sheet名称
            available_sheets = set(self.excel.sheet_names)
            
            for _, row in rules_df.iterrows():
                rule_name = str(row['规则名称']).strip()
                sheet_name = str(row['Sheet名称']).strip()
                cell_position = str(row['单元格位置']).strip()
                
                # 验证规则名称
                if not rule_name:
                    self.logger.warning(f"跳过空规则名称的行")
                    continue
                    
                # 验证sheet名称
                if sheet_name not in available_sheets:
                    self.logger.warning(f"规则'{rule_name}'指定的Sheet'{sheet_name}'不存在，已跳过")
                    continue
                    
                # 验证单元格位置格式
                if not self._is_valid_cell_position(cell_position):
                    self.logger.warning(f"规则'{rule_name}'的单元格位置'{cell_position}'格式无效，已跳过")
                    continue
                    
                # 存储规则
                self.rules[rule_name] = {
                    'sheet_name': sheet_name,
                    'cell_position': cell_position,
                    'logic': str(row['计算逻辑']) if pd.notna(row['计算逻辑']) else ''
                }
                
            if not self.rules:
                self.logger.error("没有加载到任何有效的规则")
                return False
                
            self.logger.info(f"成功加载了{len(self.rules)}条规则")
            return True
            
        except Exception as e:
            self.logger.error(f"加载规则表失败: {str(e)}")
            return False
            
    def extract_kpi_data(self) -> Dict[str, Any]:
        """根据规则提取KPI数据"""
        results = {}
        try:
            for rule_name, rule_info in self.rules.items():
                try:
                    sheet_data = pd.read_excel(
                        self.excel, 
                        sheet_name=rule_info['sheet_name']
                    )
                    
                    # 解析单元格位置
                    col, row = self._parse_cell_position(rule_info['cell_position'])
                    
                    # 验证索引是否有效
                    if row >= len(sheet_data) or col >= len(sheet_data.columns):
                        self.logger.warning(f"规则'{rule_name}'的单元格位置超出范围，已跳过")
                        continue
                    
                    # 获取数据
                    value = sheet_data.iloc[row, col]
                    
                    # 处理空值
                    if pd.isna(value):
                        self.logger.warning(f"规则'{rule_name}'的单元格值为空，已跳过")
                        continue
                        
                    # 应用计算逻辑(如果有)
                    if rule_info['logic']:
                        try:
                            value = self._apply_logic(value, rule_info['logic'])
                        except Exception as e:
                            self.logger.warning(f"规则'{rule_name}'的计算逻辑应用失败: {str(e)}")
                            continue
                            
                    # 格式化值
                    formatted_value = self._format_value(value)
                    if formatted_value is None:
                        self.logger.warning(f"规则'{rule_name}'的值格式化失败，已跳过")
                        continue
                        
                    results[rule_name] = formatted_value
                    
                except Exception as e:
                    self.logger.warning(f"处理规则'{rule_name}'时发生错误: {str(e)}")
                    continue
                    
        except Exception as e:
            self.logger.error(f"提取KPI数据失败: {str(e)}")
            
        return results
        
    def _parse_cell_position(self, position: str) -> tuple[int, int]:
        """
        解析Excel单元格位置(如'A1'或'AA1'转为行列索引)
        """
        from openpyxl.utils import column_index_from_string
        # 分离列字母和行号
        import re
        match = re.match(r'^([A-Z]+)([1-9][0-9]*)$', position)
        if not match:
            raise ValueError(f"无效的单元格位置格式: {position}")
        
        col_str, row_str = match.groups()
        # 使用openpyxl的工具函数处理列字母
        col = column_index_from_string(col_str) - 1
        row = int(row_str) - 1
        return col, row
        
    def _format_value(self, value: Any) -> Optional[Any]:
        """
        格式化值
        支持：
        - 百分比（自动检测）
        - 货币值（大于1000的数字）
        - 科学记数法（非常大或非常小的数字）
        - 一般数字（保留2位小数）
        """
        try:
            if isinstance(value, (int, float)):
                abs_value = abs(value)
                if abs_value == 0:
                    return 0
                elif abs_value < 0.0001:  # 非常小的数字
                    return f"{value:.2e}"
                elif abs_value > 1e6:     # 非常大的数字
                    return f"{value:.2e}"
                elif abs_value > 1000:    # 可能是货币值
                    return round(value, 2)  # 保留原始数值，在报告生成时格式化
                else:
                    return round(value, 2)
            elif isinstance(value, str):
                # 处理百分比
                if '%' in value:
                    value = value.strip('%')
                    try:
                        num = float(value) / 100
                        if abs(num) < 0.0001:
                            return f"{num:.2e}"
                        return round(num, 4)
                    except ValueError:
                        return value
                # 处理可能的数字字符串
                try:
                    num = float(value)
                    return self._format_value(num)  # 递归处理数字
                except ValueError:
                    return value.strip()  # 清理字符串两端的空白
            return value
        except Exception:
            return None
            
    def _format_display_value(self, value: Any) -> str:
        """
        格式化用于显示的值
        """
        if isinstance(value, (int, float)):
            abs_value = abs(value)
            if abs_value == 0:
                return "0"
            elif abs_value < 0.0001 or abs_value > 1e6:
                return f"{value:.2e}"
            elif abs_value < 1:  # 可能是百分比
                return f"{value*100:.2f}%"
            elif abs_value > 1000:  # 货币值
                return f"{value:,.2f}"
            else:
                return f"{value:.2f}"
        return str(value)
            
    def _apply_logic(self, value: Any, logic: str) -> Any:
        """应用计算逻辑"""
        try:
            # 支持基本的数学运算
            if isinstance(value, (int, float)):
                # 创建安全的局部变量环境
                local_vars = {'x': value}
                # 替换常见的数学表达式
                logic = logic.replace('value', 'x')
                # 执行计算
                result = eval(logic, {"__builtins__": {}}, local_vars)
                return result
            return value
        except Exception as e:
            raise ValueError(f"计算逻辑执行失败: {str(e)}")
        
    def generate_report(self, kpi_data: Dict[str, Any]) -> str:
        """生成报告"""
        from datetime import datetime
        
        # 获取当前时间
        now = datetime.now()
        
        # 报告头部
        report_lines = [
            "="*50,
            "采购数据分析报告",
            "="*50,
            "",
            f"生成时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
            f"实体: {self.entity_name}",
            f"周数: {self.week_number}",
            f"数据源: {self.excel_path.name}",
            "",
            "-"*50,
            "KPI指标分析结果:",
            "-"*50,
            ""
        ]
        
        # 如果没有数据
        if not kpi_data:
            report_lines.extend([
                "警告: 未能提取到任何KPI数据",
                "请检查规则配置和数据源是否正确"
            ])
            return "\n".join(report_lines)
        
        # 添加KPI数据
        for rule_name, value in kpi_data.items():
            # 格式化值的显示
            formatted_value = self._format_display_value(value)
                
            report_lines.extend([
                f"规则: {rule_name}",
                f"值: {formatted_value}",
                "-"*30,
                ""
            ])
        
        # 报告尾部
        report_lines.extend([
            "="*50,
            "报告生成完成",
            "="*50
        ])
        
        return "\n".join(report_lines)
        
    def save_report(self, report_content: str, output_path: str) -> bool:
        """保存报告到文件"""
        try:
            output_path = Path(output_path)
            
            # 确保输出目录存在
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 写入报告内容
            output_path.write_text(report_content, encoding='utf-8')
            
            self.logger.info(f"报告已保存到: {output_path}")
            return True
        except Exception as e:
            self.logger.error(f"保存报告失败: {str(e)}")
            return False

    def _is_valid_cell_position(self, position: str) -> bool:
        """
        验证单元格位置格式是否有效
        支持：
        - 单字母列：A1, B2, Z10
        - 双字母列：AA1, AB2, ZZ100
        """
        import re
        # 修改正则表达式以支持多字母列
        return bool(re.match(r'^[A-Z]{1,3}[1-9][0-9]*$', position))

class RuleValidator:
    """规则验证器"""
    
    @staticmethod
    def validate_rule(rule: Dict[str, Any]) -> tuple[bool, str]:
        """
        验证规则的完整性和有效性
        
        Returns:
            tuple[bool, str]: (是否有效, 错误信息)
        """
        # 验证必要字段
        required_fields = {'sheet_name', 'cell_position', 'logic'}
        missing_fields = required_fields - set(rule.keys())
        if missing_fields:
            return False, f"缺少必要字段: {', '.join(missing_fields)}"
            
        # 验证字段类型
        if not isinstance(rule['sheet_name'], str):
            return False, "Sheet名称必须是字符串"
        if not isinstance(rule['cell_position'], str):
            return False, "单元格位置必须是字符串"
        if not isinstance(rule['logic'], str):
            return False, "计算逻辑必须是字符串"
            
        # 验证计算逻辑语法
        if rule['logic']:
            try:
                # 测试编译逻辑表达式
                compile(rule['logic'], '<string>', 'eval')
            except SyntaxError as e:
                return False, f"计算逻辑语法错误: {str(e)}"
                
        return True, ""
        
    @staticmethod
    def validate_cell_value(value: Any, rule_name: str) -> tuple[bool, str]:
        """
        验证单元格值是否符合规则要求
        """
        # 根据规则名称确定期望的数据类型
        if "效率" in rule_name or "风险" in rule_name:
            if not isinstance(value, (int, float)):
                return False, f"{rule_name}期望数值类型，实际是{type(value)}"
            if value < 0:
                return False, f"{rule_name}不应为负值"
                
        elif "金额" in rule_name:
            if not isinstance(value, (int, float)):
                return False, f"{rule_name}期望数值类型，实际是{type(value)}"
            if value < 0:
                return False, f"{rule_name}不应为负值"
                
        elif "天数" in rule_name:
            if not isinstance(value, (int, float)):
                return False, f"{rule_name}期望数值类型，实际是{type(value)}"
            if value < 0:
                return False, f"{rule_name}不应为负值"
            if value > 365:
                return False, f"{rule_name}超过365天，可能有误"
                
        return True, ""
        
    @staticmethod
    def validate_threshold(value: float, rule_name: str) -> tuple[bool, str]:
        """
        验证值是否在合理范围内
        """
        thresholds = {
            "库存效率": (0, 200),    # 0-200%
            "缺料风险": (0, 100),    # 0-100%
            "呆滞风险": (0, 100),    # 0-100%
            "运输天数": (0, 60),     # 0-60天
            "安全时间": (0, 90),     # 0-90天
            "MOQ影响": (0, 1000000)  # 0-100万
        }
        
        for key, (min_val, max_val) in thresholds.items():
            if key in rule_name:
                if not min_val <= value <= max_val:
                    return False, f"{rule_name}的值{value}超出合理范围[{min_val}, {max_val}]"
                    
        return True, ""

class RuleManager:
    """规则管理器"""
    
    def __init__(self):
        self.rules: Dict[str, Dict[str, Any]] = {}
        self.rule_dependencies: Dict[str, List[str]] = {}
        self.execution_order: List[str] = []
        
    def add_rule(self, rule_name: str, rule_info: Dict[str, Any]) -> tuple[bool, str]:
        """
        添加规则
        
        Args:
            rule_name: 规则名称
            rule_info: 规则信息
            
        Returns:
            tuple[bool, str]: (是否成功, 错误信息)
        """
        # 验证规则名称唯一性
        if rule_name in self.rules:
            return False, f"规则名称'{rule_name}'已存在"
            
        # 验证规则格式
        is_valid, error_msg = RuleValidator.validate_rule(rule_info)
        if not is_valid:
            return False, error_msg
            
        # 分析规则依赖
        dependencies = self._extract_dependencies(rule_info['logic'])
        
        # 检查循环依赖
        if self._has_circular_dependency(rule_name, dependencies):
            return False, f"规则'{rule_name}'存在循环依赖"
            
        # 存储规则
        self.rules[rule_name] = rule_info
        self.rule_dependencies[rule_name] = dependencies
        
        # 更新执行顺序
        self._update_execution_order()
        
        return True, ""
        
    def _extract_dependencies(self, logic: str) -> List[str]:
        """从计算逻辑中提取依赖的规则名称"""
        dependencies = []
        if not logic:
            return dependencies
            
        # 查找逻辑中引用的其他规则
        for rule_name in self.rules:
            if rule_name in logic:
                dependencies.append(rule_name)
                
        return dependencies
        
    def _has_circular_dependency(self, new_rule: str, dependencies: List[str]) -> bool:
        """检查是否存在循环依赖"""
        visited = set()
        
        def dfs(rule: str) -> bool:
            if rule in visited:
                return True
            if rule not in self.rule_dependencies:
                return False
                
            visited.add(rule)
            for dep in self.rule_dependencies[rule]:
                if dfs(dep):
                    return True
            visited.remove(rule)
            return False
            
        # 临时添加新规则的依赖
        temp_deps = self.rule_dependencies.copy()
        temp_deps[new_rule] = dependencies
        
        # 从新规则开始检查
        return dfs(new_rule)
        
    def _update_execution_order(self):
        """更新规则执行顺序（拓扑排序）"""
        # 重置执行顺序
        self.execution_order = []
        visited = set()
        temp = set()
        
        def visit(rule: str):
            if rule in temp:
                raise ValueError(f"检测到循环依赖: {rule}")
            if rule in visited:
                return
                
            temp.add(rule)
            for dep in self.rule_dependencies.get(rule, []):
                visit(dep)
            temp.remove(rule)
            visited.add(rule)
            self.execution_order.append(rule)
            
        # 对所有规则进行拓扑排序
        for rule in self.rules:
            if rule not in visited:
                visit(rule)
                
    def get_execution_order(self) -> List[str]:
        """获取规则执行顺序"""
        return self.execution_order.copy()
        
    def get_rule(self, rule_name: str) -> Optional[Dict[str, Any]]:
        """获取规则信息"""
        return self.rules.get(rule_name)

class DataManager:
    """数据管理器"""
    
    def __init__(self, excel_file: pd.ExcelFile):
        self.excel_file = excel_file
        self.sheet_cache: Dict[str, pd.DataFrame] = {}
        self.value_cache: Dict[str, Any] = {}
        
    def get_sheet_data(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """获取工作表数据（带缓存）"""
        try:
            if sheet_name not in self.sheet_cache:
                self.sheet_cache[sheet_name] = pd.read_excel(
                    self.excel_file, 
                    sheet_name=sheet_name
                )
            return self.sheet_cache[sheet_name]
        except Exception as e:
            logging.error(f"读取工作表'{sheet_name}'失败: {str(e)}")
            return None
            
    def get_cell_value(self, sheet_name: str, cell_position: str) -> Optional[Any]:
        """获取单元格值（带缓存）"""
        cache_key = f"{sheet_name}:{cell_position}"
        
        if cache_key in self.value_cache:
            return self.value_cache[cache_key]
            
        try:
            sheet_data = self.get_sheet_data(sheet_name)
            if sheet_data is None:
                return None
                
            # 解析单元格位置
            from openpyxl.utils import column_index_from_string
            match = re.match(r'^([A-Z]+)([1-9][0-9]*)$', cell_position)
            if not match:
                raise ValueError(f"无效的单元格位置: {cell_position}")
                
            col_str, row_str = match.groups()
            col = column_index_from_string(col_str) - 1
            row = int(row_str) - 1
            
            # 验证索引
            if row >= len(sheet_data) or col >= len(sheet_data.columns):
                raise ValueError(f"单元格位置{cell_position}超出范围")
                
            # 获取值
            value = sheet_data.iloc[row, col]
            
            # 缓存值
            self.value_cache[cache_key] = value
            
            return value
        except Exception as e:
            logging.error(f"获取单元格值失败({cache_key}): {str(e)}")
            return None
            
    def get_column_data(self, sheet_name: str, column: str) -> Optional[pd.Series]:
        """获取列数据"""
        try:
            sheet_data = self.get_sheet_data(sheet_name)
            if sheet_data is None:
                return None
                
            col_idx = column_index_from_string(column) - 1
            return sheet_data.iloc[:, col_idx]
        except Exception as e:
            logging.error(f"获取列数据失败({sheet_name}:{column}): {str(e)}")
            return None
            
    def clear_cache(self):
        """清除缓存"""
        self.sheet_cache.clear()
        self.value_cache.clear()

class ReportGenerator:
    """报告生成器"""
    
    def __init__(self, entity_name: str, week_number: int):
        self.entity_name = entity_name
        self.week_number = week_number
        self.thresholds = {
            "库存效率": {"good": 90, "warning": 70},
            "缺料风险": {"high": 80, "medium": 50},
            "呆滞风险": {"high": 70, "medium": 40},
            "运输天数": {"high": 30, "medium": 15},
            "安全时间": {"high": 60, "medium": 30},
            "MOQ影响": {"high": 500000, "medium": 100000}
        }
        
    def generate_report(self, kpi_data: Dict[str, Any], source_file: str) -> str:
        """
        生成分析报告
        
        Args:
            kpi_data: KPI数据
            source_file: 数据源文件名
            
        Returns:
            报告文本
        """
        from datetime import datetime
        
        # 报告头部
        report_lines = [
            "="*80,
            f"{'采购数据分析报告':^80}",
            "="*80,
            "",
            f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"分析实体: {self.entity_name}",
            f"分析周数: 第{self.week_number}周",
            f"数据来源: {source_file}",
            "",
            "-"*80,
            "分析结果概述:",
            "-"*80,
            ""
        ]
        
        if not kpi_data:
            report_lines.extend([
                "警告: 未能提取到任何KPI数据",
                "建议:",
                "1. 检查数据源文件格式是否正确",
                "2. 验证规则配置是否有效",
                "3. 确认所需数据是否完整"
            ])
            return "\n".join(report_lines)
            
        # 添加概述
        summary = self._generate_summary(kpi_data)
        report_lines.extend(summary)
        report_lines.extend(["", "-"*80, "详细分析:", "-"*80, ""])
        
        # 添加详细分析
        for rule_name, value in kpi_data.items():
            analysis = self._analyze_kpi(rule_name, value)
            report_lines.extend(analysis)
            report_lines.extend(["", "-"*30, ""])
            
        # 添加建议
        recommendations = self._generate_recommendations(kpi_data)
        report_lines.extend([
            "-"*80,
            "改进建议:",
            "-"*80,
            ""
        ])
        report_lines.extend(recommendations)
        
        # 报告尾部
        report_lines.extend([
            "",
            "="*80,
            f"报告生成完成 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "="*80
        ])
        
        return "\n".join(report_lines)
        
    def _generate_summary(self, kpi_data: Dict[str, Any]) -> List[str]:
        """生成概述"""
        summary = []
        
        # 计算关键指标状态
        status_count = {"good": 0, "warning": 0, "bad": 0}
        for rule_name, value in kpi_data.items():
            status = self._get_status(rule_name, value)
            status_count[status] += 1
            
        total = len(kpi_data)
        summary.extend([
            f"分析指标总数: {total}",
            f"状态良好: {status_count['good']} ({status_count['good']/total*100:.1f}%)",
            f"需要关注: {status_count['warning']} ({status_count['warning']/total*100:.1f}%)",
            f"问题严重: {status_count['bad']} ({status_count['bad']/total*100:.1f}%)",
            ""
        ])
        
        # 添加重要发现
        findings = self._identify_key_findings(kpi_data)
        if findings:
            summary.extend(["重要发现:"])
            summary.extend(findings)
            
        return summary
        
    def _analyze_kpi(self, rule_name: str, value: Any) -> List[str]:
        """分析单个KPI"""
        analysis = [
            f"指标: {rule_name}",
            f"当前值: {self._format_value(value, rule_name)}"
        ]
        
        # 添加状态评估
        status = self._get_status(rule_name, value)
        status_text = {
            "good": "良好 ✓",
            "warning": "警告 ⚠",
            "bad": "严重 ✗"
        }
        analysis.append(f"状态: {status_text[status]}")
        
        # 添加具体分析
        analysis.extend(self._get_specific_analysis(rule_name, value))
        
        return analysis
        
    def _get_status(self, rule_name: str, value: float) -> str:
        """获取KPI状态"""
        if not isinstance(value, (int, float)):
            return "warning"
            
        thresholds = self.thresholds.get(rule_name, {})
        if not thresholds:
            return "warning"
            
        if "库存效率" in rule_name:
            if value >= thresholds["good"]:
                return "good"
            elif value >= thresholds["warning"]:
                return "warning"
            return "bad"
        elif any(key in rule_name for key in ["风险", "天数", "影响"]):
            if value >= thresholds["high"]:
                return "bad"
            elif value >= thresholds["medium"]:
                return "warning"
            return "good"
            
        return "warning"
        
    def _format_value(self, value: Any, rule_name: str) -> str:
        """格式化值"""
        if not isinstance(value, (int, float)):
            return str(value)
            
        if "效率" in rule_name or "风险" in rule_name:
            return f"{value*100:.1f}%"
        elif "金额" in rule_name:
            return f"¥{value:,.2f}"
        elif "天数" in rule_name:
            return f"{value:.1f}天"
        else:
            return f"{value:,}"
            
    def _get_specific_analysis(self, rule_name: str, value: Any) -> List[str]:
        """获取具体分析"""
        analysis = []
        
        if "库存效率" in rule_name:
            if value < 70:
                analysis.append("库存水平过低，可能影响供应链稳定性")
            elif value > 120:
                analysis.append("库存水平过高，可能导致资金积压")
                
        elif "缺料风险" in rule_name:
            if value > 80:
                analysis.append("缺料风险高，需要立即关注")
            elif value > 50:
                analysis.append("存在潜在缺料风险，建议提前准备")
                
        elif "呆滞风险" in rule_name:
            if value > 70:
                analysis.append("呆滞风险高，需要制定清理计划")
            elif value > 40:
                analysis.append("呆滞风险上升，建议关注库存周转")
                
        elif "运输天数" in rule_name:
            if value > 30:
                analysis.append("运输时间过长，影响供应链效率")
            elif value > 15:
                analysis.append("运输时间偏长，建议优化物流方案")
                
        return analysis
        
    def _identify_key_findings(self, kpi_data: Dict[str, Any]) -> List[str]:
        """识别重要发现"""
        findings = []
        
        # 分析库存效率
        if "库存效率" in kpi_data:
            efficiency = kpi_data["库存效率"]
            if efficiency < 60:
                findings.append("⚠ 库存效率严重不足，需要立即改善")
            elif efficiency > 150:
                findings.append("⚠ 库存积压严重，建议及时处理")
                
        # 分析风险指标
        risk_indicators = {k: v for k, v in kpi_data.items() if "风险" in k}
        high_risks = [k for k, v in risk_indicators.items() if v > 80]
        if high_risks:
            findings.append(f"⚠ 以下指标风险较高: {', '.join(high_risks)}")
            
        return findings
        
    def _generate_recommendations(self, kpi_data: Dict[str, Any]) -> List[str]:
        """生成改进建议"""
        recommendations = []
        
        # 库存优化建议
        if "库存效率" in kpi_data:
            efficiency = kpi_data["库存效率"]
            if efficiency < 70:
                recommendations.extend([
                    "1. 库存优化建议:",
                    "   - 检查安全库存水平设置",
                    "   - 优化补货策略和频率",
                    "   - 加强与供应商的沟通"
                ])
            elif efficiency > 120:
                recommendations.extend([
                    "1. 库存优化建议:",
                    "   - 制定库存清理计划",
                    "   - 调整采购订单量",
                    "   - 考虑供应商寄售模式"
                ])
                
        # 风险管理建议
        risk_recommendations = []
        if any(v > 70 for k, v in kpi_data.items() if "风险" in k):
            risk_recommendations.extend([
                "2. 风险管理建议:",
                "   - 建立风险预警机制",
                "   - 制定应急响应方案",
                "   - 定期评估供应商表现"
            ])
            
        if risk_recommendations:
            recommendations.extend(risk_recommendations)
            
        # 运营改进建议
        if "运输天数" in kpi_data and kpi_data["运输天数"] > 20:
            recommendations.extend([
                "3. 运营改进建议:",
                "   - 优化物流路线",
                "   - 考虑增加物流供应商",
                "   - 建立运输时效监控"
            ])
            
        return recommendations