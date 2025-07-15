from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
from pathlib import Path
import logging
import re

class ExcelAnalyzerControl:
    """
    Excel 分析 RPA 控件
    --------------------
    该类封装了库存健康分析的整个流程，可被 RPA 无缝调用。

    设计要点：
    1. **规则驱动**：所有分析逻辑依赖规则表（第一个 sheet），无需硬编码。
    2. **单文件部署**：控件内聚在一个文件，RPA 侧只需 `import excel_analyzer_control`。
    3. **易扩展**：新增业务规则 → 添加 `_analyze_*` 方法并在 `analyze_by_rule` 分派即可。
    4. **易追溯**：统一 `self.logger` 记录运行步骤与异常，方便排障。

    典型调用示例::

        from excel_analyzer_control import ExcelAnalyzerControl

        analyzer = ExcelAnalyzerControl()
        analyzer.load_excel('data_with_rules.xlsx')
        results = analyzer.analyze_all()
        analyzer.generate_report('analysis_report.xlsx', results)

    当前已内置 8 条典型业务规则，详见各 `_analyze_*` 方法的实现。
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
            
    def _col_letter_to_index(self, letters: str) -> int:
        """将Excel列字母转换为0索引的列号"""
        idx = 0
        for ch in letters.upper():
            if not ('A' <= ch <= 'Z'):
                continue
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1

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
            description = str(rule.get("Description", "")).lower()

            # 针对具体业务规则的分派
            if "inventory efficiency" in description or "库存效率" in description:
                return self._analyze_inventory_efficiency(rule)
            elif "shortage risk" in description or "缺料" in description:
                return self._analyze_shortage_risk(rule)
            elif "dead stock risk" in description or "呆滞风险" in description:
                return self._analyze_dead_stock_risk(rule)
            elif "impact of \"dead\" stock" in description or "金额最大的呆滞物料" in description:
                return self._analyze_dead_stock_value(rule)
            elif "days of transit" in description or "运输天数" in description:
                return self._analyze_transit_inconsistencies(rule)
            elif "safety time" in description or "安全时间" in description:
                return self._analyze_safety_time(rule)
            elif "no supplier" in description or "没有供应商" in description:
                return self._analyze_no_supplier(rule)
            elif "moq" in description:
                return self._analyze_moq_impact(rule)

            # 回退到通用规则类型解析
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

        t = result.get('type')
        if t in {"缺料风险", "呆滞风险", "呆滞金额", "运输天数不一致", "安全时间过长", "MOQ影响"}:
            return "; ".join([f"{pn}: {val}" if isinstance(val, (int, float)) else str(pn) for pn, val in result.get('items', [])])
        elif t == "无供应商物料":
            return ", ".join(map(str, result.get('items', [])))
        elif t == "库存效率":
            return f"{result['value']}% - {result['status']}"
        elif t == "KPI":
            return f"值: {result['value']}, 状态: {result['status']}"
        elif t == "最值":
            return f"最大值: {result['max_values']}, 最小值: {result['min_values']}"
        elif t == "单元格":
            return str(result['formatted_value'])
        else:
            return str(result) 

    def _analyze_inventory_efficiency(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 1：库存效率 (Inventory efficiency)

        数据来源：KPI 工作表 G10 单元格（百分比数值）
        判断逻辑：
            <  80  → "整体库存水平较低"
            80–120 → "整体库存水平合理"
            > 120  → "整体库存水平较高"

        返回字段：
            success   : bool  是否成功
            type      : str   "库存效率"
            value     : float 原始百分比数值
            status    : str   文本状态描述
            suggestion: str   优化建议
        """
        sheet_name = rule["Sheet"]
        value = self._get_cell_data(sheet_name, rule["Location"])
        try:
            value_f = float(value)
        except Exception:
            return {"success": False, "error": f"无法将{value}转换为数值"}

        if value_f < 80:
            status = "整体库存水平较低"
        elif 80 <= value_f <= 120:
            status = "整体库存水平合理"
        else:
            status = "整体库存水平较高"

        suggestion = "检查基础参数准确性；针对水平较低深入分析缺料，较高则优化采购计划"
        return {
            "success": True,
            "type": "库存效率",
            "value": value_f,
            "status": status,
            "suggestion": suggestion
        }

    def _fetch_top_with_pn(self, df: pd.DataFrame, value_col: str, pn_col: str, n: int, largest: bool = True) -> List[Tuple[Any, Any]]:
        # 将列字母转换为 DataFrame 的数字索引
        v_idx = self._col_letter_to_index(value_col)  # 指标列（如 AS / G / AO）
        pn_idx = self._col_letter_to_index(pn_col)    # 料号列（如 E / A / AH）

        # 取出指标列数据；`series` 与原 DataFrame 行索引保持一致
        series = df.iloc[:, v_idx]

        # 根据 `largest` 控制方向，取前 n 个极值索引
        sorter = series.nlargest(n) if largest else series.nsmallest(n)

        # 将极值行索引映射为 (PN, 数值) 二元组
        rows = sorter.index
        result: List[Tuple[Any, Any]] = []
        for idx in rows:
            pn = df.iloc[idx, pn_idx]   # 料号
            val = series.loc[idx]       # 指标数值
            result.append((pn, val))
        return result

    def _analyze_shortage_risk(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 2：缺料风险 (Shortage risk)

        数据来源：Results 工作表 AS 列（百分比）
        处理逻辑：提取 AS 列 *最小* 的 5 行，代表缺料风险最高。
        料号列：E 列。

        返回字段：
            items = [(pn, value), ...] 5 个元组
            suggestion：优化建议
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        # AS列百分比最低的5个料号, 料号在E列
        top = self._fetch_top_with_pn(df, "AS", "E", 5, largest=False)
        suggestion = "验证SAP参数，检查需求真实性，寻找替代料方案"
        return {
            "success": True,
            "type": "缺料风险",
            "items": top,
            "suggestion": suggestion
        }

    def _analyze_dead_stock_risk(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 3：呆滞风险 (Dead stock risk)

        数据来源：Results 工作表 AS 列（百分比）
        处理逻辑：提取 AS 列 *最大* 的 5 行，表示呆滞比例最高。
        料号列：E 列。
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        # AS列百分比最高的5个料号, 料号在E列
        top = self._fetch_top_with_pn(df, "AS", "E", 5, largest=True)
        suggestion = "重点处理呆滞库存，优化采购/退货/调拨"
        return {
            "success": True,
            "type": "呆滞风险",
            "items": top,
            "suggestion": suggestion
        }

    def _analyze_dead_stock_value(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 4：金额最大的呆滞物料

        数据来源：Analysis TOP short term savings 工作表 G 列（金金额度）
        处理逻辑：取 G 列最大 3 行。
        料号列：A 列。
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        # G列金额最大的3个, PN在A列
        top = self._fetch_top_with_pn(df, "G", "A", 3, largest=True)
        suggestion = "取消未交PO或转卖呆滞物料，内部调拨"
        return {
            "success": True,
            "type": "呆滞金额",
            "items": top,
            "suggestion": suggestion
        }

    def _analyze_transit_inconsistencies(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 5：运输天数不一致 (Days of transit inconsistencies)

        数据来源：Analysis TOP HIGH RISKS PARAMETERS 工作表 D 列（运输天数）
        处理逻辑：取 D 列最大 3 行，代表差异最大。
        料号列：A 列。
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        # D列天数最大三个, PN在A列
        top = self._fetch_top_with_pn(df, "D", "A", 3, largest=True)
        suggestion = "在SAP中统一供应商运输时间，标准化物流协议"
        return {
            "success": True,
            "type": "运输天数不一致",
            "items": top,
            "suggestion": suggestion
        }

    def _analyze_safety_time(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 6：安全时间过长 (Safety Time > 2 × Delivery frequency)

        数据来源：Analysis TOP HIGH RISKS PARAMETERS 工作表 N 列
        处理逻辑：取 N 列最大 1 行。
        料号列：H 列。
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        # N列最大一个, PN在H列
        items = self._fetch_top_with_pn(df, "N", "H", 1, largest=True)
        suggestion = "调整料号安全天数，优化库存策略"
        return {
            "success": True,
            "type": "安全时间过长",
            "items": items,
            "suggestion": suggestion
        }

    def _analyze_no_supplier(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 7：有需求但没有供应商 (Materials with requirements and no supplier)

        数据来源：Analysis TOP HIGH RISKS PARAMETERS 工作表 Q 列
        处理逻辑：筛选 Q 列非空料号，表示存在需求但未配置供应商。
        返回纯料号列表。
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        q_idx = self._col_letter_to_index("Q")
        pns = df.iloc[:, q_idx].dropna().tolist()
        suggestion = "检查替代料或增加供应商"
        return {
            "success": True,
            "type": "无供应商物料",
            "items": pns,
            "suggestion": suggestion
        }

    def _analyze_moq_impact(self, rule: pd.Series) -> Dict[str, Any]:
        """规则 8：MOQ 影响 (MOQ impact on potential end-of-life material)

        数据来源：Analysis TOP HIGH RISKS PARAMETERS 工作表 AO 列（MOQ 指标）
        处理逻辑：取 AO 列最大 3 行。
        料号列：AH 列。
        """
        sheet_name = rule["Sheet"]
        df = self.data_df[sheet_name]
        # AO列最大3, PN在AH列
        items = self._fetch_top_with_pn(df, "AO", "AH", 3, largest=True)
        suggestion = "与供应商谈判降低MOQ，调整采购策略"
        return {
            "success": True,
            "type": "MOQ影响",
            "items": items,
            "suggestion": suggestion
        }