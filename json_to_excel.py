import sys
import os

# 添加当前目录下的lib目录到Python路径
lib_path = os.path.join(os.path.dirname(__file__), 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import pandas as pd
import json
import numpy as np

# 函数计算辅助函数
def calculate_function_value(func_type, cell_references):
    """构建Excel函数公式，不进行实际计算
    
    Args:
        func_type: 函数类型（SUM, AVERAGE, MAX, MIN等）
        cell_references: 单元格引用列表
    
    Returns:
        Excel公式字符串（如=SUM(A1,B2)）
    """
    if not cell_references:
        return None
    
    # 将单元格引用用逗号连接
    references_str = ",".join(cell_references)
    
    # 返回Excel公式形式
    return f"={func_type.upper()}({references_str})"

def convert_text_to_number(text):
    """将文本形式的数字转换为数字类型"""
    if not isinstance(text, str):
        return text
    
    # 尝试转换为整数
    try:
        return int(text)
    except ValueError:
        pass
    
    # 尝试转换为浮点数
    try:
        return float(text)
    except ValueError:
        pass
    
    # 无法转换，返回原文本
    return text

def convert_to_excel_cell(row, col):
    """将行号和列号转换为Excel单元格引用
    
    Args:
        row: 行号（从0开始）
        col: 列号（从0开始）
    
    Returns:
        Excel单元格引用字符串（如A1, B2等）
    """
    # 将列号转换为字母
    column_letters = ""
    col_num = col
    
    while col_num >= 0:
        # Excel列从A开始（对应0），所以需要加1
        col_num += 1
        # 计算当前字母
        column_letters = chr(ord('A') + (col_num % 26) - 1) + column_letters
        # 计算下一位
        col_num = col_num // 26 - 1
        
        if col_num < 0:
            break
    
    # 行号从1开始
    row_num = row + 2
    
    return f"{column_letters}{row_num}"

def extract_cell_value(cell_data):
    """从单元格数据中提取值"""
    if isinstance(cell_data, dict):
        if "value" in cell_data:
            return cell_data["value"]
        else:
            return cell_data  # 可能是函数定义
    return cell_data

def sort_strings_by_unicode(strings):
    """按Unicode编码对字符串进行排序
    
    Args:
        strings: 字符串列表
    
    Returns:
        按Unicode编码升序排序后的字符串列表
    """
    # 使用Python的默认排序，它会自动按Unicode编码排序
    return sorted(strings)

def json_to_excel(json_file_path, output_excel_path="output.json.xlsx", sort_by_unicode=False):
    try:
        # 尝试打开并读取JSON文件
        try:
            with open(json_file_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
        except PermissionError:
            print(f"错误：无法打开文件 {json_file_path}，权限不足！")
            return False
        except json.JSONDecodeError as e:
            print(f"错误：JSON格式错误，无法解析文件 {json_file_path}，原因：{str(e)}")
            return False
        except Exception as e:
            print(f"错误：无法打开或读取文件 {json_file_path}，原因：{str(e)}")
            return False
        
        # 将JSON数据转换为DataFrame
        try:
            # 处理不同的JSON结构
            if isinstance(json_data, list):
                # 列表结构：[{"key1": value1, "key2": value2}, ...]
                data_list = json_data
            elif isinstance(json_data, dict):
                # 检查是否为嵌套字典结构
                if all(isinstance(v, dict) for v in json_data.values()):
                    # 嵌套字典结构：{"row1": {"col1": value1, "col2": value2}, ...}
                    data_list = list(json_data.values())
                else:
                    # 字典结构：{"column1": [value1, value2, ...], "column2": [...], ...}
                    data_list = json_data
            else:
                print("错误：JSON数据结构不支持，仅支持列表或字典格式")
                return False
            
            # 创建DataFrame
            df = pd.DataFrame(data_list)
        except Exception as e:
            print(f"错误：将JSON数据转换为表格时发生错误，原因：{str(e)}")
            return False
        
        # 检查DataFrame是否为空
        if df.empty:
            print("错误：JSON数据为空或格式不符合要求")
            return False
        
        # 将文本形式的数字转换为数字类型
        try:
            # 遍历所有列
            for col in df.columns:
                # 跳过已经是数字类型的列
                if df[col].dtype in ['int64', 'float64']:
                    continue
                    
                # 仅处理对象类型的列
                if df[col].dtype == 'object':
                    # 创建一个新的数组来存储转换后的值
                    converted_values = []
                    has_dicts = False
                    
                    # 遍历列中的每个值
                    for val in df[col]:
                        if isinstance(val, dict):
                            # 保留字典值（函数定义）
                            converted_values.append(val)
                            has_dicts = True
                        else:
                            # 尝试转换为数字
                            try:
                                # 尝试转换为整数
                                converted_val = int(val)
                            except ValueError:
                                try:
                                    # 尝试转换为浮点数
                                    converted_val = float(val)
                                except ValueError:
                                    # 无法转换，保留原始值
                                    converted_val = val
                            converted_values.append(converted_val)
                    
                    # 更新列的值
                    df[col] = converted_values
        except Exception as e:
            print(f"错误：转换文本数字时发生错误，原因：{str(e)}")
            return False
        
        # 处理函数单元格
        try:
            # 复制DataFrame进行处理
            calculated_df = df.copy()
            
            # 遍历所有单元格
            for row_idx in range(len(calculated_df)):
                for col_idx, col_name in enumerate(calculated_df.columns):
                    cell_data = calculated_df.iloc[row_idx, col_idx]
                    
                    # 检查是否为函数定义
                    if isinstance(cell_data, dict) and "type" in cell_data and "sub" in cell_data:
                        func_type = cell_data["type"]
                        func_subs = cell_data["sub"]
                        
                        # 收集单元格引用
                        cell_references = []
                        for sub in func_subs:
                            if "r" in sub and "c" in sub:
                                # 计算目标单元格的位置
                                target_row = row_idx + sub["r"]
                                target_col = col_idx + sub["c"]
                                
                                # 检查位置是否有效
                                if 0 <= target_row < len(calculated_df) and 0 <= target_col < len(calculated_df.columns):
                                    # 将行号和列号转换为Excel单元格引用
                                    cell_ref = convert_to_excel_cell(target_row, target_col)
                                    cell_references.append(cell_ref)
                        
                        # 构建Excel函数公式
                        if cell_references:
                            result = calculate_function_value(func_type, cell_references)
                            if result is not None:
                                # 将函数公式存入DataFrame
                                calculated_df.iloc[row_idx, col_idx] = result
                            else:
                                calculated_df.iloc[row_idx, col_idx] = f"=ERROR: {func_type}"
                        else:
                            calculated_df.iloc[row_idx, col_idx] = "=ERROR: No cell references"
        except Exception as e:
            print(f"错误：处理函数单元格时发生错误，原因：{str(e)}")
            return False
        
        # 使用处理后的DataFrame
        df = calculated_df
        
        # 如果启用了Unicode排序，对所有字符串列进行排序
        if sort_by_unicode:
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 提取非空字符串值
                    string_values = df[col].dropna().astype(str)
                    if not string_values.empty:
                        # 按Unicode编码排序
                        sorted_values = sort_strings_by_unicode(string_values)
                        # 创建映射字典
                        value_map = {v: i for i, v in enumerate(sorted_values)}
                        # 创建新列名
                        unicode_sort_col = f"{col}_unicode_sort"
                        # 添加排序后的列
                        df[unicode_sort_col] = df[col].apply(lambda x: value_map.get(str(x), -1) if pd.notna(x) and not isinstance(x, dict) else x)
        
        # 处理输出文件路径
        if os.path.exists(output_excel_path):
            base_name, ext = os.path.splitext(output_excel_path)
            counter = 1
            new_output_path = f"{base_name}_{counter}{ext}"
            while os.path.exists(new_output_path):
                counter += 1
                new_output_path = f"{base_name}_{counter}{ext}"
            output_excel_path = new_output_path
        
        # 创建Excel写入器
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                # 写入原始表格到sheet1
                df.to_excel(writer, sheet_name='原始数据', index=False)
                
                # 为每一列创建一个排序后的sheet
                for col_idx, col_name in enumerate(df.columns):
                    # 按当前列从大到小排序，保持表头
                    sorted_df = df.sort_values(by=col_name, ascending=False)
                    # 写入到新的sheet，使用更具描述性的名称
                    sheet_name = f'按{col_name}降序'
                    sorted_df.to_excel(writer, sheet_name=sheet_name, index=False)
        except PermissionError:
            print(f"错误：无法写入文件 {output_excel_path}，权限不足！")
            return False
        except Exception as e:
            print(f"错误：写入Excel文件时发生错误，原因：{str(e)}")
            return False
        
        print(f"转换完成！结果已保存到 {output_excel_path}")
        print(f"原始数据保存在 '原始数据'")
        for col_idx, col_name in enumerate(df.columns):
            print(f"按 {col_name} 列降序排序后的数据保存在 '按{col_name}降序'")
        return True
    except Exception as e:
        print(f"错误：处理文件时发生未知错误，原因：{str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 2 or len(sys.argv) > 4:
        print("使用方法：python json_to_excel.py <json文件路径> [输出Excel路径] [--sort-by-unicode]")
        sys.exit(1)
    
    json_file = sys.argv[1]
    if not os.path.exists(json_file):
        print(f"错误：文件 {json_file} 不存在！")
        sys.exit(1)
    
    output_path = "output.json.xlsx"
    if len(sys.argv) >= 3 and sys.argv[2] != "--sort-by-unicode":
        output_path = sys.argv[2]
    
    sort_by_unicode = False
    if "--sort-by-unicode" in sys.argv:
        sort_by_unicode = True
    
    success = json_to_excel(json_file, output_path, sort_by_unicode)
    if not success:
        sys.exit(1)