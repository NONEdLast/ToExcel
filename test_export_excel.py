import pandas as pd
import os
import io
import json

def parse_json_function(func_dict):
    """解析JSON格式的函数表示，转换为Excel函数格式"""
    if not isinstance(func_dict, dict) or 'type' not in func_dict or 'sub' not in func_dict:
        return str(func_dict)
    
    # 函数类型映射
    func_type = func_dict['type'].upper()
    
    # 解析参数
    params = []
    for param in func_dict['sub']:
        if not isinstance(param, dict) or 'f' not in param or 'r' not in param or 'c' not in param:
            params.append(str(param))
            continue
        
        # 参数类型：r表示相对引用，a表示绝对引用
        ref_type = param['f']
        row_offset = param['r']
        col_offset = param['c']
        
        # 计算列字母（Excel列名，如A、B、C...）
        col_letter = ''
        abs_col = abs(col_offset) if col_offset < 0 else col_offset + 3  # 假设col0是A列
        while abs_col > 0:
            abs_col -= 1
            col_letter = chr(ord('A') + abs_col % 26) + col_letter
            abs_col //= 26
        
        # 计算行号（Excel行号从1开始）
        row_num = abs(row_offset) + 1 if row_offset < 0 else row_offset + 4  # 假设row0是第1行
        
        # 构建引用格式
        if ref_type == 'a':
            # 绝对引用：$A$1
            cell_ref = f"${col_letter}${row_num}"
        else:
            # 相对引用：A1
            cell_ref = f"{col_letter}{row_num}"
        
        params.append(cell_ref)
    
    # 构建Excel函数
    excel_func = f"={func_type}({','.join(params)})"
    return excel_func

def convert_to_number_if_possible(value):
    """将字符串转换为数字类型（如果可能）"""
    if isinstance(value, str):
        value = value.strip()
        if value:
            # 如果是Excel函数（以=开头），不进行转换
            if value.startswith('='):
                return value
            try:
                return int(value)
            except ValueError:
                try:
                    return float(value)
                except ValueError:
                    return value
    return value

def test_export_excel():
    """测试Excel导出功能"""
    file_path = r"d:\AIAssist\toexcel\test_new.json"
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        # 解析JSON数据
        processed_data = []
        for key, item in json_data.items():
            processed_item = {}
            for sub_key, value in item.items():
                if isinstance(value, dict) and 'type' in value and 'sub' in value:
                    # 如果是函数字典，转换为Excel函数格式
                    processed_item[sub_key] = parse_json_function(value)
                else:
                    processed_item[sub_key] = str(value) if value is not None else None
            processed_data.append(processed_item)
        df = pd.DataFrame(processed_data)
        
        print("解析后的DataFrame:")
        print(df)
        print("\ncol3列的值:", list(df['col3']))
        
        # 创建临时Excel文件
        temp_excel_path = "test_export.xlsx"
        
        # 导出Excel时将纯数字字符串转换为数字类型
        df_to_export = df.copy()
        for col in df_to_export.columns:
            # 尝试将每列转换为数字类型
            try:
                df_to_export[col] = df_to_export[col].apply(convert_to_number_if_possible)
            except Exception:
                # 如果转换失败，保持原类型
                pass
        
        print("\n导出前的DataFrame:")
        print(df_to_export)
        print("\ncol3列的值:", list(df_to_export['col3']))
        
        # 导出Excel
        df_to_export.to_excel(temp_excel_path, index=False)
        
        print(f"\nExcel文件已导出到: {temp_excel_path}")
        print("导出成功！")
        
        return True
    except Exception as e:
        print(f"错误：导出Excel时发生错误，原因：{str(e)}")
        return False

# 运行测试
if __name__ == "__main__":
    test_export_excel()