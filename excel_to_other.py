import sys
import os
import json
import pandas as pd

# 添加当前目录下的lib目录到Python路径
lib_path = os.path.join(os.path.dirname(__file__), 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

def excel_to_csv(excel_file_path, output_csv_path="output.csv", sheet_name=0):
    """将Excel文件转换为CSV文件
    
    Args:
        excel_file_path: Excel文件路径
        output_csv_path: 输出CSV文件路径
        sheet_name: 要转换的工作表名称或索引，默认为第一个工作表
    
    Returns:
        bool: 转换是否成功
    """
    try:
        # 尝试打开Excel文件
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except PermissionError:
            print(f"错误：无法打开文件 {excel_file_path}，权限不足！")
            return False
        except Exception as e:
            print(f"错误：无法打开文件 {excel_file_path}，原因：{str(e)}")
            return False
        
        # 处理输出文件路径
        if os.path.exists(output_csv_path):
            base_name, ext = os.path.splitext(output_csv_path)
            counter = 1
            new_output_path = f"{base_name}_{counter}{ext}"
            while os.path.exists(new_output_path):
                counter += 1
                new_output_path = f"{base_name}_{counter}{ext}"
            output_csv_path = new_output_path
        
        # 写入CSV文件
        try:
            df.to_csv(output_csv_path, index=False, encoding='utf-8')
        except PermissionError:
            print(f"错误：无法写入文件 {output_csv_path}，权限不足！")
            return False
        except Exception as e:
            print(f"错误：写入CSV文件时发生错误，原因：{str(e)}")
            return False
        
        print(f"转换完成！结果已保存到 {output_csv_path}")
        return True
    except Exception as e:
        print(f"错误：处理文件时发生未知错误，原因：{str(e)}")
        return False

def excel_to_json(excel_file_path, output_json_path="output.json", sheet_name=0):
    """将Excel文件转换为JSON文件
    
    Args:
        excel_file_path: Excel文件路径
        output_json_path: 输出JSON文件路径
        sheet_name: 要转换的工作表名称或索引，默认为第一个工作表
    
    Returns:
        bool: 转换是否成功
    """
    try:
        # 尝试打开Excel文件
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except PermissionError:
            print(f"错误：无法打开文件 {excel_file_path}，权限不足！")
            return False
        except Exception as e:
            print(f"错误：无法打开文件 {excel_file_path}，原因：{str(e)}")
            return False
        
        # 处理输出文件路径
        if os.path.exists(output_json_path):
            base_name, ext = os.path.splitext(output_json_path)
            counter = 1
            new_output_path = f"{base_name}_{counter}{ext}"
            while os.path.exists(new_output_path):
                counter += 1
                new_output_path = f"{base_name}_{counter}{ext}"
            output_json_path = new_output_path
        
        # 将数据转换为与test_new.json相同的格式
        json_data = {}
        for index, row in df.iterrows():
            # 使用行索引或第一个列的值作为键
            key = str(index + 1)
            if len(df.columns) > 0:
                first_col_value = row[df.columns[0]]
                if isinstance(first_col_value, str):
                    key = first_col_value
                else:
                    key = str(first_col_value)
            
            # 创建行数据
            row_data = {}
            for col_name, value in row.items():
                # 检查值是否为NaN
                if pd.isna(value):
                    continue
                
                # 检查值是否为字典（可能包含函数定义）
                if isinstance(value, str) and value.startswith('='):
                    # 尝试解析Excel公式
                    try:
                        # 简单处理SUM、AVERAGE、MAX、MIN等函数
                        if value.startswith('=SUM('):
                            func_type = 'SUM'
                            # 这里可以添加更复杂的公式解析逻辑
                            # 目前只保存公式字符串
                            row_data[col_name] = value
                        elif value.startswith('=AVERAGE('):
                            func_type = 'AVERAGE'
                            row_data[col_name] = value
                        elif value.startswith('=MAX('):
                            func_type = 'MAX'
                            row_data[col_name] = value
                        elif value.startswith('=MIN('):
                            func_type = 'MIN'
                            row_data[col_name] = value
                        else:
                            # 其他公式直接保存
                            row_data[col_name] = value
                    except Exception as e:
                        # 解析失败，直接保存原始值
                        row_data[col_name] = value
                else:
                    # 保存普通值
                    row_data[col_name] = value
            
            json_data[key] = row_data
        
        # 写入JSON文件
        try:
            with open(output_json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=4)
        except PermissionError:
            print(f"错误：无法写入文件 {output_json_path}，权限不足！")
            return False
        except Exception as e:
            print(f"错误：写入JSON文件时发生错误，原因：{str(e)}")
            return False
        
        print(f"转换完成！结果已保存到 {output_json_path}")
        return True
    except Exception as e:
        print(f"错误：处理文件时发生未知错误，原因：{str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("使用方法：python excel_to_other.py <excel文件路径> <输出格式(csv/json)> [输出文件路径] [工作表名称/索引]")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    output_format = sys.argv[2].lower()
    output_path = None
    sheet_name = 0
    
    if not os.path.exists(excel_file):
        print(f"错误：文件 {excel_file} 不存在！")
        sys.exit(1)
    
    if len(sys.argv) >= 4:
        output_path = sys.argv[3]
    
    if len(sys.argv) >= 5:
        try:
            sheet_name = int(sys.argv[4])
        except ValueError:
            sheet_name = sys.argv[4]
    
    success = False
    if output_format == "csv":
        if output_path is None:
            output_path = "output.csv"
        success = excel_to_csv(excel_file, output_path, sheet_name)
    elif output_format == "json":
        if output_path is None:
            output_path = "output.json"
        success = excel_to_json(excel_file, output_path, sheet_name)
    else:
        print(f"错误：不支持的输出格式 {output_format}，仅支持 csv 和 json")
        sys.exit(1)
    
    if not success:
        sys.exit(1)