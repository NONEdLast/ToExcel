import sys
import os

# 添加当前目录下的lib目录到Python路径
lib_path = os.path.join(os.path.dirname(__file__), 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import pandas as pd

def sort_strings_by_unicode(strings):
    """按Unicode编码对字符串进行排序
    
    Args:
        strings: 字符串列表
    
    Returns:
        按Unicode编码升序排序后的字符串列表
    """
    # 使用Python的默认排序，它会自动按Unicode编码排序
    return sorted(strings)

def txt_to_excel(txt_file_path, output_excel_path="output.xlsx", sort_by_unicode=False):
    try:
        # 尝试打开文件以检查权限
        with open(txt_file_path, 'r') as f:
            pass
    except PermissionError:
        print(f"错误：无法打开文件 {txt_file_path}，权限不足！")
        return False
    except Exception as e:
        print(f"错误：无法打开文件 {txt_file_path}，原因：{str(e)}")
        return False
    
    try:
        # 读取txt文件，自动检测分隔符和表头
        try:
            # 尝试将第一行作为表头读取
            df = pd.read_csv(txt_file_path, sep=None, engine='python', header=0)
            has_header = True
        except Exception as e:
            # 如果读取失败，尝试不使用表头读取
            df = pd.read_csv(txt_file_path, sep=None, engine='python', header=None)
            has_header = False
        
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
                        df[unicode_sort_col] = df[col].apply(lambda x: value_map.get(str(x), -1) if pd.notna(x) else -1)
        
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
        print("使用方法：python txt_to_excel.py <txt文件路径> [输出Excel路径] [--sort-by-unicode]")
        sys.exit(1)
    
    txt_file = sys.argv[1]
    if not os.path.exists(txt_file):
        print(f"错误：文件 {txt_file} 不存在！")
        sys.exit(1)
    
    output_path = "output.xlsx"
    if len(sys.argv) >= 3 and sys.argv[2] != "--sort-by-unicode":
        output_path = sys.argv[2]
    
    sort_by_unicode = False
    if "--sort-by-unicode" in sys.argv:
        sort_by_unicode = True
    
    success = txt_to_excel(txt_file, output_path, sort_by_unicode)
    if not success:
        sys.exit(1)