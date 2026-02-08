import sys
import os

# 添加当前目录下的lib目录到Python路径
lib_path = os.path.join(os.path.dirname(__file__), 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import pandas as pd

def txt_to_excel(txt_file_path, output_excel_path="output.xlsx"):
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
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                
                # 为每一列创建一个排序后的sheet
                for col_idx, col_name in enumerate(df.columns):
                    # 按当前列从大到小排序，保持表头
                    sorted_df = df.sort_values(by=col_name, ascending=False)
                    # 写入到新的sheet，从Sheet2开始
                    sheet_name = f'Sheet{col_idx + 2}'
                    sorted_df.to_excel(writer, sheet_name=sheet_name, index=False)
        except PermissionError:
            print(f"错误：无法写入文件 {output_excel_path}，权限不足！")
            return False
        except Exception as e:
            print(f"错误：写入Excel文件时发生错误，原因：{str(e)}")
            return False
        
        print(f"转换完成！结果已保存到 {output_excel_path}")
        print(f"原始数据保存在 Sheet1")
        for col_idx, col_name in enumerate(df.columns):
            print(f"按第 {col_idx + 1} 列降序排序后的数据保存在 Sheet{col_idx + 2}")
        return True
    except Exception as e:
        print(f"错误：处理文件时发生未知错误，原因：{str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("使用方法：python txt_to_excel.py <txt文件路径>")
        sys.exit(1)
    
    txt_file = sys.argv[1]
    if not os.path.exists(txt_file):
        print(f"错误：文件 {txt_file} 不存在！")
        sys.exit(1)
    
    success = txt_to_excel(txt_file)
    if not success:
        sys.exit(1)