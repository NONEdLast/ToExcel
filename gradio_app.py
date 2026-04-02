import sys
import os
import uuid

# Add current directory to Python path so it can find project modules
current_dir = os.path.dirname(__file__)
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 首先安装依赖（如果缺少）
try:
    import gradio
    import openpyxl
    import pandas as pd
except ImportError:
    print("正在安装依赖...")
    # 使用当前Python解释器来安装依赖，确保版本兼容
    python_path = sys.executable
    os.system(f"\"{python_path}\" -m pip install gradio openpyxl pandas")
    # 重新导入
    import gradio
    import openpyxl
    import pandas as pd

# 添加当前目录下的lib目录到Python路径
lib_path = os.path.join(os.path.dirname(__file__), 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import json
import io
import gradio as gr

# 导入对应的处理脚本
import txt_to_excel
import json_to_excel
import excel_to_other
import sqlite_to_excel

# 设置临时文件目录
temp_dir = "temp_files"
os.makedirs(temp_dir, exist_ok=True)

def gradio_interface(file, detect_header, sort_by_unicode):
    """Gradio接口函数"""
    if file is None:
        return None, "请先上传文件", None
    
    try:
        file_path = file.name
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # 创建临时Excel文件
        temp_excel_path = os.path.join(temp_dir, f"result_{uuid.uuid4()}.xlsx")
        
        # 根据文件类型选择对应的处理脚本
        if file_ext == '.txt':
            print(f"处理TXT文件：{file_path}")
            success = txt_to_excel.txt_to_excel(file_path, temp_excel_path, sort_by_unicode, detect_header)
            if not success:
                return None, "错误：处理TXT文件时发生错误", None
        elif file_ext == '.json':
            print(f"处理JSON文件：{file_path}")
            success = json_to_excel.json_to_excel(file_path, temp_excel_path, sort_by_unicode)
            if not success:
                return None, "错误：处理JSON文件时发生错误", None
        else:
            return None, f"错误：不支持的文件格式 {file_ext}，当前仅支持 .txt 和 .json", None
        
        # 生成预览表格
        preview_html = "<h2>转换结果预览</h2>"
        with pd.ExcelFile(temp_excel_path) as excel_file:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                preview_html += f"<h3>{sheet_name}</h3>"
                preview_html += df.head().to_html(classes='dataframe', index=False)
        
        # 生成转换信息
        with pd.ExcelFile(temp_excel_path) as excel_file:
            num_sheets = len(excel_file.sheet_names)
            info = f"转换完成！包含 {num_sheets} 个工作表："
            for i, sheet_name in enumerate(excel_file.sheet_names, 1):
                info += f"\n{i}. {sheet_name}"
        
        return preview_html, info, temp_excel_path
    except Exception as e:
        return None, f"错误：处理文件时发生错误，原因：{str(e)}", None

def clear_cache():
    """清理缓存文件"""
    try:
        # 获取temp_files目录中的所有文件
        temp_files = os.listdir(temp_dir)
        
        if not temp_files:
            return "缓存目录为空，无需清理"
        
        # 删除所有临时文件
        for file in temp_files:
            file_path = os.path.join(temp_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        
        return f"成功清理 {len(temp_files)} 个缓存文件"
    except Exception as e:
        return f"清理缓存时发生错误，原因：{str(e)}"

def search_interface(file, sheet_name, query):
    """查找功能的Gradio接口函数"""
    if file is None:
        return "", "", "请先上传文件"
    
    try:
        file_path = file.name
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # 处理工作表名称/索引（仅Excel文件需要）
        if file_ext in ['.xlsx', '.xls']:
            try:
                # 尝试将sheet_name转换为整数（如果是数字字符串）
                sheet_name = int(sheet_name)
            except ValueError:
                # 如果转换失败，保留为字符串
                pass
            except Exception as e:
                return "", "", f"错误：解析工作表名称/索引时发生错误，原因：{str(e)}"
        
        # 根据文件类型读取内容
        df = None
        if file_ext == '.txt':
            # 读取TXT文件
            try:
                df = pd.read_csv(file_path, sep=None, engine='python', header=0)
            except Exception as e:
                try:
                    df = pd.read_csv(file_path, sep=None, engine='python', header=None)
                except Exception as e2:
                    return "", "", f"错误：读取TXT文件时发生错误，原因：{str(e2)}"
        elif file_ext == '.json':
            # 读取JSON文件
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                
                # 检查JSON格式
                if isinstance(json_data, list):
                    # 列表格式：[{}, {}, ...]
                    df = pd.DataFrame(json_data)
                elif isinstance(json_data, dict):
                    # 字典格式：{"column1": [values], "column2": [values], ...}
                    df = pd.DataFrame(json_data)
                else:
                    # 不支持的格式
                    return "", "", "错误：不支持的JSON格式，仅支持列表格式或字典格式"
            except Exception as e:
                return "", "", f"错误：读取JSON文件时发生错误，原因：{str(e)}"
        elif file_ext in ['.xlsx', '.xls']:
            # 读取Excel文件
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            except Exception as e:
                return "", "", f"错误：读取Excel文件时发生错误，原因：{str(e)}"
        else:
            return "", "", f"错误：不支持的文件格式 {file_ext}，仅支持 .txt、.json、.xlsx 和 .xls"
        
        # 生成完整表格HTML
        full_table_html = f"<h2>完整表格</h2>"
        full_table_html += df.to_html(classes='dataframe', index=False)
        
        # 生成查找结果HTML
        search_result_html = "<h2>查找结果</h2>"
        if query:
            # 执行查找
            try:
                # 将查询转换为字符串
                query_str = str(query).lower()
                
                # 创建一个布尔掩码，标记包含查询内容的行
                # 使用apply和map代替applymap（兼容pandas 2.0+）
                mask = df.apply(lambda col: col.map(lambda x: query_str in str(x).lower()))
                
                # 获取所有匹配的行
                matching_rows = df[mask.any(axis=1)]
                
                if matching_rows.empty:
                    search_result_html += "<p>未找到匹配的结果</p>"
                else:
                    search_result_html += matching_rows.to_html(classes='dataframe', index=False)
                
                info = f"查找完成！共找到 {len(matching_rows)} 行匹配的结果"
            except Exception as e:
                search_result_html += f"<p>查找时发生错误：{str(e)}</p>"
                info = f"错误：查找时发生错误，原因：{str(e)}"
        else:
            search_result_html += "<p>请输入要查找的内容</p>"
            info = "已加载文件，显示完整表格"
        
        return full_table_html, search_result_html, info
    except Exception as e:
        return "", "", f"错误：处理文件时发生错误，原因：{str(e)}"

def excel_to_other_interface(excel_file, output_format, sheet_name):
    """Excel转CSV/JSON的Gradio接口函数"""
    if excel_file is None:
        return None, "请先上传Excel文件", None
    
    try:
        excel_path = excel_file.name
        file_ext = os.path.splitext(excel_path)[1].lower()
        
        if file_ext != '.xlsx' and file_ext != '.xls':
            return None, f"错误：不支持的文件格式 {file_ext}，仅支持 .xlsx 和 .xls", None
        
        # 处理工作表名称/索引
        try:
            # 尝试将sheet_name转换为整数（如果是数字字符串）
            sheet_name = int(sheet_name)
        except ValueError:
            # 如果转换失败，保留为字符串
            pass
        except Exception as e:
            return None, f"错误：解析工作表名称/索引时发生错误，原因：{str(e)}", None
        
        # 创建临时输出文件
        if output_format == "CSV":
            temp_output_path = os.path.join(temp_dir, f"result_{uuid.uuid4()}.csv")
            success = excel_to_other.excel_to_csv(excel_path, temp_output_path, sheet_name)
        else:  # JSON
            temp_output_path = os.path.join(temp_dir, f"result_{uuid.uuid4()}.json")
            success = excel_to_other.excel_to_json(excel_path, temp_output_path, sheet_name)
        
        if not success:
            return None, f"错误：转换Excel文件时发生错误", None
        
        # 生成预览
        preview_html = f"<h2>{output_format}结果预览</h2>"
        if output_format == "CSV":
            # 读取CSV文件并生成预览
            df = pd.read_csv(temp_output_path)
            preview_html += df.head().to_html(classes='dataframe', index=False)
        else:  # JSON
            # 读取JSON文件并生成预览
            with open(temp_output_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            # 将JSON数据转换为DataFrame以生成预览
            df = pd.DataFrame.from_dict(json_data, orient='index')
            preview_html += df.head().to_html(classes='dataframe')
        
        # 生成转换信息
        info = f"转换完成！已将Excel文件转换为{output_format}格式。"
        if sheet_name != 0:
            info += f" 使用的工作表：{sheet_name}"
        
        return preview_html, info, temp_output_path
    except Exception as e:
        return None, f"错误：处理文件时发生错误，原因：{str(e)}", None

# 全局变量用于存储当前加载的表格数据
current_df = None

# 行和列操作的后端函数
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

def load_data(file, sheet_name):
    """加载文件数据并返回DataFrame"""
    global current_df
    
    if file is None:
        return pd.DataFrame(), "请先上传文件"
    
    try:
        file_path = file.name
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # 创建内存Excel对象
        excel_bytes_io = io.BytesIO()
        
        # 根据文件类型读取并转换为内存Excel
        if file_ext == '.txt' or file_ext == '.csv':
            # 读取TXT/CSV文件，将所有内容作为字符串读取
            try:
                df = pd.read_csv(file_path, sep=None, engine='python', header=0, dtype=str)
            except Exception as e:
                try:
                    df = pd.read_csv(file_path, sep=None, engine='python', header=None, dtype=str)
                except Exception as e2:
                    return pd.DataFrame(), f"错误：读取TXT/CSV文件时发生错误，原因：{str(e2)}"
            
            # 将DataFrame写入内存Excel
            with pd.ExcelWriter(excel_bytes_io, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # 重置指针到开头
            excel_bytes_io.seek(0)
            
            # 从内存Excel读取第一个Sheet
            current_df = pd.read_excel(excel_bytes_io, sheet_name=0, dtype=str, engine='openpyxl')
            
        elif file_ext == '.json':
            # 读取JSON文件，确保数字类字符串保留为字符串
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                
                # 检查JSON格式
                if isinstance(json_data, list):
                    # 列表格式：[{}, {}, ...]，将所有值转换为字符串
                    processed_data = []
                    for item in json_data:
                        processed_item = {}
                        for key, value in item.items():
                            processed_item[key] = str(value) if value is not None else None
                        processed_data.append(processed_item)
                    df = pd.DataFrame(processed_data)
                elif isinstance(json_data, dict):
                    # 检查是否是字典的字典格式
                    if all(isinstance(v, dict) for v in json_data.values()):
                        # 字典的字典格式：{"key1": {}, "key2": {}, ...}
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
                    else:
                        # 字典格式：{"column1": [values], "column2": [values], ...}，将所有值转换为字符串
                        processed_dict = {}
                        for key, values in json_data.items():
                            processed_dict[key] = [str(v) if v is not None else None for v in values]
                        df = pd.DataFrame(processed_dict)
                else:
                    # 不支持的格式
                    return pd.DataFrame(), "错误：不支持的JSON格式，仅支持列表格式或字典格式"
            except Exception as e:
                return pd.DataFrame(), f"错误：读取JSON文件时发生错误，原因：{str(e)}"
            
            # 将DataFrame写入内存Excel
            with pd.ExcelWriter(excel_bytes_io, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # 重置指针到开头
            excel_bytes_io.seek(0)
            
            # 从内存Excel读取第一个Sheet
            current_df = pd.read_excel(excel_bytes_io, sheet_name=0, dtype=str, engine='openpyxl')
            
        elif file_ext in ['.xlsx', '.xls']:
            # 处理工作表名称/索引
            try:
                # 尝试将sheet_name转换为整数（如果是数字字符串）
                sheet_name = int(sheet_name)
            except ValueError:
                # 如果转换失败，保留为字符串
                pass
            
            # 直接读取Excel文件，保持字符串类型
            current_df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, engine='openpyxl')
            
        else:
            return pd.DataFrame(), f"错误：不支持的文件格式 {file_ext}，仅支持 .txt、.json、.xlsx 和 .xls"
        
        info = f"文件加载成功！共 {len(current_df)} 行，{len(current_df.columns)} 列"
        return current_df, info
    except Exception as e:
        return pd.DataFrame(), f"错误：处理文件时发生错误，原因：{str(e)}"

def auto_convert_numbers(value):
    """自动识别并转换纯数字字符串为数字类型"""
    if value is None or value.strip() == "":
        return None
    try:
        # 尝试转换为整数
        return int(value)
    except ValueError:
        try:
            # 尝试转换为浮点数
            return float(value)
        except ValueError:
            # 保留为字符串
            return value

def parse_add_column_content(content):
    """解析添加列的输入内容，将其分为列名和列内容两部分"""
    if content is None or content.strip() == "":
        return "", ""
    
    # 按逗号分割内容
    parts = [part.strip() for part in content.split(",")]
    
    # 第一部分作为列名
    col_name = parts[0]
    
    # 剩余部分作为列内容，用逗号连接
    column_content = ",".join(parts[1:]) if len(parts) > 1 else ""
    
    return col_name, column_content

def add_row(row_content=""):
    """添加一行到表格，支持自定义行内容"""
    global current_df
    
    if current_df is None:
        return pd.DataFrame(), "请先加载文件"
    
    try:
        # 创建一个新行，默认所有值为空
        new_row = {col: None for col in current_df.columns}
        
        # 如果用户提供了行内容，解析并填充
        if row_content is not None and row_content.strip() != "":
            # 按逗号分割内容
            values = [val.strip() for val in row_content.split(",")]
            
            # 填充到新行中（只填充有效列）
            for i, col in enumerate(current_df.columns):
                if i < len(values):
                    # 保留原始字符串
                    new_row[col] = values[i]
        
        # 添加新行
        current_df = pd.concat([current_df, pd.DataFrame([new_row])], ignore_index=True)
        
        info = f"行添加成功！当前共 {len(current_df)} 行"
        return current_df, info
    except Exception as e:
        return current_df, f"错误：添加行时发生错误，原因：{str(e)}"

def delete_row(row_index):
    """删除指定行（索引从1开始）"""
    global current_df
    
    if current_df is None:
        return pd.DataFrame(), "请先加载文件"
    
    try:
        # 检查行索引是否有效
        if row_index is None or row_index.strip() == "":
            return current_df, "请输入有效的行索引"
        
        # 转换为整数
        row_idx = int(row_index)
        
        # 将1-based索引转换为0-based索引
        row_idx -= 1
        
        # 检查索引范围
        if row_idx < 0 or row_idx >= len(current_df):
            return current_df, f"行索引超出范围！表格共有 {len(current_df)} 行（索引范围：1-{len(current_df)}）"
        
        # 删除行
        current_df = current_df.drop(index=row_idx).reset_index(drop=True)
        
        info = f"行删除成功！当前共 {len(current_df)} 行"
        return current_df, info
    except ValueError:
        return current_df, "请输入有效的数字行索引"
    except Exception as e:
        return current_df, f"错误：删除行时发生错误，原因：{str(e)}"

def add_column(col_name, column_content=""):
    """添加一列到表格，支持自定义列内容"""
    global current_df
    
    if current_df is None:
        return pd.DataFrame(), "请先加载文件"
    
    try:
        # 检查列名是否为空
        if col_name is None or col_name.strip() == "":
            return current_df, "请输入有效的列名"
        
        # 检查列名是否已存在
        if col_name in current_df.columns:
            return current_df, f"列名 '{col_name}' 已存在！"
        
        # 解析列内容（如果提供）
        values = []
        if column_content is not None and column_content.strip() != "":
            # 按逗号分割内容
            values = [val.strip() for val in column_content.split(",")]
        
        # 添加新列
        if values:
            # 如果提供了值，填充这些值，其余为空
            new_col = [None] * len(current_df)
            for i, val in enumerate(values):
                if i < len(new_col):
                    new_col[i] = val
            current_df[col_name] = new_col
        else:
            # 如果没有提供值，所有值为空
            current_df[col_name] = None
        
        info = f"列添加成功！当前共 {len(current_df.columns)} 列"
        return current_df, info
    except Exception as e:
        return current_df, f"错误：添加列时发生错误，原因：{str(e)}"

def delete_column(col_input):
    """删除指定列（支持列名或列索引，索引从1开始）"""
    global current_df
    
    if current_df is None:
        return pd.DataFrame(), "请先加载文件"
    
    try:
        # 检查输入是否为空
        if col_input is None or col_input.strip() == "":
            return current_df, "请输入有效的列名或列索引"
        
        col_to_delete = None
        
        # 尝试将输入转换为整数（列索引）
        try:
            col_idx = int(col_input.strip())
            
            # 将1-based索引转换为0-based索引
            col_idx -= 1
            
            # 检查索引范围
            if col_idx < 0 or col_idx >= len(current_df.columns):
                return current_df, f"列索引超出范围！表格共有 {len(current_df.columns)} 列（索引范围：1-{len(current_df.columns)}）"
            
            # 获取对应的列名
            col_to_delete = current_df.columns[col_idx]
        except ValueError:
            # 如果转换失败，将输入视为列名
            col_name = col_input.strip()
            
            # 检查列名是否存在
            if col_name not in current_df.columns:
                return current_df, f"列名 '{col_name}' 不存在！"
            
            col_to_delete = col_name
        
        # 删除列
        current_df = current_df.drop(columns=[col_to_delete])
        
        info = f"列删除成功！当前共 {len(current_df.columns)} 列"
        return current_df, info
    except Exception as e:
        return current_df, f"错误：删除列时发生错误，原因：{str(e)}"

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

def export_table(output_format):
    """导出当前表格"""
    global current_df
    
    if current_df is None:
        return None, "请先加载文件"
    
    try:
        # 创建临时输出文件
        if output_format == "Excel":
            temp_output_path = os.path.join(temp_dir, f"result_{uuid.uuid4()}.xlsx")
            
            # 导出Excel时将纯数字字符串转换为数字类型
            df_to_export = current_df.copy()
            for col in df_to_export.columns:
                # 尝试将每列转换为数字类型
                try:
                    df_to_export[col] = df_to_export[col].apply(convert_to_number_if_possible)
                except Exception:
                    # 如果转换失败，保持原类型
                    pass
            
            df_to_export.to_excel(temp_output_path, index=False)
        elif output_format == "CSV":
            temp_output_path = os.path.join(temp_dir, f"result_{uuid.uuid4()}.csv")
            current_df.to_csv(temp_output_path, index=False)
        elif output_format == "JSON":
            temp_output_path = os.path.join(temp_dir, f"result_{uuid.uuid4()}.json")
            current_df.to_json(temp_output_path, orient='records', force_ascii=False)
        else:
            return None, "不支持的导出格式"
        
        info = f"表格导出成功！格式：{output_format}"
        return temp_output_path, info
    except Exception as e:
        return None, f"错误：导出表格时发生错误，原因：{str(e)}"


def sqlite_to_excel_interface(db_file, tables, sort_by_unicode):
    """SQLite转Excel的Gradio接口函数"""
    if db_file is None:
        return None, "请先上传SQLite数据库文件"
    
    try:
        db_path = db_file.name
        
        # 创建临时Excel文件
        temp_excel_path = os.path.join(temp_dir, f"sqlite_to_excel_{uuid.uuid4()}.xlsx")
        
        # 处理表格列表
        tables_to_convert = None
        if tables is not None and tables.strip() != "":
            tables_to_convert = [table.strip() for table in tables.split(",")]
        
        # 执行转换
        success = sqlite_to_excel.sqlite_to_excel(db_path, temp_excel_path, tables_to_convert, sort_by_unicode)
        
        if success:
            # 生成转换信息
            info = f"SQLite数据库转换为Excel成功！"
            return temp_excel_path, info
        else:
            return None, "错误：SQLite转Excel失败"
            
    except Exception as e:
        return None, f"错误：处理SQLite文件时发生错误，原因：{str(e)}"


def excel_to_sqlite_interface(excel_file, table_name, sheet_name, calculate_functions):
    """Excel转SQLite的Gradio接口函数"""
    if excel_file is None:
        return None, "请先上传Excel文件"
    
    try:
        excel_path = excel_file.name
        
        # 创建临时SQLite文件
        temp_db_path = os.path.join(temp_dir, f"excel_to_sqlite_{uuid.uuid4()}.db")
        
        # 处理工作表名称/索引
        try:
            sheet_name = int(sheet_name)
        except ValueError:
            # 如果转换失败，保留为字符串
            pass
        
        # 执行转换
        success = sqlite_to_excel.excel_to_sqlite(
            excel_path, 
            temp_db_path, 
            table_name if table_name.strip() != "" else None, 
            sheet_name, 
            calculate_functions
        )
        
        if success:
            # 生成转换信息
            info = f"Excel转换为SQLite成功！"
            return temp_db_path, info
        else:
            return None, "错误：Excel转SQLite失败"
            
    except Exception as e:
        return None, f"错误：处理Excel文件时发生错误，原因：{str(e)}"

# 创建Gradio界面
with gr.Blocks(title="文档转换工具") as app:
    gr.Markdown("# 文档转换工具")
    gr.Markdown("支持TXT/JSON与Excel文件之间的相互转换")
    
    # 创建选项卡
    with gr.Tabs():
        # 第一个选项卡：TXT/JSON转Excel
        with gr.TabItem("TXT/JSON转Excel"):
            gr.Markdown("支持上传TXT和JSON文件，自动转换为Excel并提供预览")
            
            with gr.Row():
                with gr.Column(scale=1):
                    file_input = gr.File(label="上传文件", file_types=[".txt", ".json"])
                    detect_header_checkbox = gr.Checkbox(label="检测表头", value=True)
                    sort_checkbox = gr.Checkbox(label="按Unicode编码对字符串排序", value=False)
                    convert_btn = gr.Button("转换", variant="primary")
                    clear_btn = gr.Button("清理缓存", variant="secondary")
                    info_output = gr.Textbox(label="转换信息", lines=5, interactive=False)
                    cache_info = gr.Textbox(label="缓存状态", lines=2, interactive=False)
                    excel_output = gr.File(label="下载Excel文件")
                
                with gr.Column(scale=2):
                    preview_output = gr.HTML(label="表格预览")
            
            # 设置转换按钮的点击事件
            convert_btn.click(
                fn=gradio_interface,
                inputs=[file_input, detect_header_checkbox, sort_checkbox],
                outputs=[preview_output, info_output, excel_output]
            )
            
            # 设置清理按钮的点击事件
            clear_btn.click(
                fn=clear_cache,
                outputs=cache_info
            )
            
            # 也支持文件上传后自动转换
            file_input.change(
                fn=gradio_interface,
                inputs=[file_input, detect_header_checkbox, sort_checkbox],
                outputs=[preview_output, info_output, excel_output]
            )
            
            # 添加使用说明
            gr.Markdown("## 使用说明")
            gr.Markdown("""
            1. 点击"上传文件"按钮，选择要转换的TXT或JSON文件
            2. 系统会自动开始转换，或点击"转换"按钮手动开始
            3. 在右侧可以预览转换后的表格内容
            4. 可以下载完整的Excel文件
            
            **支持的文件格式：**
            - TXT：支持自动检测分隔符，自动识别表头
            - JSON：支持两种格式：
              - 列表格式：`[{"key1": value1, "key2": value2}, ...]`
              - 字典格式：`{"column1": [value1, value2, ...], "column2": [...], ...}`
            
            **转换规则：**
            - 原始数据：原始数据
            - 按列名降序：按各列降序排序后的数据（每个列名对应一个工作表）
            
            **检测表头功能：**
            - 勾选"检测表头"选项（默认开启）后，系统会尝试将TXT文件的第一行作为表头
            - 取消勾选后，系统会将所有行作为数据读取，不使用表头
            
            **Unicode排序功能：**
            - 勾选"按Unicode编码对字符串排序"选项后，系统会为每个字符串列添加一个新列
            - 新列名格式为"原列名_unicode_sort"，包含按Unicode编码升序排序的序号
            - 支持识别和排序各种Unicode字符，包括中文、英文、数字和特殊字符
            """)
        
        # 第二个选项卡：Excel转CSV/JSON
        with gr.TabItem("Excel转CSV/JSON"):
            gr.Markdown("支持上传Excel文件，转换为CSV或JSON格式")
            
            with gr.Row():
                with gr.Column(scale=1):
                    excel_input = gr.File(label="上传Excel文件", file_types=[".xlsx", ".xls"])
                    output_format_radio = gr.Radio(
                        label="输出格式", 
                        choices=["CSV", "JSON"], 
                        value="CSV"
                    )
                    sheet_name_input = gr.Textbox(
                        label="工作表名称或索引（可选，默认使用第一个工作表）", 
                        value="0", 
                        lines=1
                    )
                    convert_excel_btn = gr.Button("转换", variant="primary")
                    clear_excel_btn = gr.Button("清理缓存", variant="secondary")
                    excel_info_output = gr.Textbox(label="转换信息", lines=5, interactive=False)
                    excel_cache_info = gr.Textbox(label="缓存状态", lines=2, interactive=False)
                    other_output = gr.File(label="下载转换后的文件")
                
                with gr.Column(scale=2):
                    excel_preview_output = gr.HTML(label="转换结果预览")
            
            # 设置转换按钮的点击事件
            convert_excel_btn.click(
                fn=excel_to_other_interface,
                inputs=[excel_input, output_format_radio, sheet_name_input],
                outputs=[excel_preview_output, excel_info_output, other_output]
            )
            
            # 设置清理按钮的点击事件
            clear_excel_btn.click(
                fn=clear_cache,
                outputs=excel_cache_info
            )
            
            # 也支持文件上传后自动转换
            excel_input.change(
                fn=excel_to_other_interface,
                inputs=[excel_input, output_format_radio, sheet_name_input],
                outputs=[excel_preview_output, excel_info_output, other_output]
            )
            
            # 添加使用说明
            gr.Markdown("## 使用说明")
            gr.Markdown("""
            1. 点击"上传Excel文件"按钮，选择要转换的Excel文件
            2. 选择输出格式（CSV或JSON）
            3. 可选：输入要转换的工作表名称或索引（默认为0，即第一个工作表）
            4. 系统会自动开始转换，或点击"转换"按钮手动开始
            5. 在右侧可以预览转换后的内容
            6. 可以下载完整的转换文件
            
            **支持的文件格式：**
            - Excel：支持 .xlsx 和 .xls 格式
            - 输出格式：CSV（逗号分隔，同test_with_header.txt格式）和JSON（同test_new.json格式）
            
            **转换规则：**
            - CSV：使用逗号作为分隔符，第一行作为表头（如果有）
            - JSON：使用与test_new.json相同的格式，将每行数据转换为一个键值对
            """)
        
        # 第三个选项卡：任意查找
        with gr.TabItem("任意查找"):
            gr.Markdown("支持上传TXT、JSON或Excel文件，查找匹配的数据")
            
            with gr.Row():
                with gr.Column(scale=1):
                    search_file_input = gr.File(label="上传文件", file_types=[".txt", ".json", ".xlsx", ".xls"])
                    search_sheet_input = gr.Textbox(
                        label="工作表名称或索引（仅Excel文件，默认使用第一个工作表）", 
                        value="0", 
                        lines=1
                    )
                    search_query_input = gr.Textbox(
                        label="查找内容", 
                        placeholder="输入要查找的内容...", 
                        lines=1
                    )
                    search_btn = gr.Button("查找", variant="primary")
                    clear_search_btn = gr.Button("清理缓存", variant="secondary")
                    search_info_output = gr.Textbox(label="查找信息", lines=5, interactive=False)
                    search_cache_info = gr.Textbox(label="缓存状态", lines=2, interactive=False)
                    clear_file_btn = gr.Button("清空文件", variant="secondary")
                
                with gr.Column(scale=2):
                    full_table_output = gr.HTML(label="完整表格")
                    search_result_output = gr.HTML(label="查找结果")
            
            # 设置查找按钮的点击事件
            search_btn.click(
                fn=search_interface,
                inputs=[search_file_input, search_sheet_input, search_query_input],
                outputs=[full_table_output, search_result_output, search_info_output]
            )
            
            # 设置清理缓存按钮的点击事件
            clear_search_btn.click(
                fn=clear_cache,
                outputs=search_cache_info
            )
            
            # 设置清空文件按钮的点击事件
            clear_file_btn.click(
                fn=lambda: ("", "", "已清空文件"),
                outputs=[full_table_output, search_result_output, search_info_output]
            )
            
            # 文件上传后自动显示表格
            search_file_input.change(
                fn=lambda file, sheet: search_interface(file, sheet, ""),
                inputs=[search_file_input, search_sheet_input],
                outputs=[full_table_output, search_result_output, search_info_output]
            )
            
            # 添加使用说明
            gr.Markdown("## 使用说明")
            gr.Markdown("""
            1. 点击"上传文件"按钮，选择要查找的文件（支持TXT、JSON、Excel格式）
            2. 对于Excel文件，可选：输入要查找的工作表名称或索引（默认为0，即第一个工作表）
            3. 在"查找内容"输入框中输入要查找的关键词
            4. 点击"查找"按钮开始查找
            5. 在右侧可以查看完整表格和查找结果
            
            **支持的文件格式：**
            - TXT：支持自动检测分隔符，自动识别表头
            - JSON：支持两种格式：
              - 列表格式：`[{"key1": value1, "key2": value2}, ...]`
              - 字典格式：`{"column1": [value1, value2, ...], "column2": [...], ...}`
            - Excel：支持 .xlsx 和 .xls 格式
            
            **查找规则：**
            - 支持模糊匹配，不要求完全匹配
            - 支持搜索所有列中的数据
            - 忽略大小写
            """)
        
        # 第四个选项卡：行和列操作
        with gr.TabItem("行和列操作"):
            gr.Markdown("支持上传TXT、JSON或Excel文件，对表格进行行和列的增删操作")
            
            with gr.Row():
                with gr.Column(scale=1):
                    # 文件加载部分
                    table_file_input = gr.File(label="上传文件", file_types=[".txt", ".json", ".xlsx", ".xls"])
                    table_sheet_input = gr.Textbox(
                        label="工作表名称或索引（仅Excel文件，默认使用第一个工作表）", 
                        value="0", 
                        lines=1
                    )
                    load_btn = gr.Button("加载文件", variant="primary")
                    
                    gr.Markdown("---")
                    
                    # 行和列操作部分
                    gr.Markdown("### 行和列操作")
                    
                    # 添加操作
                    gr.Markdown("#### 添加操作")
                    add_content_input = gr.Textbox(
                        label="添加内容", 
                        placeholder="添加行：输入行内容，用逗号分隔（例：1,2.5,测试）；添加列：先输入列名，再输入列内容，用逗号分隔（例：新列,1,2,3,4,5）", 
                        lines=1
                    )
                    
                    with gr.Row():
                        add_row_btn = gr.Button("添加行")
                        add_column_btn = gr.Button("添加列")
                    
                    gr.Markdown("---")
                    
                    # 删除操作
                    gr.Markdown("#### 删除操作")
                    delete_content_input = gr.Textbox(
                        label="删除内容", 
                        placeholder="删除行：输入行索引（从1开始）；删除列：输入列名或列索引（从1开始）", 
                        lines=1
                    )
                    
                    with gr.Row():
                        delete_row_btn = gr.Button("删除行")
                        delete_column_btn = gr.Button("删除列")
                    
                    gr.Markdown("---")
                    
                    # 导出操作部分
                    gr.Markdown("### 导出操作")
                    export_format_input = gr.Radio(
                        label="导出格式", 
                        choices=["Excel", "CSV", "JSON"], 
                        value="Excel"
                    )
                    export_btn = gr.Button("导出表格", variant="primary")
                    
                    # 信息输出
                    table_info_output = gr.Textbox(label="操作信息", lines=5, interactive=False)
                    
                    # 下载文件
                    table_output = gr.File(label="下载文件")
                
                with gr.Column(scale=2):
                    # 表格预览（改为可编辑的DataFrame）
                    table_preview_output = gr.DataFrame(label="表格预览", interactive=True, row_count=1, column_count=1)
            
            # 设置按钮的点击事件
            # 加载文件按钮
            load_btn.click(
                fn=load_data,
                inputs=[table_file_input, table_sheet_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 文件上传后自动加载
            table_file_input.change(
                fn=lambda file, sheet: load_data(file, sheet),
                inputs=[table_file_input, table_sheet_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 工作表名称/索引变化时重新加载
            table_sheet_input.change(
                fn=lambda file, sheet: load_data(file, sheet),
                inputs=[table_file_input, table_sheet_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 添加行按钮
            add_row_btn.click(
                fn=add_row,
                inputs=[add_content_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 删除行按钮
            delete_row_btn.click(
                fn=delete_row,
                inputs=[delete_content_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 添加列按钮
            add_column_btn.click(
                fn=lambda content: add_column(*parse_add_column_content(content)),
                inputs=[add_content_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 删除列按钮
            delete_column_btn.click(
                fn=delete_column,
                inputs=[delete_content_input],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 导出表格按钮
            export_btn.click(
                fn=export_table,
                inputs=[export_format_input],
                outputs=[table_output, table_info_output]
            )
            
            # 当表格内容变化时，更新全局变量current_df
            def update_current_df(dataframe):
                global current_df
                current_df = dataframe
                return dataframe, "表格内容已更新"
            
            table_preview_output.change(
                fn=update_current_df,
                inputs=[table_preview_output],
                outputs=[table_preview_output, table_info_output]
            )
            
            # 添加使用说明
            gr.Markdown("## 使用说明")
            gr.Markdown("""
            1. 点击"上传文件"按钮，选择要操作的文件（支持TXT、JSON、Excel格式）
            2. 对于Excel文件，可选：输入要操作的工作表名称或索引（默认为0，即第一个工作表）
            3. 点击"加载文件"按钮加载数据
            4. 使用行操作区域的按钮进行行的添加和删除
            5. 使用列操作区域的按钮进行列的添加和删除
            6. 使用导出操作区域的按钮将修改后的表格导出
            
            **支持的文件格式：**
            - TXT：支持自动检测分隔符，自动识别表头
            - JSON：支持两种格式：
              - 列表格式：`[{"key1": value1, "key2": value2}, ...]`
              - 字典格式：`{"column1": [value1, value2, ...], "column2": [...], ...}`
            - Excel：支持 .xlsx 和 .xls 格式
            
            **操作规则：**
            - 行和列索引：从1开始，例如要删除第一行，请输入"1"
            - 直接编辑：您可以直接在表格中编辑任何单元格的内容，编辑后会自动保存
            - 列名：输入要添加或删除的列的名称
            - 导出格式：支持Excel、CSV和JSON格式
            """)
        
        # 第五个选项卡：SQLite转换
        with gr.TabItem("SQLite转换"):
            gr.Markdown("支持SQLite数据库与Excel文件之间的相互转换")
            
            # SQLite转Excel部分
            with gr.Row():
                with gr.Column():
                    gr.Markdown("## SQLite转Excel")
                    sqlite_file_input = gr.File(label="上传SQLite数据库文件", file_types=[".db"])
                    tables_input = gr.Textbox(
                        label="要转换的表名（可选，默认转换所有表，多个表用逗号分隔）", 
                        placeholder="表1,表2,表3", 
                        lines=1
                    )
                    sqlite_sort_checkbox = gr.Checkbox(label="按Unicode编码排序字段名", value=False)
                    sqlite_to_excel_btn = gr.Button("SQLite转Excel", variant="primary")
                    sqlite_to_excel_output = gr.File(label="下载Excel文件")
                    sqlite_to_excel_info = gr.Textbox(label="转换信息", lines=3, interactive=False)
                
                # Excel转SQLite部分
                with gr.Column():
                    gr.Markdown("## Excel转SQLite")
                    excel_to_sqlite_file_input = gr.File(label="上传Excel文件", file_types=[".xlsx", ".xls"])
                    table_name_input = gr.Textbox(
                        label="输出表名（可选，默认使用Excel工作表名）", 
                        placeholder="新表名", 
                        lines=1
                    )
                    excel_sheet_input = gr.Textbox(
                        label="工作表名称或索引（默认使用第一个工作表）", 
                        value="0", 
                        lines=1
                    )
                    calculate_functions_checkbox = gr.Checkbox(label="计算Excel函数", value=True)
                    excel_to_sqlite_btn = gr.Button("Excel转SQLite", variant="primary")
                    excel_to_sqlite_output = gr.File(label="下载SQLite数据库文件")
                    excel_to_sqlite_info = gr.Textbox(label="转换信息", lines=3, interactive=False)
            
            # 设置SQLite转Excel按钮的点击事件
            sqlite_to_excel_btn.click(
                fn=sqlite_to_excel_interface,
                inputs=[sqlite_file_input, tables_input, sqlite_sort_checkbox],
                outputs=[sqlite_to_excel_output, sqlite_to_excel_info]
            )
            
            # 设置Excel转SQLite按钮的点击事件
            excel_to_sqlite_btn.click(
                fn=excel_to_sqlite_interface,
                inputs=[excel_to_sqlite_file_input, table_name_input, excel_sheet_input, calculate_functions_checkbox],
                outputs=[excel_to_sqlite_output, excel_to_sqlite_info]
            )
            
            # 添加使用说明
            gr.Markdown("## 使用说明")
            gr.Markdown("""
            ### SQLite转Excel
            1. 点击"上传SQLite数据库文件"按钮，选择要转换的SQLite数据库文件(.db)
            2. 可选：输入要转换的表名（多个表用逗号分隔，默认转换所有表）
            3. 可选：勾选"按Unicode编码排序字段名"选项
            4. 点击"SQLite转Excel"按钮开始转换
            5. 下载转换后的Excel文件
            
            ### Excel转SQLite
            1. 点击"上传Excel文件"按钮，选择要转换的Excel文件(.xlsx或.xls)
            2. 可选：输入输出表名（默认使用Excel工作表名）
            3. 可选：输入要转换的工作表名称或索引（默认为0，即第一个工作表）
            4. 可选：勾选"计算Excel函数"选项（将计算结果存储到SQLite，默认开启）
            5. 点击"Excel转SQLite"按钮开始转换
            6. 下载转换后的SQLite数据库文件
            
            **支持的功能：**
            - SQLite转Excel：支持转换整个数据库或指定表，支持Unicode排序
            - Excel转SQLite：支持转换Excel文件中的指定工作表，支持计算Excel函数
            """)

if __name__ == "__main__":
    print("启动Gradio应用...")
    # 使用local模式，避免外部资源加载问题
    app.launch(share=False, inbrowser=True, server_name="127.0.0.1", server_port=7861)
