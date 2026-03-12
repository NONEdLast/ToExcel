import sys
import os
import uuid

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
            return None, f"错误：不支持的文件格式 {file_ext}，仅支持 .txt 和 .json", None
        
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

if __name__ == "__main__":
    print("启动Gradio应用...")
    # 使用local模式，避免外部资源加载问题
    app.launch(share=False, inbrowser=True, server_name="127.0.0.1", server_port=7860)
