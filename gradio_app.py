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

# 设置临时文件目录
temp_dir = "temp_files"
os.makedirs(temp_dir, exist_ok=True)

def gradio_interface(file, sort_by_unicode):
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
            success = txt_to_excel.txt_to_excel(file_path, temp_excel_path, sort_by_unicode)
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

# 创建Gradio界面
with gr.Blocks(title="文档转Excel工具") as app:
    gr.Markdown("# 文档转Excel工具")
    gr.Markdown("支持上传TXT和JSON文件，自动转换为Excel并提供预览")
    
    with gr.Row():
        with gr.Column(scale=1):
            file_input = gr.File(label="上传文件", file_types=[".txt", ".json"])
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
        inputs=[file_input, sort_checkbox],
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
        inputs=[file_input, sort_checkbox],
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
    
    **Unicode排序功能：**
    - 勾选"按Unicode编码对字符串排序"选项后，系统会为每个字符串列添加一个新列
    - 新列名格式为"原列名_unicode_sort"，包含按Unicode编码升序排序的序号
    - 支持识别和排序各种Unicode字符，包括中文、英文、数字和特殊字符
    """)

if __name__ == "__main__":
    print("启动Gradio应用...")
    # 使用local模式，避免外部资源加载问题
    app.launch(share=False, inbrowser=True, server_name="127.0.0.1", server_port=7860)
