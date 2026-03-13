# 文档转Excel工具

一个功能强大的文档转Excel工具集，支持将TXT和JSON文件转换为Excel格式，并提供友好的Web界面。

## 功能特性

### TXT转Excel (txt_to_excel.py)
- ✅ 自动检测TXT文件的分隔符
- ✅ 智能识别表头
- ✅ 生成原始数据工作表
- ✅ 为每一列创建降序排序的工作表
- ✅ 完善的错误处理机制
- ✅ 自动处理文件冲突

### JSON转Excel (json_to_excel.py)
- ✅ 支持多种JSON结构：
  - 列表格式：`[{"key1": value1, "key2": value2}, ...]`
  - 嵌套字典：`{"row1": {"col1": value1}, "row2": {"col1": value2}}`
  - 字典格式：`{"column1": [value1, value2], "column2": [value1, value2]}`
- ✅ 自动转换文本数字为数字类型
- ✅ 支持JSON中的函数定义，直接转换为Excel公式
  - SUM、AVERAGE、MAX、MIN等常用函数
  - 支持相对单元格引用
- ✅ 生成原始数据和排序后的工作表

### Gradio Web界面 (gradio_app.py)
- ✅ 直观的文件上传界面
- ✅ 实时预览转换结果
- ✅ 支持下载Excel文件
- ✅ 友好的使用说明和帮助文档
- ✅ 本地运行，保护数据隐私

## 安装说明

### 系统要求
- Python 3.6+
- Windows系统（批处理脚本基于Windows编写）

### 自动安装

项目已包含运行脚本，会自动安装所需依赖。

### 手动安装

```bash
# 安装核心依赖
pip install pandas openpyxl

# 安装Gradio Web界面依赖（可选）
pip install gradio
```

## 使用方法

### 1. 命令行方式

#### TXT转Excel
```bash
python txt_to_excel.py <txt文件路径>
```

#### JSON转Excel
```bash
python json_to_excel.py <json文件路径>
```

或使用批处理文件：
```bash
run_json_to_excel.bat <json文件路径>
```

### 2. Web界面方式

运行批处理文件启动Web界面：

```bash
run_gradio.bat
```

然后在浏览器中打开 http://127.0.0.1:7860

## JSON函数定义格式

JSON转Excel工具支持在JSON中定义函数，格式如下：

```json
{
  "col3": {
    "type": "SUM",  // 函数类型
    "sub": [         // 要计算的单元格
      {"r": -2, "c": 0},  // r: 相对行数, c: 相对列数
      {"r": -1, "c": 0}
    ]
  }
}
```

支持的函数类型：
- SUM: 求和
- AVERAGE: 平均值
- MAX: 最大值
- MIN: 最小值

相对位置说明：
- r: 相对当前行的偏移（负数表示上方行，正数表示下方行）
- c: 相对当前列的偏移（负数表示左侧列，正数表示右侧列）

## 文件结构

```
toexcel/
├── txt_to_excel.py         # TXT转Excel脚本
├── json_to_excel.py        # JSON转Excel脚本
├── gradio_app.py           # Gradio Web界面
├── run_json_to_excel.bat   # JSON转换运行脚本
├── run_gradio.bat          # Web界面运行脚本
├── lib/                    # 本地依赖库目录
└── README.md               # 项目说明文档
```

## 示例

### TXT文件示例

```
id,name,age,salary
1,张三,25,8000
2,李四,30,12000
3,王五,28,10000
```

### JSON文件示例

```json
{
  "employee1": {
    "id": 1,
    "name": "张三",
    "age": 25,
    "salary": "8000",
    "total": {
      "type": "SUM",
      "sub": [{"r": -3, "c": 0}, {"r": -2, "c": 0}]
    }
  },
  "employee2": {
    "id": 2,
    "name": "李四",
    "age": 30,
    "salary": "12000"
  }
}
```

## 输出说明

转换后的Excel文件包含以下工作表：
- **Sheet1**: 原始数据
- **Sheet2**: 按第1列降序排序后的数据
- **Sheet3**: 按第2列降序排序后的数据
- **以此类推**...

## 注意事项

1. 确保文件路径中不含中文特殊字符
2. JSON文件必须符合UTF-8编码
3. 大型文件转换可能需要较长时间
4. Web界面默认运行在端口7860

## 错误处理

工具包含完善的错误处理机制，常见错误包括：
- 文件权限不足
- 无效的文件格式
- JSON语法错误
- 函数引用的单元格不存在

## 更新日志

### 3.13 update
- 修复JSON转Excel时，函数引用的单元格位置计算错误的问题
- 支持在JSON中指定绝对引用位置，例如`A1`、`B2`等
- 支持在JSON中定义函数，直接转换为Excel公式
- 支持在JSON中定义函数的参数，例如`SUM(A1:A10)`等

### v1.0.0
- 初始版本
- 支持TXT转Excel功能
- 支持JSON转Excel功能
- 支持Web界面
- 支持函数转换

## 改进方案

未来将继续改进程序，争取实现以下功能：
- 允许在JSON中指定绝对引用位置，例如`A1`、`B2`等
- 直接在Gradio界面对表格进行增、删、改的功能
- 通过特定的条件对表格进行分组，例如按某行数据大于或小于某个值进行分组

---

**使用提示：**
- 对于大型文件，建议使用命令行方式
- 对于包含函数的JSON文件，确保相对引用位置正确
- Web界面适合快速转换和预览小文件
