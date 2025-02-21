# ProposalLLM

A powerful AI-driven proposal generation tool that leverages large language models to automate the creation of technical proposals and responses.

[English](#english) | [中文](#chinese)

<a name="english"></a>
## English Version

### Features

1. **Automated Proposal Generation**
   - Generates point-by-point response format proposals based on requirements matrix
   - Automatically formats content with proper heading levels (1, 2, 3)
   - Preserves formatting for body text, images, and bullet points
   - Supports automatic chapter numbering
   - Detailed execution logging with timing information
   - Preserves original input files by saving outputs to separate directory

2. **Smart Content Management**
   - Automatically copies relevant product documentation to responses
   - Preserves all formatting including images, tables, and bullet points
   - Can rewrite product descriptions to match different proposal requirements
   - Uses AI to generate content for missing features
   - Supports both OpenAI and Baidu AI models

3. **Requirements Processing**
   - Generates technical requirements deviation tables
   - Automatically fills point-by-point responses in requirements matrix
   - Uses format: "Answer: Fully supported, {AI-generated response}"
   - Adds corresponding chapter numbers from the proposal
   - Provides detailed progress tracking for each requirement

4. **Document Processing**
   - Breaks down product manuals into reusable components
   - Improves proposal generation performance
   - Enables customization of features as needed
   - Supports up to 3 levels of headings
   - Maintains separate input and output directories for safe file handling

### Setup and Installation

1. **Python Environment Setup**
   ```bash
   # Create a virtual environment
   python -m venv venv
   
   # Activate virtual environment
   # On Windows
   .\venv\Scripts\activate
   # On Unix or MacOS
   source venv/bin/activate
   
   # Verify Python environment
   python --version
   ```

2. **Dependencies Installation**
   ```bash
   # Install dependencies
   pip install -r requirements.txt
   
   # Configure environment variables
   cp .env.example .env
   # Edit .env file with your API keys
   ```

3. **API Configuration**
   - Configure either ChatGPT or Baidu Qianfan model API
   - ERNIE-Speed-8K (free model) is recommended for Baidu
   - Add API keys to .env file
   - Set `USE_BAIDU=true/false` to switch between APIs
   - Configure model parameters (temperature, max tokens) in .env

4. **Document Preparation**
   - Place input files in `data/input/` directory:
     - Product manual as `标书内容.docx`
     - Requirements matrix as `需求对应表.xlsx`
   - Template file should be in `data/templates/Template.docx`
   - Ensure proper use of styles:
     - Body Text
     - Heading 1
     - Heading 2
     - Heading 3

### Usage

1. **Document Extraction**
   ```bash
   python src/Extract_Word.py
   ```
   - Generates component documents from product manual
   - Verify generated files for accuracy

2. **Requirements Setup**
   - Fill in the requirements table (`需求对应表.xlsx`):
     - Column B: Main requirements
     - Column C: Sub-requirements (generates level 2 and 3 headings)
     - Column G: Corresponding product manual section (use 'X' if none)

3. **Proposal Generation**
   ```bash
   python src/Generate.py
   ```
   - Generated files will be saved in `data/output/` directory:
     - `需求对应表_输出.xlsx`: Updated requirements matrix
     - `标书内容_输出.docx`: Generated proposal document
   - Progress and timing information will be displayed during execution

### Project Structure

```
ProposalLLM/
├── data/                    # Data directory
│   ├── input/              # Input files (original files)
│   ├── output/             # Output files (generated files)
│   └── templates/          # Template files
├── examples/               # Example files
├── src/                    # Source code
├── .env.example           # Environment template
├── requirements.txt       # Dependencies
└── README.md             # Documentation
```

<a name="chinese"></a>
## 中文版本

### 功能特点

1. **自动化标书生成**
   - 根据需求对应表自动生成点对点应答格式标书
   - 自动设置标题1、2、3级格式
   - 保持正文、图片、项目符号等格式
   - 支持自动章节编号
   - 详细的执行日志和时间统计
   - 输出文件保存在独立目录，保护原始文件

2. **智能内容管理**
   - 自动从产品文档复制相关内容到应答中
   - 完整保留图片、表格、项目符号等格式
   - 可根据不同标书需求自动重写产品描述
   - 对缺失功能使用AI自动生成内容
   - 支持OpenAI和百度AI模型

3. **需求处理**
   - 生成技术需求偏离表
   - 自动填写需求对应表中的点对点应答
   - 使用格式："答：全面支持，{AI生成的回应}"
   - 自动填写对应的标书章节号
   - 提供每个需求的处理进度跟踪

4. **文档处理**
   - 将产品手册拆分为可复用的组件
   - 提高标书生成性能
   - 支持针对性功能修改
   - 支持最多3级标题
   - 采用独立的输入输出目录，安全处理文件

### 环境配置

1. **Python环境配置**
   ```bash
   # 创建虚拟环境
   python -m venv venv
   
   # 激活虚拟环境
   # Windows系统
   .\venv\Scripts\activate
   # Unix或MacOS系统
   source venv/bin/activate
   
   # 验证Python环境
   python --version
   ```

2. **安装依赖**
   ```bash
   # 安装依赖
   pip install -r requirements.txt
   
   # 配置环境变量
   cp .env.example .env
   # 编辑.env文件，填入API密钥
   ```

3. **API配置**
   - 配置ChatGPT或百度千帆模型API
   - 推荐使用百度ERNIE-Speed-8K（免费模型）
   - 在.env文件中添加API密钥
   - 通过`USE_BAIDU=true/false`切换API
   - 在.env中配置模型参数（温度、最大token数）

4. **文档准备**
   - 在`data/input/`目录中放置输入文件：
     - 产品说明手册：`标书内容.docx`
     - 需求对应表：`需求对应表.xlsx`
   - 模板文件放在`data/templates/Template.docx`
   - 确保正确使用以下样式：
     - 正文
     - 标题1
     - 标题2
     - 标题3

### 使用方法

1. **文档提取**
   ```bash
   python src/Extract_Word.py
   ```
   - 生成产品手册对应的组件文档
   - 验证生成文件的准确性

2. **需求设置**
   - 填写需求对应表（`需求对应表.xlsx`）：
     - B列：主要需求
     - C列：子需求（用于生成二级、三级标题）
     - G列：对应产品说明书章节（如无则填'X'）

3. **标书生成**
   ```bash
   python src/Generate.py
   ```
   - 生成的文件将保存在`data/output/`目录：
     - `需求对应表_输出.xlsx`：更新后的需求对应表
     - `标书内容_输出.docx`：生成的标书文档
   - 执行过程中会显示进度和时间统计信息

### 项目结构

```
ProposalLLM/
├── data/                    # 数据目录
│   ├── input/              # 输入文件（原始文件）
│   ├── output/             # 输出文件（生成文件）
│   └── templates/          # 模板文件
├── examples/               # 示例文件
├── src/                    # 源代码
├── .env.example           # 环境变量模板
├── requirements.txt       # 项目依赖
└── README.md             # 项目文档
