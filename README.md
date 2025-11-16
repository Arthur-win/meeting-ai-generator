# 会议记录生成器 (Meeting AI Generator)

基于 AI 的智能会议记录生成系统，支持从自然语言描述中自动提取会议信息并生成专业的 Word 文档。

## ✨ 功能特点

- 🤖 **智能信息提取**：使用正则表达式和 Ollama AI 模型自动提取会议关键信息
- 📝 **自然语言理解**：支持多种中文自然语言表述方式
- 📄 **专业文档生成**：自动生成格式规范的 Word 会议记录文档
- 🎯 **会议类型识别**：自动识别技术会议、商务会议、项目会议、团队会议
- 🌐 **现代化界面**：简洁美观的 Web 界面，操作便捷

## 🚀 快速开始

### 环境要求

- Python 3.8+
- 现代浏览器（Chrome、Firefox、Edge 等）
- （可选）Ollama 本地 AI 模型

### 安装步骤

1. **克隆项目**
```bash
git clone https://github.com/your-username/meeting-ai-generator.git
cd meeting-ai-generator
```

2. **安装依赖**
```bash
pip install -r backend/requirements.txt
```

3. **启动后端服务**
```bash
cd backend
python app.py
```

4. **打开前端页面**
```bash
# 在浏览器中打开
frontend/index.html
```

## 📖 使用说明

### 输入示例

在文本框中输入会议描述，例如：

```
下周三下午三点，我想在公司三楼的大会议室开个会，参会的有市场部的李明、张娜，
技术部的王磊，还有新来的实习生赵晓雨。会议主要说三件事，一是讨论下季度的产品
推广方案，特别是线上广告投放的预算分配，李明你准备下相关数据；二是同步下新版
APP的研发进度，王磊负责这块；三是安排下月底的团建活动，张娜你牵头，大家看看
想去近郊的民宿还是去爬山，赵晓雨可以多听听大家的意见。对了，会议大概开两个
小时，记得提前把会议资料发到群里。
```

### 功能按钮

- **提取会议信息**：预览提取的结构化会议信息
- **生成文档**：直接生成并下载 Word 文档
- **清空输入**：清除输入框内容

## 🛠️ 技术栈

### 后端
- **Flask**：轻量级 Web 框架
- **Flask-CORS**：跨域资源共享支持
- **python-docx**：Word 文档生成
- **Ollama**：本地 AI 模型（可选）

### 前端
- **原生 HTML/CSS/JavaScript**：无框架依赖
- **现代化 UI 设计**：渐变色、响应式布局

## 📁 项目结构

```
meeting-ai-generator/
├── backend/
│   ├── app.py                 # Flask 后端主程序
│   ├── word_generator.py      # Word 文档生成器
│   ├── requirements.txt       # Python 依赖
│   └── output/               # 生成的文档输出目录
├── frontend/
│   └── index.html            # Web 前端页面
├── .gitignore                # Git 忽略文件
└── README.md                 # 项目说明文档
```

## 🔧 API 接口

### POST /extract
提取会议信息

**请求体：**
```json
{
  "text": "会议描述文本"
}
```

**响应：**
```json
{
  "success": true,
  "data": {
    "theme": "会议主题",
    "host": "主持人",
    "location": "会议地点",
    "attendees": "参会人员",
    "duration": "会议时长",
    "topics": [
      {
        "topic": "议题内容",
        "leader": "负责人",
        "preparation": "会前准备"
      }
    ]
  }
}
```

### POST /generate
生成并下载 Word 文档

**请求体：**
```json
{
  "text": "会议描述文本"
}
```

**响应：** Word 文档文件流

### GET /health
健康检查

**响应：**
```json
{
  "status": "healthy",
  "message": "服务运行正常"
}
```

## 🎯 支持的会议类型

- **技术会议**：技术讨论、开发评审、架构设计等
- **商务会议**：商务洽谈、市场分析、客户会议等
- **项目会议**：项目进度、任务分配、里程碑评审等
- **团队会议**：周会、例会、团队分享等

## 📝 提取的信息字段

- 会议主题
- 主持人
- 会议地点
- 参会人员
- 会议时长
- 会议议题（支持多个）
  - 议题内容
  - 负责人
  - 会前准备事项

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

1. Fork 本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 🙏 致谢

- [Flask](https://flask.palletsprojects.com/) - Web 框架
- [python-docx](https://python-docx.readthedocs.io/) - Word 文档处理
- [Ollama](https://ollama.ai/) - 本地 AI 模型

## 📧 联系方式

如有问题或建议，欢迎通过 Issue 联系我们。

---

⭐ 如果这个项目对你有帮助，请给个 Star！
