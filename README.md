# 央企项目经理竞聘简历生成器

上传 .doc/.docx 简历文件 → 自动提取文本和照片 → 编辑确认 → 一键生成 A4 竞聘简历 PDF。

## 设计风格

V4 · IBM Carbon × Apple × Stripe 三合一设计语言，专为基建央企项目经理竞聘场景定制。

## 技术栈

- Streamlit (Web UI)
- python-docx + olefile + antiword (文档解析)
- Jinja2 (HTML 模板渲染)
- Chrome Headless (PDF 转换，本地可用时)

## 本地运行

```bash
pip install -r requirements.txt
streamlit run resume_app.py
```

## 部署到 Streamlit Cloud

1. Fork 本仓库
2. 在 [share.streamlit.io](https://share.streamlit.io) 点击 "New app"
3. 选择仓库 + 分支 + `resume_app.py`
4. 点击 Deploy
