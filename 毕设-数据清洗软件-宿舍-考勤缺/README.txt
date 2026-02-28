运行方式：
1) 确保已安装依赖：PySide6, pandas, numpy, xlsxwriter
   （如需解析 Word/PDF：python-docx, pdfplumber）
2) 进入本项目目录，执行：
   python main.py

文件结构：
- lh_exporter/ui.py            UI 组件（NiceProgressBar、拖拽输入框、卡片、空白点击过滤器）
- lh_exporter/processing.py    数据处理/导出核心逻辑（来自你的第一份代码）
- lh_exporter/workers.py       三个功能的线程 Worker（界面一/界面二/简报中心）
- lh_exporter/dialogs.py       数据导入弹窗、设置中心、院系词典中心、使用说明
- lh_exporter/main_window.py   主窗口（三列 UI + 原三大功能页）
- main.py                      启动入口
