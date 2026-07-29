"""
项目管理工具 - 生产环境启动
使用 gunicorn (Linux) 或 waitress (跨平台) 作为 WSGI 服务器
"""

import os
import sys

# 确保能找到 engine 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app

if __name__ == '__main__':
    # 生产模式：关闭 debug
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 项目管理工具 (生产模式) 启动在端口 {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
