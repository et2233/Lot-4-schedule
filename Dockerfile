# 项目管理工具 - Docker 部署

FROM python:3.11-slim

WORKDIR /app

# 安装依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY engine.py app.py wsgi.py ./

# 复制 Excel 文件
COPY Construction_Schedule_with_GanttSheet.xlsx /data/

# 数据持久化目录
VOLUME ["/data"]

# 环境变量
ENV EXCEL_PATH=/data/Construction_Schedule_with_GanttSheet.xlsx
ENV DATA_DIR=/data
ENV PORT=5000

EXPOSE 5000

CMD ["sh", "-c", "python -m gunicorn -w 4 -b 0.0.0.0:${PORT} wsgi:app"]
