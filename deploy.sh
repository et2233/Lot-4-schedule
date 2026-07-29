#!/bin/bash
# ============================================
# 项目管理工具 - 一键部署脚本
# 适用于 Ubuntu 20.04+ / Debian 11+
# ============================================
set -e

APP_DIR="/opt/schedule"
DOMAIN="${1:-}"
PORT=5000

if [ -z "$DOMAIN" ]; then
    echo "用法: bash deploy.sh your-domain.com"
    echo "请提供你的域名（需已解析到本服务器 IP）"
    exit 1
fi

echo "========================================"
echo "  建筑施工进度管理 - 一键部署"
echo "  域名: $DOMAIN"
echo "========================================"

# 1. 安装依赖
echo "[1/5] 安装系统依赖..."
apt-get update -qq
apt-get install -y -qq python3 python3-pip python3-venv nginx certbot python3-certbot-nginx

# 2. 创建应用目录
echo "[2/5] 创建应用目录..."
mkdir -p $APP_DIR
cp -r /tmp/deploy_package/* $APP_DIR/
cd $APP_DIR

# 3. 安装 Python 依赖
echo "[3/5] 安装 Python 依赖..."
pip3 install -r requirements.txt

# 4. 配置 Nginx + HTTPS
echo "[4/5] 配置 Nginx..."
cat > /etc/nginx/sites-available/schedule <<EOF
server {
    listen 80;
    server_name $DOMAIN;
    
    client_max_body_size 50M;

    location / {
        proxy_pass http://127.0.0.1:$PORT;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
EOF

ln -sf /etc/nginx/sites-available/schedule /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl reload nginx

# 申请 SSL 证书
echo "[4.5/5] 申请 SSL 证书..."
certbot --nginx -d $DOMAIN --non-interactive --agree-tos --email admin@$DOMAIN || true

# 5. 配置 systemd 服务
echo "[5/5] 配置 systemd 服务..."
cat > /etc/systemd/system/schedule.service <<EOF
[Unit]
Description=Construction Schedule Manager
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=$APP_DIR
ExecStart=/usr/bin/python3 -m gunicorn -w 4 -b 127.0.0.1:$PORT wsgi:app
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable schedule
systemctl start schedule

echo ""
echo "========================================"
echo "  ✅ 部署完成!"
echo "  访问地址: https://$DOMAIN"
echo "========================================"
echo ""
echo "微信小程序配置:"
echo "  1. 小程序后台 → 开发管理 → 服务器域名"
echo "  2. request 合法域名添加: https://$DOMAIN"
echo "  3. 业务域名添加: https://$DOMAIN"
echo "  4. 使用 web-view 组件加载"
