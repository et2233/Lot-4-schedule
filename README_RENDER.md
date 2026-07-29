# 免费部署到 Render.com + 嵌入 Wix/网页

## 方案概述

因为你的网站是 Wix/Squarespace 这类建站工具，不能直接运行 Python。所以我们把 Python 程序放在免费的 **Render.com** 上跑，然后你的网页用 `iframe` 嵌入它。

## 第一步：上传代码到 GitHub

1. 去 [github.com](https://github.com) 注册/登录
2. 新建一个仓库（Repository），名字比如 `construction-schedule`
3. 把 `deploy_package` 里的所有文件上传到这个仓库
4. 重要：确保仓库里有这些文件：
   - `engine.py`
   - `app.py`
   - `wsgi.py`
   - `requirements.txt`
   - `render.yaml`
   - `Construction_Schedule_with_GanttSheet.xlsx`

## 第二步：在 Render.com 部署

1. 去 [render.com](https://render.com) 注册/登录（可以用 GitHub 账号直接登录）
2. 点击 **New +** → **Web Service**
3. 选择你的 GitHub 仓库 `construction-schedule`
4. 配置：
   - **Name**: `construction-schedule`
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn -w 2 -b 0.0.0.0:10000 wsgi:app`
   - **Plan**: Free
5. 点击 **Create Web Service**
6. 等待 2-3 分钟部署完成
7. 你会得到一个地址，比如：`https://construction-schedule.onrender.com`

## 第三步：嵌入到你的网页

### 在你的 Wix/Squarespace 网页 HTML 框中放入：

```html
<iframe
  src="https://construction-schedule.onrender.com"
  width="100%"
  height="800"
  style="border: none;"
  allowfullscreen>
</iframe>
```

把 `https://construction-schedule.onrender.com` 替换成 Render 给你的真实地址。

## 第四步：微信小程序访问

小程序页面：

```html
<web-view src="https://construction-schedule.onrender.com"></web-view>
```

小程序后台需要添加：
- **request 合法域名**: `https://construction-schedule.onrender.com`
- **业务域名**: `https://construction-schedule.onrender.com`

## ⚠️ 重要提醒

1. **Render Free 有冷启动**：如果 15 分钟没人访问，下次打开会慢 30 秒左右
2. **数据存在 Render 磁盘上**，免费版重启后数据可能丢失！建议升级到 $7/月 保留磁盘
3. 免费版每月有使用限制，日常够用

## 不想用 GitHub？

也可以直接把 `deploy_package` 压缩成 zip，在 Render 上传部署。或者联系我帮你部署。
