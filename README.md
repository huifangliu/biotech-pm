# 生物项目管理系统 MVP

当前版本是一个无第三方依赖的 MVP：

- 前端：`frontend/` 静态页面，可部署到 Vercel
- 后端：`backend/server.py`，Python 标准库 HTTP API
- 数据库：SQLite，文件在 `backend/data/biotech_pm.sqlite3`
- Excel：直接读取和导出 `.xlsx`

## 本地运行

```bash
cd biotech_pm_fullstack
python3 backend/server.py --host 0.0.0.0 --port 8000
```

浏览器打开：

```text
http://服务器IP:8000
```

默认账号：

```text
admin / admin123
demo  / demo123
```

## 已实现功能

- 登录、申请账号、忘记密码入口
- 当前账号相关项目列表
- 创建项目
- 项目详情左侧流程导航
- `1.样本信息表.xlsx` 导入样本入库
- 样本核对增删改查
- `2.实验室上机.xlsx` 导入建库、上机、浓度、操作人等信息
- 数据存放路径和备注维护
- 生信任务维护
- 导出 `生信信息表.xlsx`
- 合同文件上传和下载
- 回款总额、到账金额、回款进度
- 导出 `项目信息汇总.xlsx`

## Vercel 前端部署

把 `frontend/` 作为 Vercel 项目根目录。

如果后端不是同域名，把 `frontend/config.js` 改成：

```js
window.BIOTECH_PM_API_BASE = "https://api.example.com";
```

## 服务器 Docker 运行

```bash
cd biotech_pm_fullstack
docker compose up -d --build
```

生产环境建议在前面加 Nginx 或 Caddy，并启用 HTTPS。
