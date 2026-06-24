# 西邮物流看板服务器部署交付说明

## 交付文件

请将以下文件交给负责服务器或网站部署的同事：

- `westernpost-dashboard-prototype.html`：看板网页原型，部署时可改名为 `index.html`。
- `westernpost-dashboard-recommendation.md`：看板设计说明和数据框架文档，可作为评审资料。
- `westernpost-dashboard-server-handoff.md`：本部署说明。

## 推荐部署方式

这是一个静态 HTML 原型，最简单的上线方式是使用 Nginx、Apache、宝塔面板、对象存储静态网站或公司已有门户系统托管。

推荐服务器目录：

```text
/var/www/westernpost-dashboard
```

部署后文件结构：

```text
/var/www/westernpost-dashboard/
  index.html
```

将 `westernpost-dashboard-prototype.html` 上传到该目录，并改名为：

```text
index.html
```

## Nginx 配置示例

```nginx
server {
    listen 80;
    server_name dashboard.wpglb.com;

    root /var/www/westernpost-dashboard;
    index index.html;

    location / {
        try_files $uri $uri/ /index.html;
    }
}
```

配置后执行：

```bash
sudo nginx -t
sudo systemctl reload nginx
```

访问地址示例：

```text
http://dashboard.wpglb.com
```

如果暂时没有域名，也可以先使用：

```text
http://服务器IP
```

## 权限与安全建议

该看板包含财务、客户、利润、仓库经营等敏感经营信息。正式对外或内部试用前，建议至少完成以下控制：

1. 配置 HTTPS。
2. 限制访问范围，例如公司 VPN、IP 白名单、Nginx Basic Auth 或统一登录。
3. 不要将含真实财务和客户数据的版本放到公开互联网。
4. 后续若接入真实 API，权限应在服务端控制，不能只依赖前端隐藏字段。

## 后续更新方式

如果只是更新页面版本，可以由设计/产品同事把新的 HTML 文件重新上传并覆盖服务器上的 `index.html`。

也可以使用 `rsync` 从本地同步：

```bash
rsync -avz westernpost-dashboard-prototype.html 用户名@服务器IP:/var/www/westernpost-dashboard/index.html
```

如果后续要做真实实时数据，需要新增后端 API：

```text
浏览器页面
  -> Nginx
  -> 后端 API
  -> WMS / TMS / OMS / ERP / CRM / 管理报表数据
```

当前原型中的动态数据为前端模拟数据，适合演示看板结构、角色视角和交互方式。
