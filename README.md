# 星火 Office

这是一个支持 DOCX、XLSX、PPTX 的自托管浏览器 Office 编辑器。

文档通过浏览器内存传入编辑器，使用 ONLYOFFICE Web Apps 与 x2t WebAssembly 在用户设备上处理。最终镜像不运行 DocumentServer，不向 Office 后端上传文档，也不再依赖 `office-editor.ziziyi.com`。

## 隐私与安全

- 编辑器脚本、字体、Worker 和 WebAssembly 全部由同源服务器提供。
- 已禁用第三方 Office 插件和网盘集成。
- 严格的内容安全策略会阻止第三方网络请求。
- 构建会扫描旧作者资源域名、统计服务域名等残留，发现后直接失败。
- 星火通过限定来源的 `postMessage` 桥接，并以 `ArrayBuffer` 传输文档。

## 构建镜像

服务器需要安装 Docker，并能获取以下官方基础镜像：

```text
onlyoffice/documentserver:9.4.0.1
node:22-alpine
caddy:2-alpine
```

构建：

```bash
docker build \
  --build-arg DS_VERSION=9.4.0.1 \
  --build-arg HASH=1 \
  -t xinghuo-office:9.4.0.1-3 \
  .
```

ONLYOFFICE 镜像只在构建阶段用于提取静态资源，不会在最终容器中运行完整 DocumentServer。

启动：

```bash
docker run -d \
  --name xinghuo-office \
  --restart unless-stopped \
  -p 127.0.0.1:18080:80 \
  xinghuo-office:9.4.0.1-3
```

## 1Panel 编排

先将仓库克隆到 `/opt/office-website` 并执行上面的构建命令，然后使用：

```yaml
services:
  office:
    image: xinghuo-office:9.4.0.1-3
    container_name: xinghuo-office
    restart: unless-stopped
    ports:
      - "127.0.0.1:18080:80"
```

在 1Panel 网站中将自己的域名反向代理到 `http://127.0.0.1:18080`，并开启 HTTPS。正式环境强烈建议使用域名和 HTTPS；HTTPS 的星火网页版通常会阻止嵌入 HTTP IP 地址。

## 本地开发

```bash
pnpm install
pnpm dev
```

本项目及所使用的 ONLYOFFICE Community 组件采用 AGPL-3.0 兼容条款。通过网络提供修改后的应用时，请保留相应声明并公开对应源代码修改。
