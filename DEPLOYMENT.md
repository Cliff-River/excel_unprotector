# Excel Unprotector 部署指南

## 概述

本项目包含前端（Next.js + TypeScript）和后端（FastAPI + Python）两个部分，使用 Docker 和 Docker Compose 进行容器化部署。

## 环境要求

- Docker 20.10+
- Docker Compose 2.0+

## 快速开始

### 1. 构建并启动服务

在项目根目录下执行以下命令：

```bash
docker compose up -d --build
```

### 2. 访问应用

服务启动后，通过以下地址访问：

- **前端页面**: http://localhost:3000
- **后端 API**: http://localhost:8000
- **API 文档**: http://localhost:8000/docs

### 3. 停止服务

```bash
docker compose down
```

## 服务配置说明

### 前端服务 (frontend)

| 配置项 | 值 | 说明 |
|--------|-----|------|
| 端口映射 | 3000:3000 | 前端页面通过 3000 端口访问 |
| 健康检查 | / | 检查 Next.js 应用根路径 |
| CPU 限制 | 0.5 核 | 最大可用 CPU 资源 |
| 内存限制 | 256MB | 最大可用内存 |

### 后端服务 (backend)

| 配置项 | 值 | 说明 |
|--------|-----|------|
| 端口映射 | 8000:8000 | API 服务端口 |
| 健康检查 | /health | 直接检查 FastAPI 健康端点 |
| CPU 限制 | 1.0 核 | 最大可用 CPU 资源 |
| 内存限制 | 512MB | 最大可用内存 |

## 数据持久化

后端服务挂载了 `data` 目录，用于存放测试文件：

```yaml
volumes:
  - ./excel_unprotector_server/data:/app/data
```

## 环境变量

### 前端环境变量

| 变量名 | 默认值 | 说明 |
|--------|--------|------|
| NEXT_PUBLIC_API_BASE_URL | /api/ | API 基础路径 |
| PORT | 3000 | 服务监听端口 |

### 后端环境变量

| 变量名 | 默认值 | 说明 |
|--------|--------|------|
| PORT | 8000 | 服务监听端口 |

## 网络配置

两个服务通过自定义桥接网络 `excel-unprotector-network` 通信：

- 前端通过 `http://excel-unprotector-backend:8000` 访问后端 API
- 后端服务名 `excel-unprotector-backend` 在 Docker Compose 网络中可解析

## 健康检查

### 前端健康检查

```yaml
healthcheck:
  test: ["CMD", "wget", "--spider", "-q", "http://localhost:3000/"]
  interval: 30s
  timeout: 10s
  retries: 3
  start_period: 15s
```

### 后端健康检查

```yaml
healthcheck:
  test: ["CMD", "python", "-c", "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')"]
  interval: 30s
  timeout: 10s
  retries: 3
  start_period: 15s
```

## 服务依赖关系

前端服务依赖后端服务的健康状态：

```yaml
depends_on:
  backend:
    condition: service_healthy
```

这意味着只有当后端服务启动并通过健康检查后，前端服务才会启动。

## 常用命令

### 查看日志

```bash
# 查看所有服务日志
docker compose logs -f

# 查看前端日志
docker compose logs -f frontend

# 查看后端日志
docker compose logs -f backend
```

### 重新构建

```bash
# 重新构建所有服务
docker compose up -d --build

# 重新构建特定服务
docker compose up -d --build frontend
```

### 查看服务状态

```bash
docker compose ps
```

### 进入容器

```bash
# 进入前端容器
docker compose exec frontend sh

# 进入后端容器
docker compose exec backend bash
```

## 生产环境部署建议

1. **使用 HTTPS**: 在生产环境中，建议配置 Nginx 或使用负载均衡器提供 HTTPS 支持
2. **环境变量**: 根据实际需求调整环境变量配置
3. **资源限制**: 根据服务器配置调整 CPU 和内存限制
4. **日志管理**: 配置日志收集系统，便于监控和故障排查
5. **备份策略**: 定期备份数据目录
6. **版本控制**: 使用特定的 Docker 镜像标签，避免使用 `latest`

## 故障排查

### 服务启动失败

```bash
# 查看服务状态
docker compose ps

# 查看最近日志
docker compose logs --tail=50
```

### 健康检查失败

- 确保后端服务已成功启动
- 检查网络配置是否正确
- 查看服务日志获取详细错误信息

### 前端无法访问后端

- 检查 Next.js 配置中的 `rewrites` 是否正确
- 确保后端服务名 `excel-unprotector-backend` 在网络中可解析
- 检查后端服务是否通过健康检查

### Next.js 构建失败

- 检查依赖安装是否成功
- 确保 Node.js 版本符合要求（>= 22.0.0）
- 查看构建日志获取详细错误信息

## 前端架构说明

### Next.js 配置要点

- **App Router**: 使用 `src/app/` 目录结构
- **Standalone Output**: 启用 `output: "standalone"` 优化 Docker 镜像大小
- **API 代理**: 通过 `rewrites` 配置实现请求转发到后端

### Docker 镜像优化

前端 Dockerfile 使用多阶段构建：

1. **builder 阶段**: 安装依赖并构建应用
2. **runner 阶段**: 仅复制必要文件，使用非特权用户运行

这种方式可以显著减小最终镜像大小。