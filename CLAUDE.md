# 报表系统配置说明

## 端口配置
- **后端服务**: PORT=8080 (backend目录)
- **前端服务**: PORT=8081 (frontend目录)

## 启动命令

### 后端启动
```bash
cd backend
PORT=8080 npm start
```

### 前端启动
```bash
cd frontend
npm run dev -- --port 8081
```

## 项目结构
```
report-system/
├── backend/       # Express后端服务 (端口8080)
├── frontend/      # Vue3前端应用 (端口8081)
├── test-*.html   # 测试文件
└── CLAUDE.md     # 项目配置说明
```

## 功能特性
- 动态数据刷新：导出/预览时实时从数据库获取数据
- 数据集配置：支持配置SQL查询数据集
- 表格单元格数据集插入：双击单元格可插入数据集
- Word文档生成：支持动态数据替换

## 注意事项
- 确保端口8080和8081未被占用
- 后端依赖PostgreSQL数据库
- 前端通过/api代理访问后端服务