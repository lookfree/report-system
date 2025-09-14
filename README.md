# 📊 报告管理系统

一个基于 Vue 3 + Node.js 的智能报告管理系统，支持 Word 模板上传、动态配置和自动化报告生成。

## ✨ 核心功能

### 🎯 模板管理
- **Word 模板上传**：支持复杂格式的 Word 文档模板
- **智能表格解析**：自动识别合并表头和复杂表格结构
- **可视化编辑器**：所见即所得的模板编辑体验

### 📝 高级表格编辑
- **表格操作**：支持行/列插入、删除操作
- **单元格合并**：拖拽选择合并、一键拆分
- **尺寸调整**：列宽、行高拖拽调整
- **边框调整**：列间距精细调节
- **右键菜单**：完整的表格操作菜单

### 🔗 数据源集成
- **多数据源支持**：MySQL、PostgreSQL 等
- **动态字段映射**：自动识别数据库字段
- **实时预览**：配置过程中的数据预览
- **SQL 查询支持**：自定义查询语句

### 📋 报告生成
- **自动化生成**：根据模板和数据源生成报告
- **格式保持**：保持原有表格样式和合并关系
- **批量处理**：支持批量报告生成
- **定时任务**：可配置定时生成任务

## 🚀 快速开始

### 环境要求
- Node.js 16.x+
- npm 或 yarn

### 安装步骤

1. **克隆项目**
```bash
git clone https://github.com/lookfree/report-system.git
cd report-system
```

2. **安装依赖**
```bash
# 后端依赖
cd backend
npm install

# 前端依赖
cd ../frontend
npm install
```

3. **启动服务**
```bash
# 启动后端服务 (端口: 8080)
cd backend
npm run dev

# 启动前端服务 (端口: 8081)
cd ../frontend
npm run dev
```

4. **访问应用**
- 前端界面: http://localhost:8081
- 后端 API: http://localhost:8080

## 🏗️ 项目结构

```
report-system/
├── backend/                 # 后端服务
│   ├── routes/             # API 路由
│   ├── services/           # 业务逻辑
│   └── utils/              # 工具函数
├── frontend/               # 前端应用
│   ├── src/
│   │   ├── components/     # Vue 组件
│   │   ├── composables/    # 组合式函数
│   │   ├── views/          # 页面视图
│   │   └── router/         # 路由配置
│   └── public/             # 静态资源
└── README.md
```

## 🎯 核心页面

### 模板管理 (`/templates`)
- 模板列表展示
- 上传新模板
- 模板预览和编辑

### 模板编辑器 (`/templates/:id/editor`)
- 可视化模板编辑
- 表格高级操作
- 实时预览功能

### 数据源管理 (`/datasources`)
- 数据库连接配置
- 连接测试和验证

### 任务管理 (`/tasks`)
- 报告生成任务
- 任务状态监控

## 🔧 技术栈

### 前端
- **Vue 3**: 渐进式 JavaScript 框架
- **Element Plus**: 企业级 UI 组件库
- **Vue Router**: 官方路由管理
- **Vite**: 现代化构建工具

### 后端
- **Node.js**: JavaScript 运行时
- **Express**: Web 应用框架
- **Multer**: 文件上传处理

### 开发工具
- **ESLint**: 代码质量检查
- **Git**: 版本控制

## 📋 使用指南

### 1. 创建模板
1. 准备 Word 文档模板（支持复杂表格）
2. 在模板管理页面上传文档
3. 系统自动解析表格结构

### 2. 配置数据源
1. 添加数据库连接信息
2. 测试连接有效性
3. 查看可用字段

### 3. 编辑模板
1. 进入模板编辑器
2. 使用表格编辑功能
3. 配置数据字段映射

### 4. 生成报告
1. 创建生成任务
2. 选择模板和数据源
3. 执行任务获取报告

## 🎨 表格编辑功能

### 基础操作
- **右键菜单**: 完整的表格操作选项
- **行列管理**: 插入/删除行列
- **单元格编辑**: 直接点击编辑内容

### 高级功能
- **单元格合并**: 拖拽选择 + 右键合并
- **尺寸调整**: 拖拽列边框调整宽度
- **行高调整**: 拖拽行边框调整高度
- **列间距调整**: 精确调节列间距

### 快捷操作
- **拖拽合并**: 按住鼠标拖选单元格
- **边框拖拽**: 悬停边框显示调整手柄
- **右键菜单**: 快速访问所有功能

## 🤝 贡献指南

欢迎贡献代码！请遵循以下步骤：

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 🔗 相关链接

- [项目仓库](https://github.com/lookfree/report-system)
- [问题反馈](https://github.com/lookfree/report-system/issues)

---

**系统已完全支持复杂表格的自动识别、编辑和报告生成功能！**