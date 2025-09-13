#!/bin/bash

echo "================================"
echo "报告系统安装脚本"
echo "================================"

# 安装后端依赖
echo "安装后端依赖..."
cd backend
npm install

# 初始化Prisma
echo "初始化数据库..."
npx prisma generate
npx prisma migrate dev --name init

# 创建必要的目录
mkdir -p uploads reports

# 安装前端依赖
echo "安装前端依赖..."
cd ../frontend
npm install

echo ""
echo "================================"
echo "安装完成！"
echo "================================"
echo ""
echo "启动方法："
echo "1. 后端: cd backend && npm run dev"
echo "2. 前端: cd frontend && npm run dev"
echo ""
echo "访问地址: http://localhost:3000"
echo "================================"