const { PrismaClient } = require('@prisma/client');

async function initDataSources() {
  const prisma = new PrismaClient();
  
  try {
    console.log('🚀 开始初始化数据源...');
    
    // 示例数据源
    const dataSources = [
      {
        id: 'ds_1',
        name: '生产数据库',
        type: 'postgresql',
        host: 'localhost',
        port: 5432,
        database: 'production_db',
        username: 'admin',
        password: 'password',
        active: true
      },
      {
        id: 'ds_2', 
        name: '日志数据库',
        type: 'postgresql',
        host: 'localhost',
        port: 5433,
        database: 'logs_db',
        username: 'logger',
        password: 'password',
        active: true
      },
      {
        id: 'ds_3',
        name: '审计数据库',
        type: 'postgresql', 
        host: 'localhost',
        port: 5434,
        database: 'audit_db',
        username: 'auditor',
        password: 'password',
        active: true
      }
    ];
    
    // 使用upsert确保数据不重复
    for (const ds of dataSources) {
      const result = await prisma.rs_data_sources.upsert({
        where: { id: ds.id },
        update: ds,
        create: {
          ...ds,
          createdAt: new Date(),
          updatedAt: new Date()
        }
      });
      console.log(`✅ 数据源 "${result.name}" 已添加`);
    }
    
    // 统计总数
    const totalCount = await prisma.rs_data_sources.count({
      where: { active: true }
    });
    console.log(`📊 总共有 ${totalCount} 个可用数据源`);
    
  } catch (error) {
    console.error('❌ 初始化数据源失败:', error.message);
  } finally {
    await prisma.$disconnect();
  }
}

initDataSources();