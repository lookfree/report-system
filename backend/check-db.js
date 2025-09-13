const { PrismaClient } = require('@prisma/client');

async function checkDatabase() {
  const prisma = new PrismaClient();
  
  try {
    console.log('🔍 检查数据库状态...');
    
    // 检查模板数量
    const templateCount = await prisma.rs_report_templates.count();
    console.log(`📊 模板数量: ${templateCount}`);
    
    if (templateCount > 0) {
      // 获取最新的模板
      const latestTemplate = await prisma.rs_report_templates.findFirst({
        orderBy: { createdAt: 'desc' }
      });
      console.log(`📝 最新模板: ${latestTemplate.name} (ID: ${latestTemplate.id})`);
      
      // 检查配置数量
      const configCount = await prisma.rs_template_configs.count({
        where: { templateId: latestTemplate.id }
      });
      console.log(`⚙️ 配置数量: ${configCount}`);
      
      // 获取前几个配置示例
      const sampleConfigs = await prisma.rs_template_configs.findMany({
        where: { templateId: latestTemplate.id },
        take: 5,
        select: {
          sectionName: true,
          dataType: true,
          columnIndex: true
        }
      });
      
      console.log('📋 配置示例:');
      sampleConfigs.forEach((config, index) => {
        console.log(`  ${index + 1}. ${config.sectionName} (${config.dataType}, 列${config.columnIndex})`);
      });
    } else {
      console.log('❌ 数据库中没有模板数据');
    }
    
  } catch (error) {
    console.error('❌ 数据库查询错误:', error.message);
  } finally {
    await prisma.$disconnect();
  }
}

checkDatabase();