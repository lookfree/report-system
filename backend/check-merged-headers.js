const { PrismaClient } = require('@prisma/client');

async function checkMergedHeaders() {
  const prisma = new PrismaClient();
  
  try {
    console.log('🔍 检查合并表头解析情况...');
    
    // 获取最新的模板
    const latestTemplate = await prisma.rs_report_templates.findFirst({
      orderBy: { createdAt: 'desc' }
    });
    
    if (!latestTemplate) {
      console.log('❌ 没有找到模板');
      return;
    }
    
    console.log(`📝 检查模板: ${latestTemplate.name}`);
    
    // 查看所有配置
    const allConfigs = await prisma.rs_template_configs.findMany({
      where: { 
        templateId: latestTemplate.id
      },
      take: 20,
      select: {
        sectionName: true,
        value: true,
        columnIndex: true,
        dataType: true,
        parentSectionId: true
      },
      orderBy: { columnIndex: 'asc' }
    });
    
    console.log(`\n📋 前20个配置项:`);
    allConfigs.forEach((config, index) => {
      console.log(`  ${index + 1}. 表格: ${config.sectionName}`);
      console.log(`     值: ${config.value}`);
      console.log(`     列索引: ${config.columnIndex}`);
      console.log(`     数据类型: ${config.dataType}`);
      console.log(`     父级ID: ${config.parentSectionId || '无'}`);
      console.log('     ---');
    });
    
    // 检查某个具体表格的完整结构
    const sampleTableName = allConfigs[0]?.sectionName || '';
    const tableConfigs = await prisma.rs_template_configs.findMany({
      where: {
        templateId: latestTemplate.id,
        sectionName: sampleTableName
      },
      select: {
        value: true,
        columnIndex: true,
        dataType: true,
        parentSectionId: true
      },
      orderBy: { columnIndex: 'asc' }
    });
    
    if (tableConfigs.length > 0) {
      console.log(`\n📊 表格 "${sampleTableName}" 的完整结构:`);
      tableConfigs.forEach((config, index) => {
        const prefix = config.parentSectionId ? `  └─ ` : `├─ `;
        const parentInfo = config.parentSectionId ? ` (父ID: ${config.parentSectionId})` : '';
        console.log(`${prefix}列${config.columnIndex}: ${config.value}${parentInfo}`);
      });
    }
    
    // 检查原始表格结构数据
    console.log(`\n🔍 检查原始表格结构数据:`);
    const structure = latestTemplate.structure;
    
    if (structure && structure.sections) {
      console.log(`📋 总共找到 ${structure.sections.length} 个sections`);
      
      const tablesWithStructure = structure.sections.filter(section => section.hasTable);
      console.log(`📊 其中有表格的sections: ${tablesWithStructure.length}`);
      
      if (tablesWithStructure.length > 0) {
        console.log(`\n前3个表格的结构信息:`);
        tablesWithStructure.slice(0, 3).forEach((table, index) => {
          console.log(`\n📄 表格 ${index + 1}: ${table.title}`);
          console.log(`   hasTable: ${table.hasTable}`);
          console.log(`   tableStructure:`, JSON.stringify(table.tableStructure, null, 2));
        });
        
        // 统计不同类型的表格
        const tableTypes = {};
        tablesWithStructure.forEach(table => {
          const type = table.tableStructure?.tableType || 'undefined';
          tableTypes[type] = (tableTypes[type] || 0) + 1;
        });
        
        console.log(`\n📈 表格类型统计:`, tableTypes);
        
        // 查找有合并表头的表格
        const mergedTables = tablesWithStructure.filter(table => 
          table.tableStructure && (
            table.tableStructure.tableType === 'merged' ||
            (table.tableStructure.headers && table.tableStructure.headers.some(h => h.parentHeader))
          )
        );
        
        if (mergedTables.length > 0) {
          console.log(`\n🔗 找到 ${mergedTables.length} 个合并表格`);
          const firstMerged = mergedTables[0];
          console.log(`   示例: ${firstMerged.title}`);
          console.log(`   columnCount: ${firstMerged.tableStructure.columnCount}`);
          console.log(`   headers数量: ${firstMerged.tableStructure.headers?.length || 0}`);
        } else {
          console.log(`❌ 没有找到合并表格类型`);
        }
      }
    } else {
      console.log(`❌ 没有找到structure或sections数据`);
    }
    
  } catch (error) {
    console.error('❌ 检查错误:', error.message);
  } finally {
    await prisma.$disconnect();
  }
}

checkMergedHeaders();