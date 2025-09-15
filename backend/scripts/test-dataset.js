// 测试数据集配置脚本
const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function createTestDataset() {
  try {
    console.log('创建测试数据集配置...');

    // 测试查询1：获取模板列表
    const templates = await prisma.$queryRaw`
      SELECT id, name, "createdAt"
      FROM rs_report_templates
      LIMIT 5
    `;
    console.log('模板列表查询结果:', templates);

    // 测试查询2：获取数据源列表
    const dataSources = await prisma.$queryRaw`
      SELECT id, name, type
      FROM rs_data_sources
      LIMIT 5
    `;
    console.log('数据源列表查询结果:', dataSources);

    // 示例数据集配置
    const sampleDatasetConfigs = [
      {
        name: '模板列表',
        sqlQuery: 'SELECT id, name, "createdAt" FROM rs_report_templates LIMIT 5',
        fields: ['id', 'name', 'createdAt'],
        datasetType: 'list',
        description: '获取最近的5个报表模板'
      },
      {
        name: '系统信息',
        sqlQuery: `SELECT
          '报表系统' as system_name,
          'v1.0.0' as version,
          NOW() as current_time,
          (SELECT COUNT(*) FROM rs_report_templates) as template_count`,
        fields: ['system_name', 'version', 'current_time', 'template_count'],
        datasetType: 'single',
        description: '系统基本信息'
      },
      {
        name: '任务统计',
        sqlQuery: `SELECT
          status,
          COUNT(*) as count
        FROM rs_report_tasks
        GROUP BY status`,
        fields: ['status', 'count'],
        datasetType: 'list',
        description: '任务状态统计'
      }
    ];

    console.log('\n可用的测试数据集配置:');
    sampleDatasetConfigs.forEach((config, index) => {
      console.log(`\n${index + 1}. ${config.name}`);
      console.log(`   描述: ${config.description}`);
      console.log(`   类型: ${config.datasetType}`);
      console.log(`   SQL: ${config.sqlQuery}`);
      console.log(`   字段: ${config.fields.join(', ')}`);
    });

    console.log('\n测试数据集配置创建成功！');
    console.log('您可以在前端使用这些SQL查询来配置数据集。');

  } catch (error) {
    console.error('错误:', error);
  } finally {
    await prisma.$disconnect();
  }
}

createTestDataset();