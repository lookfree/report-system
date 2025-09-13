const { PrismaClient } = require('@prisma/client');
const { v4: uuidv4 } = require('uuid');

const prisma = new PrismaClient();

const mockDataSources = [
  {
    id: uuidv4(),
    name: '用户权限数据源',
    type: 'postgresql',
    host: 'localhost',
    port: 5432,
    database: 'security_audit',
    username: 'postgres',
    password: 'password',
    active: true,
    updatedAt: new Date()
  },
  {
    id: uuidv4(),
    name: '系统日志数据源',
    type: 'postgresql', 
    host: 'log-server.internal',
    port: 5432,
    database: 'system_logs',
    username: 'log_reader',
    password: 'log_pass',
    active: true,
    updatedAt: new Date()
  },
  {
    id: uuidv4(),
    name: '业务数据源',
    type: 'postgresql',
    host: 'business-db.company.com',
    port: 5432,
    database: 'business_data',
    username: 'business_user',
    password: 'business_pass',
    active: true,
    updatedAt: new Date()
  },
  {
    id: uuidv4(),
    name: '财务数据源',
    type: 'postgresql',
    host: 'finance-db.internal',
    port: 5432,
    database: 'finance_system',
    username: 'finance_reader',
    password: 'finance_pass',
    active: true,
    updatedAt: new Date()
  },
  {
    id: uuidv4(),
    name: '人事数据源',
    type: 'postgresql',
    host: 'hr-system.company.com',
    port: 5432,
    database: 'hr_database',
    username: 'hr_user',
    password: 'hr_pass',
    active: true,
    updatedAt: new Date()
  }
];

// Mock sample data that would be returned by these data sources
const mockSampleData = {
  '用户权限数据源': [
    {
      data_batch: '2024-12',
      org_code: 'ORG001', 
      org_name: '总公司',
      parent_org_code: '',
      person_name: '张三',
      account_type: '管理员账户',
      phone: '13800138000',
      entry_time: '2024-01-15'
    },
    {
      data_batch: '2024-12',
      org_code: 'ORG002',
      org_name: '分公司A',
      parent_org_code: 'ORG001',
      person_name: '李四',
      account_type: '普通用户',
      phone: '13900139000',
      entry_time: '2024-03-10'
    },
    {
      data_batch: '2024-12',
      org_code: 'ORG003',
      org_name: '部门B',
      parent_org_code: 'ORG002',
      person_name: '王五',
      account_type: '部门主管',
      phone: '13700137000',
      entry_time: '2024-02-20'
    },
    {
      data_batch: '2024-12',
      org_code: 'ORG004',
      org_name: '技术部',
      parent_org_code: 'ORG001',
      person_name: '赵六',
      account_type: '技术人员',
      phone: '13600136000',
      entry_time: '2024-04-05'
    },
    {
      data_batch: '2024-12',
      org_code: 'ORG005',
      org_name: '财务部',
      parent_org_code: 'ORG001',
      person_name: '孙七',
      account_type: '财务人员',
      phone: '13500135000',
      entry_time: '2024-01-30'
    }
  ],
  '系统日志数据源': [
    {
      log_id: 'LOG001',
      log_time: '2024-12-10 08:30:00',
      log_level: 'INFO',
      module: '用户登录',
      message: '用户张三成功登录系统',
      ip_address: '192.168.1.100',
      user_agent: 'Mozilla/5.0...'
    },
    {
      log_id: 'LOG002', 
      log_time: '2024-12-10 09:15:00',
      log_level: 'WARN',
      module: '权限验证',
      message: '用户李四尝试访问无权限页面',
      ip_address: '192.168.1.101',
      user_agent: 'Chrome/118.0...'
    },
    {
      log_id: 'LOG003',
      log_time: '2024-12-10 10:45:00', 
      log_level: 'ERROR',
      module: '数据库连接',
      message: '数据库连接超时',
      ip_address: '192.168.1.10',
      user_agent: 'System'
    }
  ],
  '业务数据源': [
    {
      business_id: 'BUS001',
      customer_name: '阿里巴巴集团',
      contract_amount: 1500000.00,
      contract_date: '2024-01-15',
      status: '执行中',
      manager: '张三',
      department: '销售部'
    },
    {
      business_id: 'BUS002',
      customer_name: '腾讯科技',
      contract_amount: 2800000.00,
      contract_date: '2024-03-20',
      status: '已完成',
      manager: '李四',
      department: '商务部'
    }
  ],
  '财务数据源': [
    {
      account_code: 'ACC001',
      account_name: '银行存款',
      balance: 5680000.00,
      currency: 'CNY',
      last_update: '2024-12-09',
      account_type: '资产'
    },
    {
      account_code: 'ACC002',
      account_name: '应收账款',
      balance: 2340000.00,
      currency: 'CNY', 
      last_update: '2024-12-09',
      account_type: '资产'
    }
  ],
  '人事数据源': [
    {
      employee_id: 'EMP001',
      employee_name: '张三',
      department: '技术部',
      position: '高级工程师',
      hire_date: '2020-01-15',
      salary: 25000.00,
      status: '在职'
    },
    {
      employee_id: 'EMP002',
      employee_name: '李四',
      department: '销售部',
      position: '销售经理',
      hire_date: '2019-06-01',
      salary: 22000.00,
      status: '在职'
    }
  ]
};

async function seedMockData() {
  try {
    console.log('🌱 开始插入Mock数据源...');
    
    // Check if mock data already exists
    const existingCount = await prisma.rs_data_sources.count();
    console.log(`📊 当前数据源数量: ${existingCount}`);
    
    if (existingCount > 0) {
      console.log('⚠️ 数据源表中已存在数据，清空现有数据...');
      await prisma.rs_data_sources.deleteMany({});
      console.log('✅ 现有数据已清空');
    }
    
    // Insert mock data sources
    const createdSources = await prisma.rs_data_sources.createMany({
      data: mockDataSources
    });
    
    console.log(`✅ 成功插入 ${createdSources.count} 个Mock数据源`);
    
    // Display created data sources
    const allSources = await prisma.rs_data_sources.findMany({
      orderBy: { name: 'asc' }
    });
    
    console.log('\n📋 创建的数据源列表:');
    allSources.forEach((source, index) => {
      console.log(`  ${index + 1}. ${source.name}`);
      console.log(`     - 类型: ${source.type}`);
      console.log(`     - 主机: ${source.host}:${source.port}`);
      console.log(`     - 数据库: ${source.database}`);
      console.log(`     - 状态: ${source.active ? '活跃' : '非活跃'}`);
      console.log('');
    });
    
    console.log('🎉 Mock数据源插入完成！');
    
    // Also log sample data that would be returned
    console.log('\n📄 样本数据预览 (仅用于展示，不插入数据库):');
    Object.entries(mockSampleData).forEach(([sourceName, samples]) => {
      console.log(`\n${sourceName}:`);
      console.log(JSON.stringify(samples.slice(0, 2), null, 2));
    });
    
    return allSources;
  } catch (error) {
    console.error('❌ Mock数据插入失败:', error);
    throw error;
  }
}

// Export both the function and the sample data for use in API
module.exports = {
  seedMockData,
  mockSampleData,
  mockDataSources
};

// If run directly
if (require.main === module) {
  seedMockData()
    .then(() => {
      console.log('✨ 脚本执行完成');
      process.exit(0);
    })
    .catch((error) => {
      console.error('💥 脚本执行失败:', error);
      process.exit(1);
    })
    .finally(async () => {
      await prisma.$disconnect();
    });
}