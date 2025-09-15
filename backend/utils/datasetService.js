const { PrismaClient } = require('@prisma/client');
const datasetConfig = require('../models/DatasetConfig');

class DatasetService {
  constructor() {
    // 使用 Prisma 客户端
    this.prisma = new PrismaClient();
  }

  // 执行数据集查询
  async executeDatasetQuery(templateId, cellId) {
    try {
      const config = datasetConfig.getConfig(templateId, cellId);
      if (!config) {
        return null;
      }

      // 如果是静态文本，直接返回
      if (config.displayType === 'text') {
        return {
          type: 'text',
          value: config.staticText || ''
        };
      }

      // 执行SQL查询 - 使用 Prisma 的原始查询
      // 注意：这里使用 $queryRawUnsafe 来执行动态SQL
      // 在生产环境中应该添加SQL注入防护
      const rows = await this.prisma.$queryRawUnsafe(config.sqlQuery);

      if (config.datasetType === 'single') {
        // 单条数据
        return {
          type: 'single',
          data: rows[0] || {},
          fields: config.fields
        };
      } else {
        // 列表数据
        return {
          type: 'list',
          data: rows || [],
          fields: config.fields
        };
      }
    } catch (error) {
      console.error('Dataset query error:', error);
      return {
        type: 'error',
        message: error.message
      };
    }
  }

  // 获取模板的所有数据集数据
  async getTemplateDatasets(templateId) {
    const configs = datasetConfig.getTemplateConfigs(templateId);
    const datasets = {};

    for (const config of configs) {
      const data = await this.executeDatasetQuery(templateId, config.cellId);
      datasets[config.cellId] = data;
    }

    return datasets;
  }

  // 将数据集数据填充到Excel单元格
  fillCellWithDataset(worksheet, cellAddress, datasetResult) {
    if (!datasetResult) return;

    const cell = worksheet.getCell(cellAddress);

    switch (datasetResult.type) {
      case 'text':
        cell.value = datasetResult.value;
        break;

      case 'single':
        // 单条数据，将字段值用换行分隔
        const values = datasetResult.fields.map(field =>
          `${field}: ${datasetResult.data[field] || ''}`
        );
        cell.value = values.join('\n');
        break;

      case 'list':
        // 列表数据，创建表格
        const startRow = cell.row;
        const startCol = cell.col;

        // 添加表头
        datasetResult.fields.forEach((field, index) => {
          worksheet.getCell(startRow, startCol + index).value = field;
        });

        // 添加数据行
        datasetResult.data.forEach((row, rowIndex) => {
          datasetResult.fields.forEach((field, colIndex) => {
            worksheet.getCell(startRow + rowIndex + 1, startCol + colIndex).value = row[field] || '';
          });
        });
        break;

      case 'error':
        cell.value = `Error: ${datasetResult.message}`;
        break;
    }
  }

  // 清理连接
  async closeConnection() {
    try {
      await this.prisma.$disconnect();
    } catch (error) {
      console.error('Error closing Prisma connection:', error);
    }
  }
}

// 创建单例实例
const datasetService = new DatasetService();

module.exports = datasetService;