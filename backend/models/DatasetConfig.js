// 数据集配置模型
class DatasetConfig {
  constructor() {
    this.configs = new Map();
  }

  // 添加或更新数据集配置
  addConfig(templateId, cellId, config) {
    const key = `${templateId}_${cellId}`;
    this.configs.set(key, {
      templateId,
      cellId,
      datasetName: config.datasetName,
      datasetType: config.datasetType || 'list', // 'single' or 'list'
      displayType: config.displayType || 'dataset', // 'text' or 'dataset'
      fields: config.fields || [],
      sqlQuery: config.sqlQuery || '',
      dataSourceId: config.dataSourceId,
      createdAt: new Date(),
      updatedAt: new Date()
    });
  }

  // 获取单元格的数据集配置
  getConfig(templateId, cellId) {
    const key = `${templateId}_${cellId}`;
    return this.configs.get(key);
  }

  // 获取模板的所有数据集配置
  getTemplateConfigs(templateId) {
    const configs = [];
    for (const [key, config] of this.configs) {
      if (config.templateId === templateId) {
        configs.push(config);
      }
    }
    return configs;
  }

  // 删除配置
  removeConfig(templateId, cellId) {
    const key = `${templateId}_${cellId}`;
    return this.configs.delete(key);
  }

  // 清除模板的所有配置
  clearTemplateConfigs(templateId) {
    const keysToDelete = [];
    for (const [key, config] of this.configs) {
      if (config.templateId === templateId) {
        keysToDelete.push(key);
      }
    }
    keysToDelete.forEach(key => this.configs.delete(key));
  }
}

// 创建单例实例
const datasetConfig = new DatasetConfig();

module.exports = datasetConfig;