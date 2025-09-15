// 数据集存储（内存版本，用于测试）
class DatasetStore {
  constructor() {
    this.datasets = new Map();
    this.nextId = 1;

    // 初始化一个默认的测试数据集
    this.addDataset({
      name: '模板列表',
      description: '获取系统中的报表模板列表',
      type: 'list',
      sqlQuery: 'SELECT id, name, "createdAt" FROM rs_report_templates LIMIT 5',
      fields: ['id', 'name', 'createdAt']
    });
  }

  // 获取所有数据集
  getAllDatasets() {
    return Array.from(this.datasets.values());
  }

  // 获取单个数据集
  getDataset(id) {
    return this.datasets.get(String(id));
  }

  // 添加数据集
  addDataset(dataset) {
    const id = String(this.nextId++);
    const newDataset = {
      id,
      ...dataset,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    this.datasets.set(id, newDataset);
    return newDataset;
  }

  // 更新数据集
  updateDataset(id, updates) {
    const dataset = this.datasets.get(String(id));
    if (!dataset) {
      return null;
    }
    const updatedDataset = {
      ...dataset,
      ...updates,
      updatedAt: new Date().toISOString()
    };
    this.datasets.set(String(id), updatedDataset);
    return updatedDataset;
  }

  // 删除数据集
  deleteDataset(id) {
    return this.datasets.delete(String(id));
  }
}

// 创建单例
const datasetStore = new DatasetStore();

module.exports = datasetStore;