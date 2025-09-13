const moment = require('moment');

class VariableProcessor {
  constructor(prisma) {
    this.prisma = prisma;
    
    // 配置moment为中文
    moment.locale('zh-cn');
  }

  /**
   * 处理动态变量
   * @param {string} html - 包含变量的HTML
   * @param {Object} context - 上下文数据 
   * @returns {Promise<string>} - 处理后的HTML
   */
  async processVariables(html, context = {}) {
    console.log('🔄 Processing dynamic variables in HTML...');
    
    // 处理所有动态字段
    const fieldRegex = /<span[^>]*class="dynamic-field"[^>]*>(.*?)<\/span>/g;
    
    let processedHtml = html;
    let match;
    
    while ((match = fieldRegex.exec(html)) !== null) {
      const fieldSpan = match[0];
      const fieldContent = match[1];
      
      // 提取字段属性
      const fieldName = this.extractAttribute(fieldSpan, 'data-field-name');
      const fieldType = this.extractAttribute(fieldSpan, 'data-field-type');
      
      console.log(`📊 Processing field: ${fieldName} (${fieldType})`);
      
      let value = '';
      
      switch (fieldType) {
        case 'DATE':
          value = await this.processDateVariable(fieldSpan, context);
          break;
        case 'SYSTEM':
          value = await this.processSystemVariable(fieldSpan, context);
          break;
        case 'DYNAMIC':
          value = await this.processDynamicQuery(fieldSpan, context);
          break;
        case 'FIXED':
          value = this.extractAttribute(fieldSpan, 'data-default-value') || '';
          break;
        default:
          value = fieldContent;
          break;
      }
      
      // 替换字段内容
      processedHtml = processedHtml.replace(fieldSpan, value);
      console.log(`✅ Field ${fieldName} processed: ${value}`);
    }
    
    // 处理动态表格
    processedHtml = await this.processDynamicTables(processedHtml, context);
    
    console.log('✅ Variable processing completed');
    return processedHtml;
  }

  /**
   * 处理日期变量
   */
  async processDateVariable(fieldSpan, context) {
    const dateFormat = this.extractAttribute(fieldSpan, 'data-date-format');
    const now = moment();
    
    switch (dateFormat) {
      case 'YYYY-MM-DD':
        return now.format('YYYY-MM-DD');
      case 'YYYY年MM月DD日':
        return now.format('YYYY年MM月DD日');
      case 'YYYY年M月':
        return now.format('YYYY年M月');
      case 'YYYY年':
        return now.format('YYYY年');
      case 'M月':
        return now.format('M月');
      case 'PREV_MONTH':
        return now.subtract(1, 'month').format('M月');
      case 'NEXT_MONTH':
        return now.add(1, 'month').format('M月');
      default:
        return now.format(dateFormat);
    }
  }

  /**
   * 处理系统变量
   */
  async processSystemVariable(fieldSpan, context) {
    const systemVariable = this.extractAttribute(fieldSpan, 'data-system-variable');
    
    switch (systemVariable) {
      case 'CURRENT_USER':
        return context.currentUser || '系统用户';
      case 'DEPARTMENT':
        return context.department || '信息安全室';
      case 'REPORT_TIME':
        return moment().format('YYYY-MM-DD HH:mm:ss');
      case 'SYSTEM_VERSION':
        return context.systemVersion || 'v1.0.0';
      default:
        return systemVariable;
    }
  }

  /**
   * 处理动态SQL查询
   */
  async processDynamicQuery(fieldSpan, context) {
    try {
      const dataSourceId = this.extractAttribute(fieldSpan, 'data-source-id');
      const sqlQuery = this.extractAttribute(fieldSpan, 'data-sql');
      
      if (!dataSourceId || !sqlQuery) {
        return '查询配置错误';
      }
      
      // 这里应该连接数据源执行查询
      // 暂时返回模拟数据
      const { mockSampleData } = require('../scripts/seed-mock-data');
      const sampleData = mockSampleData['审计数据库'];
      
      if (sampleData && sampleData.length > 0) {
        // 返回第一条记录的第一个字段值
        const firstRecord = sampleData[0];
        const firstValue = Object.values(firstRecord)[0];
        return String(firstValue);
      }
      
      return '无数据';
    } catch (error) {
      console.error('Dynamic query error:', error);
      return '查询错误';
    }
  }

  /**
   * 处理动态表格
   */
  async processDynamicTables(html, context) {
    const tableRegex = /<div[^>]*class="dynamic-table"[^>]*>.*?<\/div>/gs;
    
    let processedHtml = html;
    let match;
    
    while ((match = tableRegex.exec(html)) !== null) {
      const tableDiv = match[0];
      
      // 提取表格属性
      const title = this.extractAttribute(tableDiv, 'data-table-title');
      const dataSourceId = this.extractAttribute(tableDiv, 'data-source-id');
      const sqlQuery = this.extractAttribute(tableDiv, 'data-sql');
      const dataMode = this.extractAttribute(tableDiv, 'data-mode') || 'SQL';
      
      console.log(`📋 Processing table: ${title} (${dataMode})`);
      
      try {
        let tableData = [];
        
        if (dataMode === 'SQL' && dataSourceId && sqlQuery) {
          // 执行SQL查询获取数据
          tableData = await this.executeTableQuery(dataSourceId, sqlQuery);
        } else if (dataMode === 'DATASET') {
          // 从数据集获取数据
          tableData = await this.getDatasetData(dataSourceId, tableDiv);
        }
        
        // 生成表格HTML
        const newTableHtml = this.generateTableHtml(title, tableData, tableDiv);
        processedHtml = processedHtml.replace(tableDiv, newTableHtml);
        
        console.log(`✅ Table ${title} processed with ${tableData.length} rows`);
      } catch (error) {
        console.error('Table processing error:', error);
        const errorHtml = `<div><h4>${title}</h4><p>表格数据加载失败: ${error.message}</p></div>`;
        processedHtml = processedHtml.replace(tableDiv, errorHtml);
      }
    }
    
    return processedHtml;
  }

  /**
   * 执行表格查询
   */
  async executeTableQuery(dataSourceId, sqlQuery) {
    // 这里应该连接真实数据源执行查询
    // 暂时返回模拟数据
    const { mockSampleData } = require('../scripts/seed-mock-data');
    const sampleData = mockSampleData['审计数据库'];
    
    return sampleData || [];
  }

  /**
   * 获取数据集数据
   */
  async getDatasetData(dataSourceId, tableDiv) {
    const columnsAttr = this.extractAttribute(tableDiv, 'data-columns');
    
    if (!columnsAttr) {
      throw new Error('数据集列配置缺失');
    }
    
    const columns = JSON.parse(columnsAttr);
    const { mockSampleData } = require('../scripts/seed-mock-data');
    const sampleData = mockSampleData['审计数据库'];
    
    if (!sampleData) {
      return [];
    }
    
    // 根据列配置过滤和映射数据
    return sampleData.map(row => {
      const mappedRow = {};
      columns.forEach(col => {
        mappedRow[col.title] = row[col.field] || '';
      });
      return mappedRow;
    }).slice(0, 10); // 限制显示条数
  }

  /**
   * 生成表格HTML
   */
  generateTableHtml(title, data, originalDiv) {
    if (!data || data.length === 0) {
      return `<div><h4>${title}</h4><p>暂无数据</p></div>`;
    }
    
    // 获取表头
    const headers = Object.keys(data[0]);
    
    let html = `<div><h4>${title}</h4>`;
    html += '<table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">';
    
    // 表头
    html += '<thead><tr style="background-color: #f5f5f5;">';
    headers.forEach(header => {
      html += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${header}</th>`;
    });
    html += '</tr></thead>';
    
    // 表格数据
    html += '<tbody>';
    data.forEach(row => {
      html += '<tr>';
      headers.forEach(header => {
        html += `<td style="border: 1px solid #ddd; padding: 8px;">${row[header] || ''}</td>`;
      });
      html += '</tr>';
    });
    html += '</tbody></table></div>';
    
    return html;
  }

  /**
   * 提取HTML属性值
   */
  extractAttribute(html, attributeName) {
    const regex = new RegExp(`${attributeName}="([^"]*)"`, 'i');
    const match = html.match(regex);
    return match ? match[1] : '';
  }
}

module.exports = VariableProcessor;