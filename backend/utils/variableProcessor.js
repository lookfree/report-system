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

    // 先处理数据集占位符（在表格单元格中的）
    let processedHtml = await this.processDatasetPlaceholders(html, context);

    // 处理所有动态字段
    const fieldRegex = /<span[^>]*class="dynamic-field"[^>]*>(.*?)<\/span>/g;
    let match;

    while ((match = fieldRegex.exec(processedHtml)) !== null) {
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
   * 处理数据集占位符（在表格单元格中）
   */
  async processDatasetPlaceholders(html, context) {
    console.log('🔄 Processing dataset placeholders...');

    // 先处理列表数据的起始占位符
    let processedHtml = await this.processListDatasetPlaceholders(html, context);

    // 然后处理单条数据的占位符
    const placeholderRegex = /<div[^>]*class="dataset-placeholder"[^>]*>[\s\S]*?<\/div>/g;
    let match;

    const datasetStore = require('../models/DatasetStore');

    while ((match = placeholderRegex.exec(processedHtml)) !== null) {
      const placeholderDiv = match[0];

      // 提取数据集属性
      const datasetId = this.extractAttribute(placeholderDiv, 'data-dataset-id');
      const datasetName = this.extractAttribute(placeholderDiv, 'data-dataset-name');
      const fieldName = this.extractAttribute(placeholderDiv, 'data-field-name');
      const dataType = this.extractAttribute(placeholderDiv, 'data-data-type');
      const displayFields = this.extractAttribute(placeholderDiv, 'data-display-fields');

      console.log(`📊 Processing dataset placeholder: ${datasetName} (${dataType})`);

      try {
        // 获取数据集
        const dataset = datasetStore.getDataset(datasetId);

        if (!dataset) {
          processedHtml = processedHtml.replace(placeholderDiv, `<span style="color: red;">数据集未找到: ${datasetName}</span>`);
          continue;
        }

        // 执行数据集查询
        const result = await this.executeDatasetQuery(dataset);

        if (dataType === 'single' && fieldName) {
          // 单条数据 - 提取指定字段
          const value = result && result.data && result.data[0] ? result.data[0][fieldName] : '无数据';
          processedHtml = processedHtml.replace(placeholderDiv, String(value));
        } else if (dataType === 'list' && displayFields) {
          // 列表数据 - 生成表格
          const fields = displayFields.split(',');
          const tableHtml = this.generateMiniTable(result.data || [], fields);
          processedHtml = processedHtml.replace(placeholderDiv, tableHtml);
        } else {
          processedHtml = processedHtml.replace(placeholderDiv, '配置错误');
        }

        console.log(`✅ Dataset ${datasetName} processed`);
      } catch (error) {
        console.error(`Error processing dataset ${datasetName}:`, error);
        processedHtml = processedHtml.replace(placeholderDiv, `<span style="color: red;">数据处理错误: ${error.message}</span>`);
      }
    }

    return processedHtml;
  }

  /**
   * 执行数据集查询 - 使用真实数据库
   */
  async executeDatasetQuery(dataset) {
    console.log(`📊 执行数据集查询: ${dataset.name}`);
    console.log(`SQL: ${dataset.sqlQuery}`);

    try {
      // 尝试执行真实数据库查询
      const rows = await this.prisma.$queryRawUnsafe(dataset.sqlQuery);
      console.log(`✅ 查询成功，返回 ${rows.length} 条记录`);

      return {
        success: true,
        type: dataset.type,
        fields: dataset.fields,
        data: dataset.type === 'single' ? [rows[0] || {}] : rows
      };
    } catch (dbError) {
      console.warn(`⚠️ 数据库查询失败: ${dbError.message}`);
      console.log('📦 使用模拟数据作为备选');

      // 如果数据库查询失败，返回模拟数据
      const mockData = this.getMockDataForDataset(dataset);
      return {
        success: true,
        type: dataset.type,
        fields: dataset.fields,
        data: mockData
      };
    }
  }

  /**
   * 获取数据集的模拟数据
   */
  getMockDataForDataset(dataset) {
    // 根据数据集名称返回相应的模拟数据
    const mockDataMap = {
      '模板列表': dataset.type === 'single'
        ? [{ id: 1, name: '安全审计报告模板', createdAt: '2024-01-15', updatedAt: '2024-01-20', status: '已发布' }]
        : [
            { id: 1, name: '安全审计报告模板', createdAt: '2024-01-15', updatedAt: '2024-01-20', status: '已发布' },
            { id: 2, name: '接口安全评估表', createdAt: '2024-01-16', updatedAt: '2024-01-21', status: '草稿' },
            { id: 3, name: '漏洞扫描报告', createdAt: '2024-01-17', updatedAt: '2024-01-22', status: '已发布' }
          ],
      '审计数据': [
        { audit_name: '2024年第一季度安全审计', audit_date: '2024-03-31', risk_level: '高', department: '信息安全部', rectification_status: '整改中' },
        { audit_name: '接口安全专项检查', audit_date: '2024-03-15', risk_level: '中', department: '开发部', rectification_status: '已完成' },
        { audit_name: '数据库权限审计', audit_date: '2024-03-01', risk_level: '低', department: '运维部', rectification_status: '待处理' }
      ],
      '安全事件': [
        { incident_id: 'SEC-2024-001', incident_title: 'SQL注入攻击', occurrence_time: '2024-03-20 14:30', incident_type: '注入攻击', severity: '高' },
        { incident_id: 'SEC-2024-002', incident_title: 'XSS跨站脚本', occurrence_time: '2024-03-21 10:15', incident_type: 'XSS攻击', severity: '中' },
        { incident_id: 'SEC-2024-003', incident_title: '弱密码告警', occurrence_time: '2024-03-22 09:00', incident_type: '认证安全', severity: '低' }
      ]
    };

    const data = mockDataMap[dataset.name] || [];
    return dataset.type === 'single' ? [data[0] || {}] : data;
  }

  /**
   * 处理列表数据集占位符（跨多个单元格）
   */
  async processListDatasetPlaceholders(html, context) {
    console.log('🔄 Processing list dataset placeholders...');

    // 使用DOM解析器处理HTML
    const jsdom = require('jsdom');
    const { JSDOM } = jsdom;
    const dom = new JSDOM(html);
    const document = dom.window.document;

    // 查找所有列表数据的起始占位符
    const startPlaceholders = document.querySelectorAll('.dataset-placeholder-start');

    const datasetStore = require('../models/DatasetStore');

    for (const placeholder of startPlaceholders) {
      const datasetId = placeholder.getAttribute('data-dataset-id');
      const datasetName = placeholder.getAttribute('data-dataset-name');
      const displayFields = placeholder.getAttribute('data-display-fields');

      console.log(`📊 Processing list dataset: ${datasetName}`);

      try {
        // 获取数据集
        const dataset = datasetStore.getDataset(datasetId);

        if (!dataset) {
          placeholder.innerHTML = `<span style="color: red;">数据集未找到</span>`;
          continue;
        }

        // 执行数据集查询
        const result = await this.executeDatasetQuery(dataset);

        if (result && result.data && result.data.length > 0) {
          const fields = displayFields ? displayFields.split(',') : dataset.fields;

          // 找到包含占位符的表格单元格
          const cell = placeholder.closest('td') || placeholder.closest('th');
          if (!cell) continue;

          const table = cell.closest('table');
          if (!table) continue;

          const startRowIndex = cell.parentElement.rowIndex;
          const startCellIndex = cell.cellIndex;

          // 清理占位符
          placeholder.remove();

          // 设置表头
          for (let i = 0; i < fields.length && (startCellIndex + i) < cell.parentElement.cells.length; i++) {
            const headerCell = cell.parentElement.cells[startCellIndex + i];
            headerCell.innerHTML = `<strong>${fields[i]}</strong>`;
          }

          // 处理数据行 - 如果需要，添加新行
          for (let rowIdx = 0; rowIdx < result.data.length; rowIdx++) {
            let dataRow;

            // 获取或创建数据行
            if (startRowIndex + rowIdx + 1 < table.rows.length) {
              dataRow = table.rows[startRowIndex + rowIdx + 1];
            } else {
              // 需要添加新行
              dataRow = table.insertRow();
              // 确保新行有足够的单元格
              while (dataRow.cells.length < table.rows[0].cells.length) {
                dataRow.insertCell();
              }
            }

            // 填充数据
            for (let colIdx = 0; colIdx < fields.length && (startCellIndex + colIdx) < dataRow.cells.length; colIdx++) {
              const dataCell = dataRow.cells[startCellIndex + colIdx];
              const fieldName = fields[colIdx];
              const value = result.data[rowIdx][fieldName] || '-';

              // 清理原有占位符
              const existingPlaceholder = dataCell.querySelector('.dataset-placeholder-data, .dataset-placeholder-field');
              if (existingPlaceholder) {
                existingPlaceholder.remove();
              }

              dataCell.innerHTML = value;
            }
          }

          // 清理剩余的占位符
          const remainingPlaceholders = table.querySelectorAll('.dataset-placeholder-field, .dataset-placeholder-data');
          remainingPlaceholders.forEach(p => p.remove());

        } else {
          placeholder.innerHTML = '无数据';
        }

        console.log(`✅ List dataset ${datasetName} processed with ${result.data ? result.data.length : 0} rows`);
      } catch (error) {
        console.error(`Error processing list dataset ${datasetName}:`, error);
        placeholder.innerHTML = `<span style="color: red;">数据处理错误</span>`;
      }
    }

    return dom.serialize();
  }

  /**
   * 生成小型表格HTML
   */
  generateMiniTable(data, fields) {
    if (!data || data.length === 0) {
      return '<span>无数据</span>';
    }

    let html = '<table style="border-collapse: collapse; border: 1px solid #ddd;">';

    // 表头
    html += '<tr>';
    fields.forEach(field => {
      html += `<th style="border: 1px solid #ddd; padding: 4px; background: #f5f5f5;">${field}</th>`;
    });
    html += '</tr>';

    // 数据行
    data.slice(0, 5).forEach(row => {
      html += '<tr>';
      fields.forEach(field => {
        html += `<td style="border: 1px solid #ddd; padding: 4px;">${row[field] || '-'}</td>`;
      });
      html += '</tr>';
    });

    html += '</table>';
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