const mammoth = require('mammoth');
const CoverDetector = require('./coverDetector');

class UniversalWordParser {
  constructor() {
    this.config = {
      // 支持的标题级别
      headerLevels: [1, 2, 3, 4, 5, 6],
      // 表格检测策略
      tableStrategies: ['auto', 'first-two-rows', 'dynamic'],
      // 合并表头检测
      mergeDetection: ['explicit-colspan', 'implicit-count', 'content-analysis']
    };
  }

  async parseTemplate(filePath, options = {}) {
    try {
      console.log('Starting Word template parsing...');
      
      // 使用自定义样式映射来保留更多格式信息
      const styleMap = [
        "p[style-name='Title'] => h1.document-title",
        "p[style-name='Heading 1'] => h1",
        "p[style-name='Heading 2'] => h2",
        "p[style-name='Heading 3'] => h3",
        "p[style-name='Normal'] => p.normal-text",
        "b => b",
        "i => i",
        "u => u",
        "table => table.word-table"
      ];
      
      const result = await mammoth.convertToHtml({
        path: filePath,
        styleMap: styleMap,
        convertImage: mammoth.images.imgElement(function(image) {
          return image.read("base64").then(function(imageBuffer) {
            return {
              src: "data:" + image.contentType + ";base64," + imageBuffer
            };
          });
        })
      });
      
      let html = result.value;
      
      // 应用中文字体样式增强
      html = this.enhanceChineseFontStyles(html);
      
      console.log('Raw HTML preview:', html.substring(0, 500));
      console.log('HTML size:', html.length, 'characters');
      
      // 使用动态策略解析
      const structure = await this.extractStructure(html, options);
      
      // 保存原始HTML内容用于预览
      structure.originalHtml = html;
      
      console.log('Template parsing completed successfully');
      return structure;
    } catch (error) {
      console.error('Error parsing Word template:', error);
      throw error;
    }
  }
  
  async extractStructure(html, options = {}) {
    const structure = {
      title: '',
      sections: [],
      metadata: {
        parseStrategy: options.strategy || 'auto',
        parsedAt: new Date().toISOString(),
        totalTables: 0,
        totalSections: 0
      }
    };
    
    // 1. 提取标题 - 支持多种格式
    structure.title = this.extractTitle(html);
    
    // 2. 提取章节 - 动态识别
    const sections = await this.extractSections(html, options);
    
    // 3. 确保sections是数组
    if (!Array.isArray(sections)) {
      console.error('extractSections returned non-array:', typeof sections, sections);
      throw new Error(`extractSections returned ${typeof sections} instead of array`);
    }
    
    structure.sections = sections;
    
    // 4. 更新元数据
    structure.metadata.totalSections = sections.length;
    structure.metadata.totalTables = sections.filter(s => s.hasTable).length;
    
    return structure;
  }
  
  extractTitle(html) {
    // 多种标题提取策略
    const strategies = [
      // 策略1: 标准H1标签
      () => {
        const match = html.match(/<h1[^>]*>(.*?)<\/h1>/i);
        return match ? this.cleanHtml(match[1]) : null;
      },
      // 策略2: 文档开头的粗体大字
      () => {
        const match = html.match(/<p[^>]*>\s*<strong[^>]*>(.*?)<\/strong>\s*<\/p>/i);
        if (match && match[1].length < 100) {
          return this.cleanHtml(match[1]);
        }
        return null;
      },
      // 策略3: 第一个非空段落（如果较短）
      () => {
        const match = html.match(/<p[^>]*>(.*?)<\/p>/i);
        if (match) {
          const text = this.cleanHtml(match[1]).trim();
          if (text.length > 5 && text.length < 150) {
            return text;
          }
        }
        return null;
      }
    ];
    
    for (const strategy of strategies) {
      const title = strategy();
      if (title) return title;
    }
    
    return '未知文档标题';
  }
  
  async extractSections(html, options = {}) {
    const sections = [];
    
    // 首先找到所有表格
    const allTables = [...html.matchAll(/<table[^>]*>.*?<\/table>/gis)];
    console.log(`Found ${allTables.length} tables in document`);
    
    // 动态识别标题模式
    const headerPattern = this.detectHeaderPattern(html);
    console.log('Detected header pattern:', headerPattern);
    
    const headerMatches = [...html.matchAll(headerPattern.regex)];
    
    // 如果找到标题且表格数量不多，按标题分割sections
    if (headerMatches.length > 0 && allTables.length < 50) {
      for (let index = 0; index < headerMatches.length; index++) {
        const match = headerMatches[index];
        const level = this.extractHeaderLevel(match, headerPattern);
        const text = this.cleanHtml(match[2] || match[1]);
        
        // 提取内容范围
        const startPos = match.index + match[0].length;
        const endPos = headerMatches[index + 1] ? headerMatches[index + 1].index : html.length;
        const content = html.substring(startPos, endPos);
        
        // 多策略表格解析（异步）
        const tableStructure = await this.extractTableStructureAdvanced(content, options);

        // 生成分节类型标识
        const sectionType = this.inferSectionType(text, tableStructure);

        // 提取静态字段（占位符）
        const fields = this.extractStaticFields(content);
        
        const section = {
          id: `section_${index}`,
          level: level,
          title: text,
          hasTable: tableStructure.hasTable,
          hasContent: content.trim().length > 0,
          contentPreview: this.cleanHtml(content).substring(0, 100),
          tableStructure: tableStructure,
          rawContent: content.substring(0, 200),
          originalHtml: content,
          type: sectionType,
          fields: fields
        };
        
        sections.push(section);
        
        // 每10个section添加一个异步yield避免阻塞
        if (index % 10 === 0) {
          await this.asyncYield();
        }
      }
    } else {
      // 如果没有找到标题，批量处理每个表格
      console.log('No headers found, creating sections for each table with batch processing');
      
      const BATCH_SIZE = 5; // 每批处理5个表格
      const totalBatches = Math.ceil(allTables.length / BATCH_SIZE);
      
      for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
        const startIndex = batchIndex * BATCH_SIZE;
        const endIndex = Math.min(startIndex + BATCH_SIZE, allTables.length);
        const batchTables = allTables.slice(startIndex, endIndex);
        
        console.log(`Processing batch ${batchIndex + 1}/${totalBatches}, tables ${startIndex + 1}-${endIndex}`);
        
        // 并行处理批次内的表格
        const batchPromises = batchTables.map(async (tableMatch, batchTableIndex) => {
          const globalIndex = startIndex + batchTableIndex;
          return await this.processTableWithTimeout(tableMatch, globalIndex, html, options);
        });
        
        try {
          const batchSections = await Promise.all(batchPromises);
          sections.push(...batchSections.filter(section => section !== null));
        } catch (error) {
          console.error(`Batch ${batchIndex + 1} processing failed:`, error.message);
          // 继续处理下一批次
        }
        
        // 批次间暂停，避免内存压力
        await this.asyncYield();
        
        // 记录进度
        const progress = Math.round((endIndex / allTables.length) * 100);
        console.log(`Progress: ${progress}% (${endIndex}/${allTables.length})`);
      }
    }
    
    // 如果还是没有找到任何sections，创建fallback
    if (sections.length === 0) {
      sections.push(this.createFallbackSection(html));
    }
    
    console.log(`Successfully processed ${sections.length} sections`);
    return sections;
  }
  
  detectHeaderPattern(html) {
    // 检测文档使用的标题模式
    const patterns = [
      // 标准HTML标题
      {
        name: 'html-headers',
        regex: /<h([1-6])[^>]*>(.*?)<\/h\1>/gi,
        levelExtractor: (match) => parseInt(match[1])
      },
      // 粗体数字标题 (如 "1. 章节名")
      {
        name: 'numbered-bold',
        regex: /<p[^>]*>\s*<strong[^>]*>\s*(\d+\.?\s*)(.*?)<\/strong>/gi,
        levelExtractor: (match) => 1
      },
      // 粗体文本作为标题
      {
        name: 'bold-headers',
        regex: /<p[^>]*>\s*<strong[^>]*>(.*?)<\/strong>\s*<\/p>/gi,
        levelExtractor: (match) => 2
      },
      // 特殊格式：● 开头的标题
      {
        name: 'bullet-headers',
        regex: /<p[^>]*>\s*[●•]\s*(.*?)<\/p>/gi,
        levelExtractor: (match) => 3
      }
    ];
    
    // 统计每种模式的匹配数量
    let bestPattern = patterns[0];
    let maxMatches = 0;
    
    for (const pattern of patterns) {
      const matches = [...html.matchAll(pattern.regex)];
      if (matches.length > maxMatches) {
        maxMatches = matches.length;
        bestPattern = pattern;
      }
    }
    
    return bestPattern;
  }
  
  extractHeaderLevel(match, pattern) {
    try {
      return pattern.levelExtractor(match);
    } catch (error) {
      return 1; // 默认级别
    }
  }

  // 根据标题与是否包含表格推断分节类型
  inferSectionType(title, tableStructure) {
    try {
      const hasTable = !!(tableStructure && tableStructure.hasTable);
      const t = (title || '').toString();
      const isDetail = /明细|细则/i.test(t);
      const isSummary = /汇总|合计|总计/i.test(t);
      if (hasTable && isDetail) return 'detail';
      if (hasTable && isSummary) return 'summary';
      return 'static';
    } catch (e) {
      return 'static';
    }
  }

  // 提取静态字段占位符，返回结构化字段定义
  extractStaticFields(sectionHtml) {
    try {
      if (!sectionHtml) return [];
      const placeholderRegex = /\{\{\s*([A-Z0-9_\-]+)\s*\}\}/g;
      const unique = new Set();
      const fields = [];
      let match;
      while ((match = placeholderRegex.exec(sectionHtml)) !== null) {
        const originalId = match[1] || '';
        const id = originalId.toLowerCase();
        if (id && !unique.has(id)) {
          unique.add(id);
          let inputType = 'text';
          const lower = id.toLowerCase();
          if (/date|riqi|日期/.test(lower)) inputType = 'date';
          if (/summary|zongjie|总结|content|描述|说明/.test(lower)) inputType = 'textarea';
          fields.push({
            id,
            label: id.replace(/_/g, ' ').toUpperCase(),
            inputType,
            defaultValue: '',
            anchor: { placeholder: `{{${originalId}}}` }
          });
        }
      }
      return fields;
    } catch (e) {
      return [];
    }
  }
  
  async extractTableStructureAdvanced(html, options = {}) {
    const tableStructure = {
      hasTable: false,
      headers: [],
      rowCount: 0,
      columnCount: 0,
      tableType: 'simple',
      parseStrategy: 'none',
      allTables: []
    };
    
    // 查找所有表格
    const tableMatches = html.match(/<table[^>]*>.*?<\/table>/gis);
    
    if (!tableMatches || tableMatches.length === 0) {
      return tableStructure;
    }
    
    tableStructure.hasTable = true;
    tableStructure.parseStrategy = options.tableStrategy || 'auto';
    
    // 异步解析所有表格，带超时保护
    for (let tableIndex = 0; tableIndex < tableMatches.length; tableIndex++) {
      try {
        const tableHtml = tableMatches[tableIndex];
        const tableInfo = await this.parseIndividualTableWithTimeout(tableHtml, tableIndex, options);
        if (tableInfo) {
          tableStructure.allTables.push(tableInfo);
        }
      } catch (error) {
        console.warn(`Failed to parse table ${tableIndex}:`, error.message);
      }
    }
    
    // 选择最佳表格作为主表格
    const mainTable = this.selectMainTable(tableStructure.allTables);
    if (mainTable) {
      Object.assign(tableStructure, mainTable);
    }
    
    return tableStructure;
  }
  
  parseIndividualTable(tableHtml, tableIndex, options = {}) {
    const allRowMatches = [...tableHtml.matchAll(/<tr[^>]*>(.*?)<\/tr>/gis)];
    
    if (allRowMatches.length === 0) {
      return { hasTable: false, headers: [], rowCount: 0, columnCount: 0 };
    }
    
    // 多策略表头识别
    const headerStrategies = [
      () => this.analyzeTableStructureStandard(allRowMatches),
      () => this.analyzeTableStructureByContent(allRowMatches),
      () => this.analyzeTableStructureByFormat(allRowMatches)
    ];
    
    let bestResult = null;
    let maxConfidence = 0;
    
    for (const strategy of headerStrategies) {
      try {
        const result = strategy();
        if (result && result.confidence > maxConfidence) {
          maxConfidence = result.confidence;
          bestResult = result;
        }
      } catch (error) {
        console.warn('Table analysis strategy failed:', error.message);
      }
    }
    
    return bestResult || {
      hasTable: true,
      headers: [],
      rowCount: allRowMatches.length,
      columnCount: 0,
      tableType: 'unknown',
      tableIndex: tableIndex,
      confidence: 0
    };
  }
  
  analyzeTableStructureStandard(rowMatches) {
    // 使用现有的标准分析逻辑
    const structure = {
      headers: [],
      columnCount: 0,
      tableType: 'simple',
      confidence: 0.5
    };
    
    // 解析行结构
    const rows = rowMatches.map(match => {
      const cellMatches = [...match[1].matchAll(/<t[hd][^>]*>(.*?)<\/t[hd]>/gis)];
      return cellMatches.map(cell => {
        const cellHtml = cell[0];
        // 提取colspan和rowspan属性
        const colspanMatch = cellHtml.match(/colspan=["']?(\d+)["']?/i);
        const rowspanMatch = cellHtml.match(/rowspan=["']?(\d+)["']?/i);
        
        return {
          content: this.cleanHtml(cell[1]).trim(),
          colspan: colspanMatch ? parseInt(colspanMatch[1]) : 1,
          rowspan: rowspanMatch ? parseInt(rowspanMatch[1]) : 1,
          isHeader: cellHtml.includes('<th'),
          rawHtml: cellHtml
        };
      });
    });
    
    if (rows.length === 0) return structure;
    
    // 检测合并表头
    const hasImplicitMerge = rows.length > 1 && rows[0].length < rows[1].length;
    const hasExplicitMerge = rows[0] && rows[0].some(cell => cell.colspan > 1);
    
    if (hasImplicitMerge || hasExplicitMerge) {
      structure.tableType = 'merged';
      structure.headers = this.extractMergedHeaders(rows);
      structure.confidence = 0.8;
    } else {
      structure.tableType = 'simple';
      structure.headers = this.extractSimpleHeaders(rows[0] || []);
      structure.confidence = 0.7;
    }
    
    // 计算实际列数，考虑合并单元格
    structure.columnCount = this.calculateActualColumnCount(rows);
    return structure;
  }
  
  // 计算表格的实际列数，考虑合并单元格
  calculateActualColumnCount(rows) {
    if (!rows || rows.length === 0) return 0;
    
    let maxColumnCount = 0;
    
    // 遍历每一行，计算实际占用的列数
    for (const row of rows) {
      let currentRowColumnCount = 0;
      for (const cell of row) {
        currentRowColumnCount += cell.colspan || 1;
      }
      maxColumnCount = Math.max(maxColumnCount, currentRowColumnCount);
    }
    
    return maxColumnCount;
  }
  
  analyzeTableStructureByContent(rowMatches) {
    // 基于内容分析表头
    const structure = {
      headers: [],
      columnCount: 0,
      tableType: 'content-based',
      confidence: 0.3
    };
    
    // TODO: 实现基于内容的表头识别
    // 例如：识别数字、日期、文本模式来推断表头
    
    return structure;
  }
  
  analyzeTableStructureByFormat(rowMatches) {
    // 基于格式分析表头
    const structure = {
      headers: [],
      columnCount: 0,
      tableType: 'format-based',
      confidence: 0.4
    };
    
    // TODO: 实现基于格式的表头识别
    // 例如：粗体、居中、背景色等
    
    return structure;
  }
  
  selectMainTable(allTables) {
    if (allTables.length === 0) return null;
    
    // 选择置信度最高的表格作为主表格
    return allTables.reduce((best, current) => {
      const currentScore = (current.confidence || 0) * (current.columnCount || 0);
      const bestScore = (best.confidence || 0) * (best.columnCount || 0);
      return currentScore > bestScore ? current : best;
    });
  }
  
  extractNearbyTitle(text) {
    // 从表格前面的文本中提取可能的标题
    const strategies = [
      // 特殊符号开头的标题（优先级最高）
      () => {
        const bulletMatch = text.match(/[●•]\s*([^<\n]+)/);
        if (bulletMatch) {
          const title = bulletMatch[1].trim();
          if (title.length > 2 && title.length < 150) {
            console.log('Found bullet title:', title);
            return title;
          }
        }
        return null;
      },
      // 段落中包含特殊符号的文本
      () => {
        const match = text.match(/<p[^>]*>.*?[●•]\s*([^<]+).*?<\/p>/);
        if (match) {
          const title = this.cleanHtml(match[1]).trim();
          if (title.length > 2 && title.length < 150) {
            console.log('Found paragraph bullet title:', title);
            return title;
          }
        }
        return null;
      },
      // 粗体文本（取最后一个）
      () => {
        const matches = [...text.matchAll(/<strong[^>]*>(.*?)<\/strong>/g)];
        if (matches.length > 0) {
          const lastMatch = matches[matches.length - 1];
          const title = this.cleanHtml(lastMatch[1]).trim();
          if (title.length > 2 && title.length < 150) {
            console.log('Found bold title:', title);
            return title;
          }
        }
        return null;
      },
      // 段落中的短文本（取最后一个段落）
      () => {
        const matches = [...text.matchAll(/<p[^>]*>(.*?)<\/p>/g)];
        if (matches.length > 0) {
          const lastMatch = matches[matches.length - 1];
          const cleanText = this.cleanHtml(lastMatch[1]).trim();
          if (cleanText.length > 2 && cleanText.length < 150 && !cleanText.includes('表格')) {
            console.log('Found paragraph title:', cleanText);
            return cleanText;
          }
        }
        return null;
      },
      // 数字标题 "1." "1.1" 等
      () => {
        const match = text.match(/(\d+(?:\.\d+)*\.?\s*[^<\n]{3,100})/);
        if (match) {
          const title = match[1].trim();
          console.log('Found numbered title:', title);
          return title;
        }
        return null;
      }
    ];
    
    for (const strategy of strategies) {
      const title = strategy();
      if (title) return title;
    }
    
    return null;
  }

  createFallbackSection(html) {
    // 当无法识别标题时的后备方案
    const tableStructure = this.extractTableStructureAdvanced(html);
    const fields = this.extractStaticFields(html);
    return {
      id: 'section_0',
      level: 1,
      title: '文档内容',
      hasTable: tableStructure.hasTable,
      hasContent: true,
      contentPreview: this.cleanHtml(html).substring(0, 100),
      tableStructure: tableStructure,
      originalHtml: html,
      type: 'static',
      fields: fields
    };
  }
  
  // 复用现有的方法
  extractMergedHeaders(rows) {
    // ... 复用现有的合并表头逻辑
    const headers = [];
    
    if (rows.length < 2) {
      return this.extractSimpleHeaders(rows[0] || []);
    }
    
    const firstRow = rows[0];
    const secondRow = rows[1];
    
    const isImplicitMerge = firstRow.length < secondRow.length;
    let headerIndex = 0;
    let secondRowIndex = 0;
    
    if (isImplicitMerge) {
      for (let i = 0; i < firstRow.length; i++) {
        const firstCell = firstRow[i];
        
        if (i === firstRow.length - 1) {
          const parentName = firstCell.content || `合并列${i + 1}`;
          const remainingSecondRowCells = secondRow.length - secondRowIndex;
          
          for (let j = 0; j < remainingSecondRowCells; j++) {
            const secondCell = secondRow[secondRowIndex + j];
            const subColumnName = secondCell.content || `${parentName}子列${j + 1}`;
            
            headers.push({
              index: headerIndex++,
              name: subColumnName,
              originalName: subColumnName,
              parentHeader: parentName,
              dataType: 'FIXED',
              value: '',
              sqlQuery: '',
              dataSourceId: null
            });
          }
          secondRowIndex += remainingSecondRowCells;
        } else {
          const columnName = firstCell.content || `列${i + 1}`;
          headers.push({
            index: headerIndex++,
            name: columnName,
            originalName: columnName,
            parentHeader: null,
            dataType: 'FIXED',
            value: '',
            sqlQuery: '',
            dataSourceId: null
          });
        }
      }
    } else {
      // 标准合并处理
      for (let i = 0; i < firstRow.length; i++) {
        const firstCell = firstRow[i];
        
        if (firstCell.colspan > 1) {
          const parentName = firstCell.content || `合并列${i + 1}`;
          
          for (let j = 0; j < firstCell.colspan; j++) {
            let subColumnName = '';
            
            if (secondRowIndex < secondRow.length) {
              const secondCell = secondRow[secondRowIndex];
              subColumnName = secondCell.content || `${parentName}子列${j + 1}`;
              secondRowIndex++;
            } else {
              subColumnName = `${parentName}子列${j + 1}`;
            }
            
            headers.push({
              index: headerIndex++,
              name: subColumnName,
              originalName: subColumnName,
              parentHeader: parentName,
              dataType: 'FIXED',
              value: '',
              sqlQuery: '',
              dataSourceId: null
            });
          }
        } else {
          const columnName = firstCell.content || `列${i + 1}`;
          headers.push({
            index: headerIndex++,
            name: columnName,
            originalName: columnName,
            parentHeader: null,
            dataType: 'FIXED',
            value: '',
            sqlQuery: '',
            dataSourceId: null
          });
          
          if (secondRowIndex < secondRow.length) {
            secondRowIndex++;
          }
        }
      }
    }
    
    return headers;
  }
  
  extractSimpleHeaders(firstRow) {
    return firstRow.map((cell, index) => ({
      index: index,
      name: cell.content || `列${index + 1}`,
      originalName: cell.content,
      parentHeader: null,
      dataType: 'FIXED',
      value: '',
      sqlQuery: '',
      dataSourceId: null
    }));
  }
  
  cleanHtml(html) {
    return html
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .trim();
  }
  
  // 简单的表格样式增强 - 暂时禁用以避免重复
  enhanceTableStylesDisabled(html) {
    if (!html) return html;
    
    // 处理表格
    html = html.replace(/<table([^>]*)>/gi, (match, attrs, offset) => {
      // 获取这个表格的内容
      const tableEndIndex = html.indexOf('</table>', offset);
      const tableContent = html.substring(offset, tableEndIndex);
      const firstRow = tableContent.match(/<tr[^>]*>(.*?)<\/tr>/si);
      
      if (firstRow) {
        // 计算列数
        const cellMatches = firstRow[1].match(/<t[dh][^>]*>/gi) || [];
        const cellCount = cellMatches.length;
        
        if (cellCount === 2) {
          // 两列表格 - 类似"审计概述"的格式
          return `<table class="two-column-table" style="border-collapse: collapse; width: 100%; margin: 20px 0; border: 1px solid #333; background: white;"${attrs}>`;
        } else if (cellCount > 2) {
          // 多列表格
          return `<table class="multi-column-table" style="border-collapse: collapse; width: 100%; margin: 20px 0; border: 1px solid #333; background: white;"${attrs}>`;
        }
      }
      return `<table style="border-collapse: collapse; width: 100%; margin: 20px 0; border: 1px solid #333; background: white;"${attrs}>`;
    });
    
    // 处理两列表格的特殊样式
    html = html.replace(/<table class="two-column-table"[^>]*>(.*?)<\/table>/gis, (match, content) => {
      // 处理每一行
      let processedContent = content.replace(/<tr([^>]*)>/gi, (trMatch, trAttrs, trOffset) => {
        // 找到这一行的结束
        const trEndIndex = content.indexOf('</tr>', trOffset);
        const rowContent = content.substring(trOffset, trEndIndex);
        
        // 处理这一行的单元格
        let cellIndex = 0;
        const processedRow = rowContent.replace(/<td([^>]*)>/gi, (tdMatch, tdAttrs) => {
          cellIndex++;
          if (cellIndex === 1) {
            // 第一列 - 标签列
            return `<td style="border: 1px solid #333; padding: 10px 20px; width: 25%; min-width: 120px; background-color: #f8f8f8; text-align: right; vertical-align: middle;"${tdAttrs}>`;
          } else {
            // 第二列 - 内容列
            return `<td style="border: 1px solid #333; padding: 10px 20px; width: 75%; vertical-align: middle;"${tdAttrs}>`;
          }
        });
        
        return `<tr${trAttrs}>` + processedRow.substring(processedRow.indexOf('>') + 1);
      });
      
      return match.replace(content, processedContent);
    });
    
    // 处理普通表格单元格
    html = html.replace(/<td([^>]*)>/gi, (match, attrs) => {
      // 如果已经有style属性，不要重复添加
      if (attrs.includes('style=')) {
        return match;
      }
      return `<td style="border: 1px solid #333; padding: 10px 15px; vertical-align: middle;"${attrs}>`;
    });
    
    html = html.replace(/<th([^>]*)>/gi, (match, attrs) => {
      if (attrs.includes('style=')) {
        return match;
      }
      return `<th style="border: 1px solid #333; padding: 10px 15px; vertical-align: middle; background-color: #f8f8f8; font-weight: bold; text-align: center;"${attrs}>`;
    });
    
    return html;
  }
  
  // 增强整个文档的样式 - 暂时禁用
  enhanceFullDocumentDisabled(html) {
    if (!html) return html;
    
    let enhanced = html;
    
    // 检测并处理封面
    enhanced = this.detectAndEnhanceCover(enhanced);
    
    // 处理两列表格的特殊样式
    enhanced = enhanced.replace(/<table class="two-column-table"([^>]*)>(.*?)<\/table>/gis, (match, attrs, content) => {
      // 处理每一行
      let processedContent = content.replace(/<tr([^>]*)>/gi, '<tr$1>');
      
      // 处理第一列（标签列）
      processedContent = processedContent.replace(/<tr([^>]*)>\s*<td([^>]*)>/gi, 
        '<tr$1><td style="border: 1px solid #000; padding: 10px 15px; width: 120px; vertical-align: middle;"$2>');
      
      return `<table class="audit-overview-table" style="border-collapse: collapse; width: 100%; border: 1px solid #000; margin: 20px 0; font-family: 宋体, SimSun; font-size: 12pt;"${attrs}>${processedContent}</table>`;
    });
    
    // 处理标题
    enhanced = enhanced.replace(/<h1([^>]*)>([^<]+)<\/h1>/gi, 
      '<h1 style="font-size: 22pt; font-weight: bold; text-align: center; margin: 30px 0; font-family: 黑体, SimHei; letter-spacing: 2px;">$2</h1>');
    
    // 处理段落
    enhanced = enhanced.replace(/<p([^>]*)>([^<]+)<\/p>/gi, (match, attrs, content) => {
      // 检查是否是居中段落
      if (attrs.includes('text-align') && attrs.includes('center')) {
        return `<p style="text-align: center; margin: 15px 0; font-size: 16pt; font-family: 黑体, SimHei;">${content}</p>`;
      }
      // 普通段落
      return `<p style="margin: 10px 0; line-height: 1.8; text-indent: 2em; font-size: 12pt; font-family: 宋体, SimSun; text-align: justify;">${content}</p>`;
    });
    
    // 处理列表
    enhanced = enhanced.replace(/<ul([^>]*)>/gi, '<ul style="margin: 10px 0; padding-left: 30px; font-size: 12pt;"$1>');
    enhanced = enhanced.replace(/<ol([^>]*)>/gi, '<ol style="margin: 10px 0; padding-left: 30px; font-size: 12pt;"$1>');
    enhanced = enhanced.replace(/<li([^>]*)>/gi, '<li style="margin: 5px 0; line-height: 1.8;"$1>');
    
    return enhanced;
  }
  
  // 检测并增强封面 - 简化版本
  detectAndEnhanceCover(html) {
    console.log('开始检测封面特征...');
    
    // 简单检测：前500字符内有多个短文本段落
    const firstContent = html.substring(0, 1000);
    const paragraphs = firstContent.match(/<p[^>]*>([^<]{1,100})<\/p>/gi) || [];
    
    if (paragraphs.length >= 3) {
      const texts = paragraphs.map(p => p.replace(/<[^>]*>/g, '').trim());
      console.log('找到前几个段落:', texts.slice(0, 5));
      
      // 检查是否像封面
      const hasShortTexts = texts.every(t => t.length < 50);
      const hasKeywords = texts.some(t => 
        t.includes('API') || t.includes('报') || t.includes('审计') || 
        t.includes('公司') || t.includes('日') || t.includes('室')
      );
      
      if (hasShortTexts && (hasKeywords || paragraphs.length >= 4)) {
        console.log('检测到封面!');
        
        // 创建简单封面
        const coverHtml = this.createSimpleCover(texts);
        
        // 找到第一个表格或长段落
        const tablePos = html.indexOf('<table');
        const longParagraph = html.match(/<p[^>]*>[^<]{100,}<\/p>/i);
        const contentStart = Math.min(
          tablePos > 0 ? tablePos : Infinity,
          longParagraph ? html.indexOf(longParagraph[0]) : Infinity
        );
        
        if (contentStart < Infinity) {
          return coverHtml + html.substring(contentStart);
        }
        return coverHtml + html;
      }
    }
    
    // 不做封面处理
    return html;
  }
  
  // 创建简单封面
  createSimpleCover(texts) {
    const title1 = texts[0] || '';
    const title2 = texts[1] || '';
    const title3 = texts[2] || '';
    const title4 = texts[3] || '';
    
    return `
      <div style="
        min-height: 100vh;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        page-break-after: always;
        font-family: '宋体', SimSun;
      ">
        <h1 style="font-size: 22pt; margin: 20px 0;">${title1}</h1>
        <h1 style="font-size: 22pt; margin: 20px 0;">${title2}</h1>
        <p style="font-size: 14pt; margin: 20px 0;">${title3}</p>
        <p style="font-size: 12pt; margin: 20px 0;">${title4}</p>
      </div>
    `;
  }
  
  // 原有的复杂检测逻辑 - 暂时禁用
  detectAndEnhanceCoverOld(html) {
    const coverConfig = {
      // 方式1：关键字检测
      keywords: {
        title: ['API', '接口', '系统', '平台', '报告', '方案'],
        type: ['审计', '日报', '周报', '月报', '专题', '分析', '总结', '报告']
      },
      // 方式2：布局特征检测
      layoutPatterns: {
        minCenteredTexts: 2,  // 至少2个居中文本
        maxTextLength: 50,     // 每个文本最大长度
        datePattern: /[（(]?\d{4}年\d{1,2}月\d{1,2}日[）)]?/  // 日期格式
      },
      // 方式3：特殊标记检测
      markers: ['[COVER]', '【封面】', '{{COVER_PAGE}}']
    };
    
    // 检查特殊标记（优先级最高）
    for (const marker of coverConfig.markers) {
      if (html.includes(marker)) {
        console.log('检测到封面标记:', marker);
        html = html.replace(marker, '');
        return this.insertCoverAtBeginning(html);
      }
    }
    
    // 检查居中文本布局
    const centerPattern = /<p[^>]*style="[^"]*text-align:\s*center[^"]*"[^>]*>([^<]{1,100})<\/p>/gi;
    const centerMatches = [...html.matchAll(centerPattern)];
    
    if (centerMatches.length >= coverConfig.layoutPatterns.minCenteredTexts) {
      const texts = centerMatches.map(m => m[1].trim());
      
      // 方式1：智能关键字匹配
      const hasTitleKeyword = texts.some(t => 
        coverConfig.keywords.title.some(kw => t.includes(kw))
      );
      const hasTypeKeyword = texts.some(t => 
        coverConfig.keywords.type.some(kw => t.includes(kw))
      );
      
      // 方式2：检查是否有日期格式
      const hasDate = texts.some(t => 
        coverConfig.layoutPatterns.datePattern.test(t)
      );
      
      // 方式3：检查文档开头位置
      const isAtBeginning = html.indexOf(centerMatches[0][0]) < 200;
      
      // 综合判断（满足以下任一条件即认为是封面）
      const isCover = 
        (hasTitleKeyword && hasTypeKeyword) ||  // 有标题和类型关键字
        (centerMatches.length >= 3 && hasDate) ||  // 3个以上居中文本且有日期
        (isAtBeginning && centerMatches.length >= 2 && texts[0].length < 30);  // 在文档开头且有短标题
      
      if (isCover) {
        console.log('检测到封面布局，居中文本:', texts.slice(0, 3));
        
        // 提取封面信息
        const coverInfo = this.extractCoverInfo(texts, hasDate);
        const coverHtml = this.createCoverHtml(coverInfo);
        
        // 移除原始的居中文本
        let enhancedHtml = html;
        centerMatches.slice(0, Math.min(4, centerMatches.length)).forEach(match => {
          enhancedHtml = enhancedHtml.replace(match[0], '');
        });
        
        return coverHtml + enhancedHtml;
      }
    }
    
    // 方式4：检查是否只有标题没有实质内容（极简封面）
    const cleanText = html.replace(/<[^>]*>/g, '').trim();
    if (cleanText.length < 200 && centerMatches.length > 0) {
      console.log('检测到极简封面');
      const texts = centerMatches.map(m => m[1].trim());
      const coverInfo = this.extractCoverInfo(texts, false);
      return this.createCoverHtml(coverInfo) + html;
    }
    
    return html;
  }
  
  // 提取封面信息
  extractCoverInfo(texts, hasDate) {
    const info = {
      mainTitle: '',
      subTitle: '',
      date: ''
    };
    
    // 智能分配标题和日期
    if (texts.length >= 3 && hasDate) {
      // 最后一个通常是日期
      const datePattern = /[（(]?\d{4}年\d{1,2}月\d{1,2}日[）)]?/;
      const dateIndex = texts.findIndex(t => datePattern.test(t));
      
      if (dateIndex >= 0) {
        info.date = texts[dateIndex];
        texts.splice(dateIndex, 1);
      }
    }
    
    // 分配主副标题
    if (texts.length >= 2) {
      info.mainTitle = texts[0];
      info.subTitle = texts[1];
    } else if (texts.length === 1) {
      // 尝试拆分单个标题
      const title = texts[0];
      if (title.includes('API') || title.includes('接口')) {
        info.mainTitle = title;
        info.subTitle = '专题审计日报';
      } else {
        info.mainTitle = '审计报告';
        info.subTitle = title;
      }
    }
    
    // 如果没有日期，生成默认日期
    if (!info.date) {
      const today = new Date();
      info.date = `（${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日）`;
    }
    
    return [info.mainTitle, info.subTitle, info.date];
  }
  
  // 在文档开头插入默认封面
  insertCoverAtBeginning(html) {
    const defaultCover = ['省公司 API 接口', '专题审计日报', new Date().toLocaleDateString('zh-CN', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    })];
    return this.createCoverHtml(defaultCover) + html;
  }
  
  // 创建封面HTML
  createCoverHtml(titles) {
    // 通用封面创建，不预设默认值
    const mainTitle = titles[0] || '';
    const subTitle = titles[1] || '';
    const dateStr = titles[2] || '';
    
    return `
      <div class="word-cover" style="
        width: 100%;
        min-height: 100vh;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        position: relative;
        page-break-after: always;
        background: white;
      ">
        <!-- 页面边框装饰 -->
        <div style="
          position: absolute;
          top: 60px;
          left: 60px;
          right: 60px;
          bottom: 60px;
        ">
          <!-- 角落装饰 -->
          <div style="position: absolute; top: 0; left: 0; width: 60px; height: 60px; border-top: 2px solid #666; border-left: 2px solid #666;"></div>
          <div style="position: absolute; top: 0; right: 0; width: 60px; height: 60px; border-top: 2px solid #666; border-right: 2px solid #666;"></div>
          <div style="position: absolute; bottom: 0; left: 0; width: 60px; height: 60px; border-bottom: 2px solid #666; border-left: 2px solid #666;"></div>
          <div style="position: absolute; bottom: 0; right: 0; width: 60px; height: 60px; border-bottom: 2px solid #666; border-right: 2px solid #666;"></div>
        </div>
        
        <!-- 标题内容 -->
        <div style="text-align: center; z-index: 1;">
          <h1 style="
            font-size: 42pt;
            font-weight: bold;
            color: #1e88e5;
            font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;
            letter-spacing: 8px;
            margin: 0 0 30px 0;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
          ">${mainTitle}</h1>
          <h1 style="
            font-size: 42pt;
            font-weight: bold;
            color: #1e88e5;
            font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;
            letter-spacing: 8px;
            margin: 0;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
          ">${subTitle}</h1>
        </div>
        
        <!-- 日期 -->
        <div style="
          position: absolute;
          bottom: 200px;
          width: 100%;
          text-align: center;
        ">
          <p style="
            font-size: 18pt;
            color: #333;
            font-family: '宋体', SimSun, serif;
            letter-spacing: 2px;
          ">${dateStr.includes('（') ? dateStr : `（${dateStr}）`}</p>
        </div>
      </div>
      <div style="page-break-after: always;"></div>
    `;
  }
  
  // 异步yield，避免长时间阻塞事件循环
  async asyncYield() {
    return new Promise(resolve => setImmediate(resolve));
  }
  
  // 带超时保护的表格处理
  async processTableWithTimeout(tableMatch, globalIndex, html, options) {
    const timeout = 30000; // 30秒超时
    
    const timeoutPromise = new Promise((_, reject) => {
      setTimeout(() => reject(new Error('Table processing timeout')), timeout);
    });
    
    const processPromise = this.processSingleTable(tableMatch, globalIndex, html, options);
    
    try {
      return await Promise.race([processPromise, timeoutPromise]);
    } catch (error) {
      console.error(`Table ${globalIndex + 1} processing failed:`, error.message);
      return null; // 返回null而不是抛出错误，让批次处理继续
    }
  }
  
  // 处理单个表格
  async processSingleTable(tableMatch, globalIndex, html, options) {
    const tableHtml = tableMatch[0];
    const tableStructure = await this.extractTableStructureAdvanced(tableHtml, options);
    
    // 查找表格前后的文本作为标题
    const tablePos = tableMatch.index;
    const beforeText = html.substring(Math.max(0, tablePos - 500), tablePos);
    
    console.log(`\n=== 表格 ${globalIndex + 1} 标题提取 ===`);
    
    let title = this.extractNearbyTitle(beforeText);
    if (!title) {
      const widerText = html.substring(Math.max(0, tablePos - 1000), tablePos);
      title = this.extractNearbyTitle(widerText);
    }
    
    if (!title) {
      title = `表格 ${globalIndex + 1}`;
      console.log('No title found, using fallback:', title);
    } else {
      console.log('Found title:', title);
    }
    
    const type = this.inferSectionType(title, tableStructure);
    const fields = this.extractStaticFields(tableHtml);

    return {
      id: `table_section_${globalIndex}`,
      level: 1,
      title: title,
      hasTable: tableStructure.hasTable,
      hasContent: true,
      contentPreview: this.cleanHtml(tableHtml).substring(0, 100),
      tableStructure: tableStructure,
      rawContent: tableHtml.substring(0, 200),
      originalHtml: tableHtml,
      type: type,
      fields: fields
    };
  }
  
  // 带超时保护的单个表格解析
  async parseIndividualTableWithTimeout(tableHtml, tableIndex, options) {
    const timeout = 10000; // 10秒超时
    
    const timeoutPromise = new Promise((_, reject) => {
      setTimeout(() => reject(new Error('Individual table parsing timeout')), timeout);
    });
    
    const parsePromise = Promise.resolve(this.parseIndividualTable(tableHtml, tableIndex, options));
    
    return await Promise.race([parsePromise, timeoutPromise]);
  }
  
  // 增强中文字体样式处理
  enhanceChineseFontStyles(html) {
    if (!html) return html;
    
    console.log('🎨 开始增强中文字体样式...');
    
    // 在文档开头添加完整的CSS样式
    const chineseFontCSS = `
      <style>
        /* 基础字体样式 - 针对中文优化 */
        body, * {
          font-family: "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "Segoe UI", Tahoma, "Helvetica Neue", Helvetica, Arial, "PingFang SC", "Hiragino Sans GB", "Heiti SC", "WenQuanYi Micro Hei", sans-serif !important;
          font-synthesis: none;
          text-rendering: optimizeLegibility;
          -webkit-font-smoothing: antialiased;
          -moz-osx-font-smoothing: grayscale;
        }
        
        /* Word标准字体映射 */
        .font-songti, p, td, th, span, div {
          font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", "Times New Roman", serif !important;
          font-size: 12pt;
          line-height: 1.6;
        }
        
        .font-heiti, h1, h2, h3, h4, h5, h6 {
          font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
          font-weight: bold;
        }
        
        .font-yahei {
          font-family: "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
        }
        
        /* 标题字体优化 */
        h1 {
          font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
          font-size: 18pt;
          font-weight: bold;
          text-align: center;
          margin: 24pt 0 12pt 0;
          line-height: 1.3;
        }
        
        h2 {
          font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
          font-size: 16pt;
          font-weight: bold;
          margin: 18pt 0 6pt 0;
          line-height: 1.3;
        }
        
        h3 {
          font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
          font-size: 14pt;
          font-weight: bold;
          margin: 12pt 0 6pt 0;
          line-height: 1.3;
        }
        
        /* 段落字体优化 */
        p {
          font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
          font-size: 12pt;
          line-height: 1.75;
          margin: 6pt 0;
          text-align: justify;
          text-justify: inter-ideograph;
        }
        
        /* 表格字体优化 */
        table {
          font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
          border-collapse: collapse;
          width: 100%;
          margin: 12pt 0;
        }
        
        td, th {
          font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
          font-size: 10.5pt;
          line-height: 1.5;
          padding: 4pt 6pt;
          border: 0.5pt solid #000000;
          vertical-align: middle;
        }
        
        th {
          font-weight: bold;
          text-align: center;
          background-color: #f8f8f8;
        }
        
        /* 列表字体优化 */
        ul, ol, li {
          font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
          font-size: 12pt;
          line-height: 1.6;
        }
        
        /* 强调文本 */
        strong, b {
          font-weight: bold;
        }
        
        em, i {
          font-style: italic;
        }
        
        /* 确保中文字符正确显示 */
        .chinese-content {
          font-variant-numeric: normal;
          font-feature-settings: normal;
        }
        
        /* 针对特定中文标点的处理 */
        .punctuation-fix {
          font-kerning: normal;
          text-spacing: ideograph-alpha ideograph-numeric;
        }
        
        /* Word表格样式兼容 */
        .word-table td:first-child {
          font-weight: bold;
          background-color: #f5f5f5;
          min-width: 120px;
        }
      </style>
    `;
    
    // 将CSS插入到HTML开头
    if (html.includes('<body>') || html.includes('<html>')) {
      // 如果是完整HTML文档
      html = html.replace('<head>', `<head>${chineseFontCSS}`)
        .replace(/<body([^>]*)>/, '<body$1 class="chinese-content punctuation-fix">');
    } else {
      // 如果是HTML片段，在开头添加样式
      html = `${chineseFontCSS}<div class="chinese-content punctuation-fix">${html}</div>`;
    }
    
    // 为所有段落添加宋体类
    html = html.replace(/<p([^>]*)>/gi, '<p$1 class="font-songti">');
    
    // 为所有标题添加黑体类
    html = html.replace(/<(h[1-6])([^>]*)>/gi, '<$1$2 class="font-heiti">');
    
    // 为表格添加字体优化
    html = html.replace(/<table([^>]*)>/gi, '<table$1 class="word-table font-songti">');
    
    // 处理表格单元格
    html = html.replace(/<(td|th)([^>]*)>/gi, '<$1$2 class="font-songti">');
    
    console.log('✅ 中文字体样式增强完成');
    return html;
  }
}

module.exports = UniversalWordParser;