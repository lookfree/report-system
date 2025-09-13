const mammoth = require('mammoth');
const htmlDocx = require('html-docx-js');
const fs = require('fs').promises;
const path = require('path');

class HtmlWordConverter {
  /**
   * 将Word文档转换为HTML（保留完整格式）
   */
  async wordToHtml(filePath) {
    try {
      const result = await mammoth.convertToHtml({
        path: filePath,
        options: {
          // 保留所有样式和格式
          styleMap: [
            // 标题样式映射
            "p[style-name='Heading 1'] => h1:fresh",
            "p[style-name='Heading 2'] => h2:fresh", 
            "p[style-name='Heading 3'] => h3:fresh",
            "p[style-name='Heading 4'] => h4:fresh",
            "p[style-name='标题 1'] => h1:fresh",
            "p[style-name='标题 2'] => h2:fresh",
            "p[style-name='标题 3'] => h3:fresh",
            "p[style-name='标题 4'] => h4:fresh",
            "p[style-name='Title'] => h1:fresh",
            "p[style-name='Subtitle'] => h2:fresh",
            
            // 文本格式
            "b => strong",
            "i => em", 
            "u => u",
            "strike => del",
            
            // 段落保持原样
            "p => p:fresh",
            "p[style-name='Normal'] => p:fresh",
            "p[style-name='正文'] => p:fresh",
            
            // 表格样式 - 改进复杂表格处理
            "table => table.word-table:fresh",
            "tr => tr:fresh",
            "td => td:fresh",
            "th => th:fresh",
            
            // 列表样式
            "ul => ul",
            "ol => ol", 
            "li => li"
          ],
          
          // 保留样式信息
          includeDefaultStyleMap: true,
          includeEmbeddedStyleMap: true,
          
          // 保留图片
          convertImage: mammoth.images.imgElement(function(image) {
            return image.read("base64").then(function(imageBuffer) {
              return {
                src: "data:" + image.contentType + ";base64," + imageBuffer
              };
            });
          }),
          
          // 保留更多文档结构 - 改进复杂表格和分页处理
          transformDocument: mammoth.transforms.compose([
            // 保留段落对齐
            mammoth.transforms.paragraph(function(element) {
              if (element.alignment) {
                const alignment = element.alignment;
                return {
                  ...element,
                  attributes: {
                    ...element.attributes,
                    'data-align': alignment
                  }
                };
              }
              return element;
            }),
            
            // 保留表格属性
            mammoth.transforms.table(function(element) {
              return {
                ...element,
                attributes: {
                  ...element.attributes,
                  'data-table-complex': 'true'
                }
              };
            }),
            
            // 保留表格行属性
            mammoth.transforms.tableRow(function(element) {
              return {
                ...element,
                attributes: {
                  ...element.attributes,
                  'data-row': 'true'
                }
              };
            }),
            
            // 保留表格单元格属性 - 处理合并单元格
            mammoth.transforms.tableCell(function(element) {
              const attributes = {
                ...element.attributes,
                'data-cell': 'true'
              };
              
              // 保留colspan和rowspan
              if (element.colSpan && element.colSpan > 1) {
                attributes.colspan = element.colSpan;
              }
              if (element.rowSpan && element.rowSpan > 1) {
                attributes.rowspan = element.rowSpan;
              }
              
              return {
                ...element,
                attributes: attributes
              };
            })
          ]),
          
          // 处理分页符
          convertBreak: function(breakType) {
            if (breakType === 'page') {
              return '<div class="page-break"></div>';
            }
            return null;
          }
        }
      });

      // 处理复杂表格结构
      console.log('🔄 开始处理复杂表格结构...');
      const processedHtml = this.processComplexTables(result.value);
      
      // 添加自定义样式
      const styledHtml = this.addCustomStyles(processedHtml);
      
      // 输出转换警告信息以便诊断
      if (result.messages && result.messages.length > 0) {
        console.log('📋 Word转换警告信息:');
        result.messages.forEach((message, index) => {
          console.log(`  ${index + 1}. ${message.type}: ${message.message}`);
        });
      }
      
      // 输出原始HTML以便调试复杂表格问题
      console.log('🔍 原始mammoth转换结果长度:', result.value.length);
      console.log('🔍 处理后HTML长度:', processedHtml.length);
      if (result.value.includes('<table')) {
        const tableCount = (result.value.match(/<table/g) || []).length;
        console.log(`📊 检测到 ${tableCount} 个表格`);
      }
      
      return {
        html: styledHtml,
        messages: result.messages,
        rawHtml: result.value, // 添加原始HTML用于调试
        processedHtml: processedHtml // 添加处理后的HTML用于调试
      };
    } catch (error) {
      console.error('Word to HTML conversion error:', error);
      throw new Error('文档转换失败: ' + error.message);
    }
  }

  /**
   * 添加自定义样式到HTML
   */
  addCustomStyles(html) {
    const styles = `
      <style>
        /* Word文档基础样式 */
        body {
          font-family: '宋体', 'SimSun', '微软雅黑', 'Microsoft YaHei', Arial, sans-serif;
          font-size: 12pt;
          color: #000;
          line-height: 1.5;
          margin: 0;
          padding: 20px;
          background: white;
        }
        
        /* Word表格样式 - 改进复杂表格支持 */
        table, .word-table {
          border-collapse: collapse;
          width: 100%;
          margin: 10px 0;
          border: 1px solid #000;
          font-size: 12pt;
          table-layout: auto; /* 允许自适应列宽 */
          word-wrap: break-word;
          word-break: break-all;
        }
        
        table td, table th, .word-table td, .word-table th {
          border: 1px solid #000;
          padding: 5px 8px;
          vertical-align: top;
          text-align: left;
          line-height: 1.4;
          word-wrap: break-word;
          overflow-wrap: break-word;
          max-width: 200px; /* 防止单元格过宽 */
        }
        
        table th, .word-table th {
          background-color: #f2f2f2;
          font-weight: bold;
          text-align: center;
        }
        
        /* 合并单元格样式 */
        table td[colspan], table th[colspan],
        .word-table td[colspan], .word-table th[colspan] {
          text-align: center;
          font-weight: bold;
        }
        
        table td[rowspan], table th[rowspan],
        .word-table td[rowspan], .word-table th[rowspan] {
          vertical-align: middle;
        }
        
        /* 复杂表格特殊处理 */
        table[data-table-complex="true"] {
          margin: 15px 0;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        /* 跨页表格处理 */
        table.page-break-table {
          page-break-inside: auto;
          break-inside: auto;
        }
        
        table.page-break-table thead {
          display: table-header-group;
        }
        
        table.page-break-table tbody {
          display: table-row-group;
        }
        
        table.page-break-table tr {
          page-break-inside: avoid;
          break-inside: avoid;
        }
        
        /* Word标题样式 */
        h1 { 
          font-size: 18pt; 
          font-weight: bold; 
          margin: 20px 0 15px 0;
          color: #000;
          text-align: center;
          font-family: '黑体', 'SimHei', sans-serif;
        }
        
        h2 { 
          font-size: 16pt; 
          font-weight: bold; 
          margin: 15px 0 10px 0;
          color: #000;
          font-family: '黑体', 'SimHei', sans-serif;
        }
        
        h3 { 
          font-size: 14pt; 
          font-weight: bold; 
          margin: 12px 0 8px 0;
          color: #000;
          font-family: '黑体', 'SimHei', sans-serif;
        }
        
        h4 { 
          font-size: 12pt; 
          font-weight: bold; 
          margin: 10px 0 6px 0;
          color: #000;
          font-family: '黑体', 'SimHei', sans-serif;
        }
        
        /* Word段落样式 */
        p {
          margin: 6px 0;
          line-height: 1.5;
          text-align: justify;
          text-indent: 0;
          font-size: 12pt;
        }
        
        /* Word列表样式 */
        ul, ol {
          margin: 6px 0;
          padding-left: 20px;
        }
        
        li {
          margin: 3px 0;
          line-height: 1.5;
        }
        
        /* 文本格式 */
        strong, b {
          font-weight: bold;
        }
        
        em, i {
          font-style: italic;
        }
        
        u {
          text-decoration: underline;
        }
        
        /* 保留原有的内联样式 */
        [style*="font-weight: bold"], [style*="font-weight:bold"] {
          font-weight: bold !important;
        }
        
        [style*="font-style: italic"], [style*="font-style:italic"] {
          font-style: italic !important;
        }
        
        [style*="text-decoration: underline"], [style*="text-decoration:underline"] {
          text-decoration: underline !important;
        }
        
        [style*="text-align: center"], [style*="text-align:center"] {
          text-align: center !important;
        }
        
        [style*="text-align: right"], [style*="text-align:right"] {
          text-align: right !important;
        }
        
        [style*="text-align: left"], [style*="text-align:left"] {
          text-align: left !important;
        }
        
        /* 动态字段样式 */
        .dynamic-field {
          background-color: #e3f2fd;
          padding: 2px 6px;
          border-radius: 3px;
          border: 1px dashed #2196f3;
          display: inline-block;
          min-width: 60px;
        }
        
        /* 可编辑单元格样式 */
        .editable-cell {
          background-color: #fff9c4;
          min-height: 20px;
        }
        
        .editable-cell:focus {
          outline: 2px solid #2196f3;
          background-color: #fff;
        }
        
        /* 图片样式 */
        img {
          max-width: 100%;
          height: auto;
          display: block;
          margin: 10px auto;
        }
        
        /* 分页符 */
        .page-break {
          page-break-before: always;
          break-before: page;
        }
      </style>
    `;
    
    return styles + html;
  }

  /**
   * 处理动态数据替换
   */
  async processDynamicData(html, prisma) {
    try {
      const cheerio = require('cheerio');
      const $ = cheerio.load(html);
      
      // 处理动态字段
      const dynamicFields = $('.dynamic-field');
      for (let i = 0; i < dynamicFields.length; i++) {
        const field = $(dynamicFields[i]);
        const fieldType = field.attr('data-field-type');
        const dataSourceId = field.attr('data-source-id');
        const sqlQuery = field.attr('data-sql');
        const defaultValue = field.attr('data-default');
        
        let value = defaultValue || '';
        
        if (fieldType === 'DYNAMIC' && dataSourceId && sqlQuery) {
          try {
            // 这里需要根据实际的数据源配置执行SQL查询
            // 简化示例：返回当前时间
            value = new Date().toLocaleString('zh-CN');
          } catch (error) {
            console.error('SQL查询失败:', error);
            value = defaultValue || '数据获取失败';
          }
        }
        
        field.text(value);
        field.removeClass('dynamic-field');
        field.removeAttr('data-field-type data-source-id data-sql data-default');
      }
      
      // 处理动态表格
      const dynamicTables = $('.dynamic-table');
      for (let i = 0; i < dynamicTables.length; i++) {
        const tableDiv = $(dynamicTables[i]);
        const dataSourceId = tableDiv.attr('data-source-id');
        const sqlQuery = tableDiv.attr('data-sql');
        const tableTitle = tableDiv.attr('data-table-title');
        
        if (dataSourceId && sqlQuery) {
          try {
            // 简化示例：生成示例数据
            const sampleData = [
              { col1: '数据1', col2: '值1', col3: '结果1' },
              { col1: '数据2', col2: '值2', col3: '结果2' },
              { col1: '数据3', col2: '值3', col3: '结果3' }
            ];
            
            let tableHtml = `<h4>${tableTitle}</h4><table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">`;
            
            if (sampleData.length > 0) {
              // 表头
              tableHtml += '<thead><tr style="background-color: #f5f5f5;">';
              Object.keys(sampleData[0]).forEach(key => {
                tableHtml += `<th style="border: 1px solid #ddd; padding: 8px;">${key}</th>`;
              });
              tableHtml += '</tr></thead>';
              
              // 数据行
              tableHtml += '<tbody>';
              sampleData.forEach(row => {
                tableHtml += '<tr>';
                Object.values(row).forEach(value => {
                  tableHtml += `<td style="border: 1px solid #ddd; padding: 8px;">${value}</td>`;
                });
                tableHtml += '</tr>';
              });
              tableHtml += '</tbody>';
            }
            
            tableHtml += '</table>';
            tableDiv.html(tableHtml);
          } catch (error) {
            console.error('表格数据查询失败:', error);
            tableDiv.find('tbody').html('<tr><td colspan="3" style="text-align: center; color: #999;">数据获取失败</td></tr>');
          }
        }
        
        tableDiv.removeClass('dynamic-table');
        tableDiv.removeAttr('data-table-id data-table-title data-source-id data-sql');
        tableDiv.css('border', 'none').css('padding', '0');
      }
      
      return $.html();
    } catch (error) {
      console.error('动态数据处理失败:', error);
      return html;
    }
  }

  /**
   * 将HTML转换为Word文档
   */
  async htmlToWord(html, outputPath, prisma = null) {
    try {
      // 处理动态数据
      if (prisma) {
        html = await this.processDynamicData(html, prisma);
      }
      
      // 准备HTML文档
      const fullHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { 
              font-family: 'Microsoft YaHei', Arial; 
              font-size: 14px;
              line-height: 1.6;
              color: #333;
            }
            table { 
              border-collapse: collapse; 
              width: 100%;
              margin: 10px 0;
            }
            table td, table th { 
              border: 1px solid #000; 
              padding: 8px;
              vertical-align: top;
            }
            table th {
              background-color: #f0f0f0;
              font-weight: bold;
            }
            h1 { font-size: 24px; margin: 20px 0 10px 0; }
            h2 { font-size: 20px; margin: 15px 0 8px 0; }
            h3 { font-size: 16px; margin: 10px 0 5px 0; }
            p { margin: 8px 0; }
          </style>
        </head>
        <body>
          ${html}
        </body>
        </html>
      `;

      // 转换为Word文档
      const docx = htmlDocx.asBlob(fullHtml, {
        orientation: 'portrait',
        margins: {
          top: 720,    // 1 inch = 720 twips
          right: 720,
          bottom: 720,
          left: 720,
          header: 360,
          footer: 360,
          gutter: 0
        }
      });

      // 将Blob转换为Buffer
      const buffer = Buffer.from(await docx.arrayBuffer());
      
      // 保存文件
      await fs.writeFile(outputPath, buffer);
      
      return outputPath;
    } catch (error) {
      console.error('HTML to Word conversion error:', error);
      throw new Error('导出Word文档失败: ' + error.message);
    }
  }

  /**
   * 处理表格使其可编辑
   */
  makeTablesEditable(html) {
    // 使用正则表达式或DOM解析来处理表格
    const processedHtml = html.replace(
      /<td>/g,
      '<td contenteditable="true" class="editable-cell">'
    );
    
    return processedHtml;
  }

  /**
   * 插入动态字段标记
   */
  insertDynamicFields(html, fields) {
    let processedHtml = html;
    
    fields.forEach(field => {
      const placeholder = `{{${field.name}}}`;
      const fieldHtml = `
        <span class="dynamic-field" 
              data-field-id="${field.id}"
              data-field-name="${field.name}"
              data-field-type="${field.type}"
              contenteditable="false">
          ${placeholder}
        </span>
      `;
      
      processedHtml = processedHtml.replace(
        new RegExp(placeholder, 'g'),
        fieldHtml
      );
    });
    
    return processedHtml;
  }

  /**
   * 提取HTML中的表格结构
   */
  extractTableStructure(html) {
    const tables = [];
    const tableRegex = /<table[^>]*>(.*?)<\/table>/gs;
    let match;
    
    while ((match = tableRegex.exec(html)) !== null) {
      const tableHtml = match[0];
      const structure = this.parseTableHtml(tableHtml);
      tables.push(structure);
    }
    
    return tables;
  }

  /**
   * 解析单个表格HTML
   */
  parseTableHtml(tableHtml) {
    // 提取表头
    const headerRegex = /<th[^>]*>(.*?)<\/th>/gs;
    const headers = [];
    let headerMatch;
    
    while ((headerMatch = headerRegex.exec(tableHtml)) !== null) {
      headers.push({
        text: this.stripHtml(headerMatch[1]),
        html: headerMatch[1]
      });
    }
    
    // 提取行数和列数
    const rowRegex = /<tr[^>]*>/g;
    const rows = (tableHtml.match(rowRegex) || []).length;
    
    return {
      headers,
      rows,
      columns: headers.length || this.countColumns(tableHtml),
      html: tableHtml
    };
  }

  /**
   * 计算表格列数
   */
  countColumns(tableHtml) {
    const firstRowMatch = /<tr[^>]*>(.*?)<\/tr>/s.exec(tableHtml);
    if (firstRowMatch) {
      const cellRegex = /<t[dh][^>]*>/g;
      const cells = firstRowMatch[1].match(cellRegex) || [];
      return cells.length;
    }
    return 0;
  }

  /**
   * 去除HTML标签
   */
  stripHtml(html) {
    return html.replace(/<[^>]*>/g, '').trim();
  }

  /**
   * 验证Word文档格式
   */
  async validateWordDocument(filePath) {
    try {
      const buffer = await fs.readFile(filePath);
      
      // 检查文件头魔数
      const isDocx = buffer[0] === 0x50 && buffer[1] === 0x4B; // PK (ZIP格式)
      const isDoc = buffer[0] === 0xD0 && buffer[1] === 0xCF; // 旧版doc格式
      
      if (!isDocx && !isDoc) {
        throw new Error('不是有效的Word文档格式');
      }
      
      if (isDoc) {
        throw new Error('检测到旧版Word格式(.doc)，请使用.docx格式');
      }
      
      return true;
    } catch (error) {
      throw new Error('文档验证失败: ' + error.message);
    }
  }

  /**
   * 处理复杂表格结构
   */
  processComplexTables(html) {
    const cheerio = require('cheerio');
    const $ = cheerio.load(html);
    
    // 处理所有表格
    $('table').each(function() {
      const table = $(this);
      
      // 添加复杂表格标记
      table.attr('data-table-complex', 'true');
      
      // 检查是否有合并单元格
      let hasMergedCells = false;
      table.find('td[colspan], td[rowspan], th[colspan], th[rowspan]').each(function() {
        hasMergedCells = true;
        $(this).addClass('merged-cell');
      });
      
      // 如果表格很大，添加分页处理
      const rowCount = table.find('tr').length;
      if (rowCount > 20) {
        table.addClass('page-break-table');
      }
      
      // 处理表格结构
      const hasHeader = table.find('th').length > 0;
      if (!hasHeader) {
        // 如果没有表头，将第一行设为表头
        const firstRow = table.find('tr').first();
        if (firstRow.length > 0) {
          firstRow.find('td').each(function() {
            const td = $(this);
            const th = $('<th>').html(td.html());
            // 复制所有属性
            Array.from(td[0].attributes).forEach(attr => {
              th.attr(attr.name, attr.value);
            });
            td.replaceWith(th);
          });
          
          // 创建thead和tbody结构
          if (!table.find('thead').length) {
            const thead = $('<thead>').append(firstRow);
            table.prepend(thead);
            
            const remainingRows = table.find('tr').not(firstRow);
            if (remainingRows.length > 0) {
              const tbody = $('<tbody>').append(remainingRows);
              table.append(tbody);
            }
          }
        }
      }
      
      // 确保表格有正确的结构
      if (table.find('thead').length === 0 && table.find('th').length > 0) {
        const headerRow = table.find('tr').has('th').first();
        if (headerRow.length > 0) {
          const thead = $('<thead>').append(headerRow);
          table.prepend(thead);
          
          const remainingRows = table.find('tr').not(headerRow);
          if (remainingRows.length > 0 && table.find('tbody').length === 0) {
            const tbody = $('<tbody>').append(remainingRows);
            table.append(tbody);
          }
        }
      }
      
      console.log(`📊 处理表格: ${rowCount} 行, 合并单元格: ${hasMergedCells}`);
    });
    
    return $.html();
  }

  /**
   * 处理合并单元格
   */
  processMergedCells(html) {
    // 处理colspan和rowspan
    const processedHtml = html.replace(
      /<td\s+colspan="(\d+)"/g,
      (match, colspan) => {
        return `<td colspan="${colspan}" class="merged-cell"`;
      }
    ).replace(
      /<td\s+rowspan="(\d+)"/g,
      (match, rowspan) => {
        return `<td rowspan="${rowspan}" class="merged-cell"`;
      }
    );
    
    return processedHtml;
  }

  /**
   * 清理Word HTML（去除多余的样式和标签）
   */
  cleanWordHtml(html) {
    // 移除Word特有的标签和样式
    let cleanHtml = html
      .replace(/<o:p\s*\/?>|<\/o:p>/g, '') // 移除Office标签
      .replace(/<!\[if[^>]*>.*?<!\[endif\]>/gs, '') // 移除条件注释
      .replace(/mso-[^;"]*/g, '') // 移除MSO样式
      .replace(/style="[^"]*"/g, (match) => {
        // 只保留必要的样式
        const important = match.match(/(font-size|color|font-weight|text-align|background-color):[^;]*/g);
        if (important) {
          return `style="${important.join(';')}"`;
        }
        return '';
      });
    
    return cleanHtml;
  }
}

module.exports = new HtmlWordConverter();