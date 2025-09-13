const mammoth = require('mammoth');
const fs = require('fs').promises;

class WordParser {
  async parseTemplate(filePath) {
    try {
      // 首先检查文件是否存在
      await fs.access(filePath);
      
      // 尝试解析文档
      const result = await mammoth.convertToHtml({ path: filePath });
      const html = result.value;
      
      // 检查是否成功解析到内容
      if (!html || html.trim().length === 0) {
        throw new Error('文档解析结果为空，可能是文档格式不支持或文档损坏');
      }
      
      // Parse HTML to extract structure
      const structure = this.extractStructure(html);
      
      return structure;
    } catch (error) {
      console.error('Error parsing Word template:', error);
      
      // 提供更友好的错误信息
      if (error.message && error.message.includes('Could not find the body element')) {
        throw new Error('文档格式错误：无法解析Word文档结构。请确保文档是有效的.docx格式，且未被密码保护或损坏。');
      } else if (error.code === 'ENOENT') {
        throw new Error('文件不存在：' + filePath);
      } else if (error.message && error.message.includes('file is encrypted')) {
        throw new Error('文档被密码保护，请先解除密码保护后重新上传。');
      } else {
        throw new Error('文档解析失败：' + (error.message || '未知错误'));
      }
    }
  }
  
  extractStructure(html) {
    const structure = {
      title: '',
      sections: []
    };
    
    // Extract title
    const titleMatch = html.match(/<h1[^>]*>(.*?)<\/h1>/i);
    if (titleMatch) {
      structure.title = this.cleanHtml(titleMatch[1]);
    }
    
    // Extract all headers and analyze their content
    const headerPattern = /<h([1-3])[^>]*>(.*?)<\/h\1>/gi;
    const matches = [...html.matchAll(headerPattern)];
    
    matches.forEach((match, index) => {
      const level = match[1];
      const text = this.cleanHtml(match[2]);
      const startPos = match.index + match[0].length;
      const endPos = matches[index + 1] ? matches[index + 1].index : html.length;
      const content = html.substring(startPos, endPos);
      
      // Extract table structure if exists
      const tableStructure = this.extractTableStructure(content);
      
      const section = {
        id: `section_${index}`,
        level: parseInt(level),
        title: text,
        hasTable: tableStructure.hasTable,
        hasContent: content.trim().length > 0,
        contentPreview: this.cleanHtml(content).substring(0, 100),
        tableStructure: tableStructure
      };
      
      structure.sections.push(section);
    });
    
    // If no headers found, analyze the whole document
    if (structure.sections.length === 0) {
      const tableStructure = this.extractTableStructure(html);
      if (tableStructure.hasTable) {
        structure.sections.push({
          id: 'section_0',
          level: 1,
          title: '表格数据',
          hasTable: true,
          hasContent: true,
          contentPreview: '文档包含表格数据',
          tableStructure: tableStructure
        });
      } else {
        // Extract paragraphs as fallback
        const paragraphs = html.match(/<p[^>]*>(.*?)<\/p>/gi) || [];
        paragraphs.forEach((para, index) => {
          const text = this.cleanHtml(para);
          if (text.trim()) {
            structure.sections.push({
              id: `section_${index}`,
              level: 1,
              title: text.substring(0, 50) + (text.length > 50 ? '...' : ''),
              hasTable: false,
              hasContent: true,
              contentPreview: text.substring(0, 100),
              tableStructure: { hasTable: false }
            });
          }
        });
      }
    }
    
    return structure;
  }
  
  extractTableStructure(html) {
    const tableStructure = {
      hasTable: false,
      headers: [],
      rowCount: 0,
      columnCount: 0,
      tableType: 'simple' // simple, merged
    };
    
    // Find all tables in the content
    const tableMatches = html.match(/<table[^>]*>.*?<\/table>/gis);
    
    if (tableMatches && tableMatches.length > 0) {
      tableStructure.hasTable = true;
      
      // Parse the first table
      const firstTable = tableMatches[0];
      
      // Extract all rows to analyze structure
      const allRowMatches = [...firstTable.matchAll(/<tr[^>]*>(.*?)<\/tr>/gis)];
      
      if (allRowMatches.length > 0) {
        // Analyze table structure
        const analyzedStructure = this.analyzeTableStructure(allRowMatches);
        tableStructure.headers = analyzedStructure.headers;
        tableStructure.columnCount = analyzedStructure.columnCount;
        tableStructure.tableType = analyzedStructure.tableType;
      }
      
      tableStructure.rowCount = allRowMatches.length;
    }
    
    return tableStructure;
  }
  
  analyzeTableStructure(rowMatches) {
    const structure = {
      headers: [],
      columnCount: 0,
      tableType: 'simple'
    };
    
    // Parse each row to understand the structure
    const rows = rowMatches.map(match => {
      const cellMatches = [...match[1].matchAll(/<t[hd][^>]*(?:\s+colspan=["'](\d+)["'])?[^>]*>(.*?)<\/t[hd]>/gis)];
      return cellMatches.map(cell => ({
        content: this.cleanHtml(cell[2]).trim(),
        colspan: parseInt(cell[1]) || 1
      }));
    });
    
    if (rows.length === 0) return structure;
    
    // Check if it's a merged header table
    const firstRowSpans = rows[0].some(cell => cell.colspan > 1);
    const hasMultipleHeaderRows = rows.length > 1 && rows[1].length > rows[0].length;
    
    // 检查第一行的单元格数量与第二行的单元格数量差异，推断合并表头
    const firstRowCellCount = rows[0].length;
    const secondRowCellCount = rows.length > 1 ? rows[1].length : 0;
    const hasImpliedMerge = secondRowCellCount > firstRowCellCount;
    
    if (firstRowSpans || hasMultipleHeaderRows || hasImpliedMerge) {
      structure.tableType = 'merged';
      structure.headers = this.extractMergedHeaders(rows);
    } else {
      structure.tableType = 'simple';
      structure.headers = this.extractSimpleHeaders(rows[0] || []);
    }
    
    structure.columnCount = structure.headers.length;
    return structure;
  }
  
  extractMergedHeaders(rows) {
    const headers = [];
    
    if (rows.length < 2) {
      return this.extractSimpleHeaders(rows[0] || []);
    }
    
    const firstRow = rows[0];
    const secondRow = rows[1];
    
    console.log('First row cells:', firstRow.map(c => ({ content: c.content, colspan: c.colspan })));
    console.log('Second row cells:', secondRow.map(c => ({ content: c.content, colspan: c.colspan })));
    
    // 检查是否为隐式合并：第一行比第二行少
    const isImplicitMerge = firstRow.length < secondRow.length;
    console.log('Is implicit merge:', isImplicitMerge, 'first row:', firstRow.length, 'second row:', secondRow.length);
    
    let headerIndex = 0;
    let secondRowIndex = 0;
    
    if (isImplicitMerge) {
      // 隐式合并：第一行的最后一个单元格对应第二行的多个单元格
      for (let i = 0; i < firstRow.length; i++) {
        const firstCell = firstRow[i];
        
        if (i === firstRow.length - 1) {
          // 最后一个单元格，对应剩余的所有第二行单元格
          const parentName = firstCell.content || `合并列${i + 1}`;
          const remainingSecondRowCells = secondRow.length - secondRowIndex;
          
          console.log(`Processing implicit merge: parent="${parentName}", remaining cells: ${remainingSecondRowCells}, starting from index: ${secondRowIndex}`);
          
          // 将剩余的第二行单元格都作为这个合并表头的子列
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
          // 前面的单独列，一一对应 - 但不要跳过第二行的单元格
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
          // 注意：在隐式合并中，第一行前面的列不对应第二行的单元格
          // 第二行的所有单元格都属于最后一个合并表头
        }
      }
    } else {
      // 显式合并或常规处理
      for (let i = 0; i < firstRow.length; i++) {
        const firstCell = firstRow[i];
        
        if (firstCell.colspan > 1) {
          // 显式合并表头单元格，需要根据colspan生成对应数量的子列
          const parentName = firstCell.content || `合并列${i + 1}`;
          
          // 从第二行获取对应的子列名称
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
          // 单独的列
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
          
          // 如果有对应的第二行单元格，跳过一个
          if (secondRowIndex < secondRow.length) {
            secondRowIndex++;
          }
        }
      }
    }
    
    // 如果第二行还有剩余的单元格，可能是表格结构解析不完整，补充处理
    while (secondRowIndex < secondRow.length) {
      const remainingCell = secondRow[secondRowIndex];
      headers.push({
        index: headerIndex++,
        name: remainingCell.content || `额外列${headerIndex}`,
        originalName: remainingCell.content,
        parentHeader: null,
        dataType: 'FIXED',
        value: '',
        sqlQuery: '',
        dataSourceId: null
      });
      secondRowIndex++;
    }
    
    console.log('Generated headers:', headers.map(h => ({ name: h.name, parentHeader: h.parentHeader })));
    
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
}

module.exports = WordParser;