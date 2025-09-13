const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType } = require('docx');
const fs = require('fs').promises;
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const { Client } = require('pg');

class ReportGenerator {
  constructor(prisma) {
    this.prisma = prisma;
  }
  
  async generateReport(templateId, taskId = null) {
    try {
      // Get template with configs
      const template = await this.prisma.rs_report_templates.findUnique({
        where: { id: templateId },
        include: {
          rs_template_configs: {
            include: {
              rs_data_sources: true
            }
          }
        }
      });
      
      if (!template) {
        throw new Error('Template not found');
      }
      
      // Create report record
      const reportName = `${template.name}_${new Date().toISOString().split('T')[0]}.docx`;
      const reportPath = path.join(process.env.REPORTS_DIR, `${uuidv4()}_${reportName}`);
      
      const report = await this.prisma.rs_generated_reports.create({
        data: {
          templateId,
          taskId,
          reportName,
          filePath: reportPath,
          status: 'PENDING'
        }
      });
      
      try {
        // Generate document content
        const sections = await this.generateSections(template);
        
        // Create Word document
        const doc = new Document({
          sections: [{
            properties: {},
            children: sections
          }]
        });
        
        // Write to file
        const buffer = await Packer.toBuffer(doc);
        await fs.writeFile(reportPath, buffer);
        
        // Update report status
        await this.prisma.rs_generated_reports.update({
          where: { id: report.id },
          data: { status: 'SUCCESS' }
        });
        
        return report;
      } catch (error) {
        // Update report status on error
        await this.prisma.rs_generated_reports.update({
          where: { id: report.id },
          data: { status: 'FAILED' }
        });
        throw error;
      }
    } catch (error) {
      console.error('Report generation error:', error);
      throw error;
    }
  }
  
  async generateSections(template) {
    const sections = [];
    const structure = template.structure;
    
    // Add title
    if (structure.title) {
      sections.push(
        new Paragraph({
          text: structure.title,
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 }
        })
      );
    }
    
    // Process each section
    for (const section of structure.sections) {
      // Add section header
      sections.push(
        new Paragraph({
          text: section.title,
          heading: section.level === 1 ? HeadingLevel.HEADING_1 : 
                   section.level === 2 ? HeadingLevel.HEADING_2 : 
                   HeadingLevel.HEADING_3,
          spacing: { before: 200, after: 200 }
        })
      );
      
      // Get config for this section
      const config = template.rs_template_configs.find(c => c.sectionId === section.id);
      
      if (config) {
        const content = await this.getContentForConfig(config);
        
        if (section.hasTable && Array.isArray(content)) {
          // Generate table with section structure info
          const table = this.generateTable(content, section);
          sections.push(table);
        } else {
          // Generate text content
          sections.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: String(content),
                  size: 24
                })
              ],
              spacing: { after: 200 }
            })
          );
        }
      }
    }
    
    return sections;
  }
  
  async getContentForConfig(config) {
    switch (config.dataType) {
      case 'FIXED':
        return config.value || '';
        
      case 'MANUAL':
        return config.value || '(待填写)';
        
      case 'DYNAMIC':
        if (config.sqlQuery && config.rs_data_sources) {
          return await this.executeSQLQuery(config.sqlQuery, config.rs_data_sources);
        }
        return '(无数据)';
        
      default:
        return '';
    }
  }
  
  async executeSQLQuery(query, dataSource) {
    const client = new Client({
      host: dataSource.host,
      port: dataSource.port,
      database: dataSource.database,
      user: dataSource.username,
      password: dataSource.password
    });
    
    try {
      await client.connect();
      const result = await client.query(query);
      await client.end();
      
      if (result.rows.length === 0) {
        return '(无数据)';
      }
      
      // If single value, return it
      if (result.rows.length === 1 && Object.keys(result.rows[0]).length === 1) {
        return Object.values(result.rows[0])[0];
      }
      
      // Return full result set for tables
      return result.rows;
    } catch (error) {
      console.error('SQL query error:', error);
      return `(查询错误: ${error.message})`;
    }
  }
  
  generateTable(data, templateSection) {
    if (!data || data.length === 0) {
      return new Paragraph({ text: '(无数据)' });
    }
    
    const headers = Object.keys(data[0]);
    const tableStructure = templateSection?.tableStructure;
    
    // Check if it's a merged header table
    if (tableStructure?.tableType === 'merged' && tableStructure.headers) {
      return this.generateMergedTable(data, tableStructure);
    }
    
    // Generate simple table
    return this.generateSimpleTable(data, headers);
  }
  
  generateSimpleTable(data, headers) {
    // Create header row
    const headerRow = new TableRow({
      children: headers.map(header =>
        new TableCell({
          children: [new Paragraph({ 
            text: header,
            alignment: AlignmentType.CENTER
          })],
          shading: { fill: 'E0E0E0' }
        })
      )
    });
    
    // Create data rows
    const dataRows = data.map(row =>
      new TableRow({
        children: headers.map(header =>
          new TableCell({
            children: [new Paragraph({ 
              text: String(row[header] || ''),
              alignment: AlignmentType.LEFT
            })]
          })
        )
      })
    );
    
    return new Table({
      rows: [headerRow, ...dataRows],
      width: { size: 100, type: 'pct' }
    });
  }
  
  generateMergedTable(data, tableStructure) {
    // Group headers by parent
    const headerGroups = {};
    const allHeaders = [];
    
    tableStructure.headers.forEach(header => {
      if (header.parentHeader) {
        if (!headerGroups[header.parentHeader]) {
          headerGroups[header.parentHeader] = [];
        }
        headerGroups[header.parentHeader].push(header);
      } else {
        allHeaders.push({ type: 'single', header });
      }
    });
    
    // Add grouped headers
    Object.keys(headerGroups).forEach(parentName => {
      allHeaders.push({ 
        type: 'group', 
        parentName, 
        children: headerGroups[parentName] 
      });
    });
    
    // Create first header row (parent headers)
    const firstHeaderRow = new TableRow({
      children: allHeaders.map(group => {
        if (group.type === 'single') {
          return new TableCell({
            children: [new Paragraph({ 
              text: group.header.name,
              alignment: AlignmentType.CENTER
            })],
            shading: { fill: 'E0E0E0' }
          });
        } else {
          return new TableCell({
            children: [new Paragraph({ 
              text: group.parentName,
              alignment: AlignmentType.CENTER
            })],
            shading: { fill: 'E0E0E0' },
            columnSpan: group.children.length
          });
        }
      })
    });
    
    // Create second header row (sub headers)
    const secondHeaderRow = new TableRow({
      children: allHeaders.flatMap(group => {
        if (group.type === 'single') {
          return [new TableCell({
            children: [new Paragraph({ text: '' })], // Empty cell for alignment
            shading: { fill: 'F5F5F5' }
          })];
        } else {
          return group.children.map(child => 
            new TableCell({
              children: [new Paragraph({ 
                text: child.name,
                alignment: AlignmentType.CENTER
              })],
              shading: { fill: 'F5F5F5' }
            })
          );
        }
      })
    });
    
    // Create data rows
    const dataRows = data.map(row => {
      const cells = tableStructure.headers.map(header => 
        new TableCell({
          children: [new Paragraph({ 
            text: String(row[header.name] || row[header.originalName] || ''),
            alignment: AlignmentType.LEFT
          })]
        })
      );
      
      return new TableRow({ children: cells });
    });
    
    return new Table({
      rows: [firstHeaderRow, secondHeaderRow, ...dataRows],
      width: { size: 100, type: 'pct' }
    });
  }
}

module.exports = ReportGenerator;