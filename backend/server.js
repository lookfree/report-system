const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const fsPromises = require('fs').promises;
const { v4: uuidv4 } = require('uuid');
require('dotenv').config();

const { PrismaClient } = require('@prisma/client');
const UniversalWordParser = require('./utils/universalWordParser');
const ReportGenerator = require('./utils/reportGenerator');
const TaskScheduler = require('./utils/taskScheduler');
const htmlWordConverter = require('./utils/htmlWordConverter');
const VariableProcessor = require('./utils/variableProcessor');
const datasetConfig = require('./models/DatasetConfig');
const datasetService = require('./utils/datasetService');
const datasetStore = require('./models/DatasetStore');

const app = express();
const prisma = new PrismaClient();
const wordParser = new UniversalWordParser();
const reportGenerator = new ReportGenerator(prisma);
const taskScheduler = new TaskScheduler(prisma, reportGenerator);
const variableProcessor = new VariableProcessor(prisma);

// HTML reconstruction function
function reconstructHtmlFromSections(sections, title = '', originalHtml = '') {
  console.log('🔧 Reconstructing HTML from sections...');
  console.log(`📊 Sections count: ${sections ? sections.length : 0}`);
  console.log(`📄 Title: ${title || 'No title'}`);
  console.log(`🔤 Original HTML length: ${originalHtml ? originalHtml.length : 0}`);
  
  // If we have original HTML, prefer that over section reconstruction
  if (originalHtml && originalHtml.length > 0) {
    console.log('✅ Using original HTML content');
    return originalHtml;
  }
  
  let html = '';
  
  // Add title if provided
  if (title) {
    html += `<h1>${title}</h1>\n`;
  }
  
  // Process each section
  if (sections && Array.isArray(sections)) {
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      console.log(`📄 Processing section ${i}: ${section.title || 'Untitled'}`);
      
      // Try multiple possible content fields
      let sectionHtml = '';
      if (section.rawContent) {
        sectionHtml = section.rawContent;
        console.log(`  ✅ Using rawContent (${sectionHtml.length} chars)`);
      } else if (section.originalHtml) {
        sectionHtml = section.originalHtml;
        console.log(`  ✅ Using originalHtml (${sectionHtml.length} chars)`);
      } else if (section.content) {
        sectionHtml = section.content;
        console.log(`  ✅ Using content (${sectionHtml.length} chars)`);
      } else if (section.html) {
        sectionHtml = section.html;
        console.log(`  ✅ Using html (${sectionHtml.length} chars)`);
      } else if (section.contentPreview) {
        // Use content preview as fallback
        sectionHtml = `<p>${section.contentPreview}</p>`;
        console.log(`  ⚠️ Using contentPreview fallback (${sectionHtml.length} chars)`);
      } else {
        // Create section header at minimum
        if (section.title) {
          sectionHtml = `<h2>${section.title}</h2>`;
          console.log(`  ⚠️ Using title only (${sectionHtml.length} chars)`);
        }
      }
      
      html += sectionHtml;
    }
  }
  
  const resultLength = html.length;
  console.log(`🎯 Reconstruction result: ${resultLength} characters`);
  
  // If we still don't have much content, this indicates a data structure issue
  if (resultLength < 100) {
    console.warn('⚠️ Very short reconstruction result - possible data structure issue');
    console.log('Sample section structure:', sections && sections[0] ? JSON.stringify(sections[0], null, 2) : 'No sections');
  }
  
  return html;
}

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' })); // 增加JSON大小限制
app.use(express.urlencoded({ extended: true, limit: '50mb' })); // 增加URL编码大小限制

// 增加请求超时时间
app.use((req, res, next) => {
  // 为模板上传设置更长的超时时间
  if (req.path === '/api/templates/upload') {
    req.setTimeout(300000); // 5分钟超时
    res.setTimeout(300000);
  }
  next();
});

// Create upload directories
const createDirs = async () => {
  const dirs = [process.env.UPLOAD_DIR, process.env.REPORTS_DIR];
  for (const dir of dirs) {
    await fsPromises.mkdir(dir, { recursive: true });
  }
};

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: async (req, file, cb) => {
    await createDirs();
    cb(null, process.env.UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${uuidv4()}_${file.originalname}`;
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage,
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB 文件大小限制
    fieldSize: 10 * 1024 * 1024  // 10MB 字段大小限制
  },
  fileFilter: (req, file, cb) => {
    // 处理文件名编码
    try {
      if (file.originalname) {
        // 确保文件名是正确的UTF-8编码
        file.originalname = Buffer.from(file.originalname, 'latin1').toString('utf8');
      }
    } catch (error) {
      console.warn('File name encoding issue:', error);
    }
    
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.docx' || ext === '.doc') {
      cb(null, true);
    } else {
      cb(new Error('Only Word documents are allowed'));
    }
  }
});


// Routes

// Upload and parse template
app.post('/api/templates/upload', upload.single('template'), async (req, res) => {
  try {
    const { name } = req.body;
    const file = req.file;
    
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    // 检查文件格式（通过文件头魔数判断）
    const fileBuffer = await fsPromises.readFile(file.path);
    const isRealDocx = fileBuffer[0] === 0x50 && fileBuffer[1] === 0x4B; // PK (ZIP格式标识)
    const isOldDoc = fileBuffer[0] === 0xD0 && fileBuffer[1] === 0xCF; // 旧版doc格式标识
    
    if (isOldDoc) {
      // 删除上传的无效文件
      await fsPromises.unlink(file.path).catch(console.error);
      return res.status(400).json({ 
        error: '文档格式错误：检测到旧版Word格式(.doc)或WPS文档。请使用Microsoft Word另存为.docx格式（Office Open XML格式），或使用WPS Office导出为标准.docx格式后重新上传。' 
      });
    }
    
    if (!isRealDocx) {
      // 删除上传的无效文件
      await fsPromises.unlink(file.path).catch(console.error);
      return res.status(400).json({ 
        error: '文档格式错误：文件不是有效的.docx格式。请确保使用Microsoft Word 2007或更新版本保存，或从WPS导出为标准.docx格式。' 
      });
    }
    
    // Parse the Word template
    const structure = await wordParser.parseTemplate(file.path);
    
    // Debug: Check if originalHtml is preserved
    console.log(`📊 Structure analysis after parsing:`);
    console.log(`  - Title: "${structure.title}"`);
    console.log(`  - Sections count: ${structure.sections ? structure.sections.length : 0}`);
    console.log(`  - Original HTML length: ${structure.originalHtml ? structure.originalHtml.length : 0}`);
    
    if (!structure.originalHtml || structure.originalHtml.length === 0) {
      console.error('⚠️ WARNING: originalHtml is missing or empty after parsing!');
    } else {
      console.log(`✅ Original HTML preserved: ${structure.originalHtml.length} characters`);
    }
    
    // 处理文件名编码问题
    let cleanFileName = file.originalname;
    try {
      // 尝试处理可能的编码问题
      if (Buffer.isBuffer(file.originalname)) {
        cleanFileName = file.originalname.toString('utf8');
      } else if (typeof file.originalname === 'string') {
        // 检查是否包含乱码，如果有则使用用户输入的名称
        const hasGarbledText = /[��]/.test(file.originalname) || 
                              /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/.test(file.originalname) ||
                              /[åèæã]/g.test(file.originalname) || // 常见的UTF-8乱码字符
                              file.originalname.includes('ã') ||
                              file.originalname.includes('æ') ||
                              /[^\x00-\x7F\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u3040-\u309f\u30a0-\u30ff]/.test(file.originalname);
        if (hasGarbledText && name) {
          cleanFileName = name + '.docx';
        }
      }
    } catch (error) {
      console.warn('Error processing filename encoding:', error);
      cleanFileName = name ? name + '.docx' : 'template.docx';
    }
    
    // Create a minimal structure for database storage - only essential metadata
    const tableSections = structure.sections ? structure.sections.filter(s => s.hasTable) : [];
    const minimalStructure = {
      title: structure.title || "",
      totalTables: tableSections.length
    };

    console.log(`Original structure size: ${JSON.stringify(structure).length} characters`);
    console.log(`Minimal structure size: ${JSON.stringify(minimalStructure).length} characters`);
    console.log('Minimal structure summary:', {
      title: minimalStructure.title,
      totalTables: minimalStructure.totalTables
    });
    
    // Store full structure in a separate cache file for configuration access
    const cacheDir = path.join(__dirname, 'cache');
    if (!fs.existsSync(cacheDir)) {
      fs.mkdirSync(cacheDir);
    }
    
    const templateId = `template_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const cacheFilePath = path.join(cacheDir, `${templateId}.json`);
    fs.writeFileSync(cacheFilePath, JSON.stringify(structure));
    console.log(`Full structure cached to: ${cacheFilePath}`);

    // Save minimal template to database
    console.log('🔄 Attempting to save template to database...');
    console.log('Template data:', {
      id: templateId,
      name: name || cleanFileName.replace(/\.(doc|docx)$/i, ''),
      originalFileName: cleanFileName,
      filePath: file.path,
      structureSize: JSON.stringify(minimalStructure).length
    });
    
    // 正确的数据库插入逻辑：
    // 1. 先插入模板基本信息到 rs_report_templates
    // 2. 再插入每个表格配置到 rs_template_configs
    
    console.log('🔄 Inserting template and configurations to database...');
    
    try {
      // 设置较长的超时时间以处理大文件
      const startTime = Date.now();
      console.log(`Starting database operations at ${new Date().toISOString()}`);
      
      // 插入模板基本信息
      const template = await prisma.rs_report_templates.create({
        data: {
          id: templateId,
          name: name || cleanFileName.replace(/\.(doc|docx)$/i, ''),
          originalFileName: cleanFileName,
          filePath: file.path,
          structure: minimalStructure, // 只存储基本信息
          updatedAt: new Date()
        }
      });
      
      console.log('✅ Template basic info saved:', template.id);
      
      // 分批插入表格配置到 rs_template_configs - 优化大文件处理
      const tableSections = structure.sections ? structure.sections.filter(s => s.hasTable) : [];
      console.log(`Processing ${tableSections.length} table sections...`);
      
      const BATCH_SIZE = 10; // 每批处理10个表格
      let totalInserted = 0;
      
      for (let batch = 0; batch < Math.ceil(tableSections.length / BATCH_SIZE); batch++) {
        const startIndex = batch * BATCH_SIZE;
        const endIndex = Math.min(startIndex + BATCH_SIZE, tableSections.length);
        const batchSections = tableSections.slice(startIndex, endIndex);
        
        const configs = batchSections.map((section, index) => {
          const globalIndex = startIndex + index;
          const configId = `config_${templateId}_${globalIndex}`;
          
          // 只存储基本信息，避免大对象
          const essentialInfo = {
            tableIndex: globalIndex,
            columnCount: section.tableStructure?.columnCount || 0,
            rowCount: section.tableStructure?.rowCount || 0,
            hasHeaders: !!(section.tableStructure?.headers?.length)
          };
          
          return {
            id: configId,
            templateId: templateId,
            sectionId: section.id || `table_${globalIndex}`,
            sectionName: section.title || `表格${globalIndex + 1}`,
            dataType: 'MANUAL', // 默认手动填写
            value: JSON.stringify(essentialInfo), // 只存储基本信息
            sqlQuery: null,
            dataSourceId: null,
            columnIndex: globalIndex,
            parentSectionId: null,
            updatedAt: new Date()
          };
        });
        
        if (configs.length > 0) {
          await prisma.rs_template_configs.createMany({
            data: configs
          });
          totalInserted += configs.length;
          console.log(`✅ Batch ${batch + 1}: Inserted ${configs.length} configurations (Total: ${totalInserted}/${tableSections.length})`);
        }
      }
      
      console.log(`✅ All table configurations saved: ${totalInserted} total`);
      
      const endTime = Date.now();
      const processingTime = (endTime - startTime) / 1000;
      console.log(`⏱️ Total processing time: ${processingTime.toFixed(2)} seconds`);
      
      res.json({
        success: true,
        template,
        configsCount: totalInserted,
        processingTime: processingTime
      });
      
    } catch (dbError) {
      console.error('❌ Database insertion failed:', dbError);
      throw dbError;
    }
    
    return;
    
    // TODO: 修复数据库插入问题后恢复这部分代码
    try {
      const template = await prisma.rs_report_templates.create({
        data: {
          id: templateId,
          name: name || cleanFileName.replace(/\.(doc|docx)$/i, ''),
          originalFileName: cleanFileName,
          filePath: file.path,
          structure: minimalStructure,
          updatedAt: new Date()
        }
      });
      
      console.log('✅ Template saved to database successfully:', template.id);
      
      // 成功保存到数据库
      res.json({
        success: true,
        template
      });
      return; // 确保成功后直接返回
    } catch (dbError) {
      console.error('❌ Database save failed:', dbError);
      console.error('Error name:', dbError.name);
      console.error('Error message:', dbError.message);
      
      // 尝试保存一个更简化的版本
      console.log('🔄 Attempting to save ultra-minimal version...');
      const ultraMinimal = {
        title: structure.title || "",
        totalTables: tableSections.length
      };
      
      try {
        const template = await prisma.rs_report_templates.create({
          data: {
            id: templateId,
            name: name || cleanFileName.replace(/\.(doc|docx)$/i, ''),
            originalFileName: cleanFileName,
            filePath: file.path,
            structure: ultraMinimal,
            updatedAt: new Date()
          }
        });
        console.log('✅ Ultra-minimal template saved successfully:', template.id);
      } catch (secondError) {
        console.error('❌ Even ultra-minimal save failed:', secondError);
        throw dbError; // 抛出原始错误
      }
    }
    
    res.json({
      success: true,
      template
    });
  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get all templates
app.get('/api/templates', async (req, res) => {
  try {
    const templates = await prisma.rs_report_templates.findMany({
      orderBy: { createdAt: 'desc' }
    });
    
    // 在返回前处理可能的编码问题
    const cleanTemplates = templates.map(template => {
      let cleanOriginalFileName = template.originalFileName;
      
      if (cleanOriginalFileName) {
        try {
          // 检查是否包含乱码字符
          const hasGarbledText = /[��]/.test(cleanOriginalFileName) || 
                                /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/.test(cleanOriginalFileName) ||
                                /[åèæã]/g.test(cleanOriginalFileName) || // 常见的UTF-8乱码字符
                                cleanOriginalFileName.includes('ã') ||
                                cleanOriginalFileName.includes('æ') ||
                                /[^\x00-\x7F\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u3040-\u309f\u30a0-\u30ff]/.test(cleanOriginalFileName);
          
          if (hasGarbledText) {
            // 如果有乱码，使用模板名称 + .docx
            cleanOriginalFileName = template.name + '.docx';
          }
        } catch (error) {
          console.warn('Error processing template filename:', error);
          cleanOriginalFileName = template.name + '.docx';
        }
      }
      
      return {
        ...template,
        originalFileName: cleanOriginalFileName
      };
    });
    
    res.json(cleanTemplates);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get template with configs
app.get('/api/templates/:id', async (req, res) => {
  try {
    const template = await prisma.rs_report_templates.findUnique({
      where: { id: req.params.id },
      include: {
        rs_template_configs: {
          include: {
            rs_data_sources: true
          }
        }
      }
    });
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // 尝试加载缓存的HTML内容
    let htmlContent = null;
    if (template.id) {
      const cacheFile = path.join(__dirname, 'cache', `${template.id}.json`);
      try {
        if (fs.existsSync(cacheFile)) {
          const cacheData = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
          if (cacheData.htmlContent) {
            htmlContent = cacheData.htmlContent;
            console.log(`Loaded HTML content for template ${template.id}, length: ${htmlContent.length}`);
          } else if (cacheData.sections && Array.isArray(cacheData.sections)) {
            // 重建HTML内容从sections数据，优先使用原始HTML
            htmlContent = reconstructHtmlFromSections(cacheData.sections, cacheData.title, cacheData.originalHtml);
            console.log(`Reconstructed HTML content for template ${template.id}, length: ${htmlContent.length}`);
          }
        }
      } catch (error) {
        console.error('Error loading cached HTML:', error);
      }
    }
    
    res.json({
      ...template,
      htmlContent
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get full template structure from cache for configuration
app.get('/api/templates/:id/full-structure', async (req, res) => {
  try {
    const cacheFilePath = path.join(__dirname, 'cache', `${req.params.id}.json`);
    
    if (!fs.existsSync(cacheFilePath)) {
      return res.status(404).json({ error: 'Full structure cache not found' });
    }
    
    const fullStructure = JSON.parse(fs.readFileSync(cacheFilePath, 'utf8'));
    res.json(fullStructure);
  } catch (error) {
    console.error('Error loading full structure:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get template HTML preview
app.get('/api/templates/:id/preview', async (req, res) => {
  try {
    const { id } = req.params;
    
    // 首先从数据库获取模板信息
    const template = await prisma.rs_report_templates.findUnique({
      where: { id }
    });
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // 尝试从缓存获取完整结构（包含HTML）
    const cacheFilePath = path.join(__dirname, 'cache', `${id}.json`);
    
    if (fs.existsSync(cacheFilePath)) {
      const fullStructure = JSON.parse(fs.readFileSync(cacheFilePath, 'utf8'));
      if (fullStructure.originalHtml) {
        return res.json({
          success: true,
          html: fullStructure.originalHtml,
          title: fullStructure.title || '未知文档'
        });
      }
    }
    
    // 如果缓存不存在或没有HTML，重新解析文件
    if (fs.existsSync(template.filePath)) {
      console.log('Re-parsing template for HTML preview:', template.filePath);
      const structure = await wordParser.parseTemplate(template.filePath);
      
      res.json({
        success: true,
        html: structure.originalHtml || '<p>无法生成HTML预览</p>',
        title: structure.title || template.name
      });
    } else {
      res.status(404).json({ error: 'Template file not found' });
    }
  } catch (error) {
    console.error('Error getting template preview:', error);
    res.status(500).json({ error: error.message });
  }
});

// Delete template
app.delete('/api/templates/:id', async (req, res) => {
  try {
    const { id } = req.params;
    
    // Check if template exists
    const template = await prisma.rs_report_templates.findUnique({
      where: { id }
    });
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // Delete the template (cascade will handle related records)
    await prisma.rs_report_templates.delete({
      where: { id }
    });
    
    // Delete the physical file
    try {
      await fsPromises.unlink(template.filePath);
    } catch (fileError) {
      console.warn('Could not delete template file:', fileError.message);
    }
    
    res.json({ success: true, message: 'Template deleted successfully' });
  } catch (error) {
    console.error('Delete template error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Duplicate template
app.post('/api/templates/:id/duplicate', async (req, res) => {
  try {
    const { id } = req.params;
    const { name } = req.body;
    
    // Get original template with configs
    const originalTemplate = await prisma.rs_report_templates.findUnique({
      where: { id },
      include: {
        rs_template_configs: true
      }
    });
    
    if (!originalTemplate) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // Copy the physical file
    const originalPath = originalTemplate.filePath;
    const newFileName = `${uuidv4()}_${originalTemplate.originalFileName}`;
    const newFilePath = path.join(process.env.UPLOAD_DIR, newFileName);
    
    try {
      await fsPromises.copyFile(originalPath, newFilePath);
    } catch (fileError) {
      return res.status(500).json({ error: 'Failed to copy template file' });
    }
    
    // Create new template
    const newTemplateId = uuidv4();
    const newTemplate = await prisma.rs_report_templates.create({
      data: {
        id: newTemplateId,
        name,
        originalFileName: originalTemplate.originalFileName,
        filePath: newFilePath,
        structure: originalTemplate.structure,
        updatedAt: new Date()
      }
    });
    
    // Copy all configs
    if (originalTemplate.rs_template_configs && originalTemplate.rs_template_configs.length > 0) {
      await prisma.rs_template_configs.createMany({
        data: originalTemplate.rs_template_configs.map(config => ({
          templateId: newTemplate.id,
          sectionId: config.sectionId,
          sectionName: config.sectionName,
          dataType: config.dataType,
          value: config.value,
          sqlQuery: config.sqlQuery,
          dataSourceId: config.dataSourceId,
          columnIndex: config.columnIndex,
          parentSectionId: config.parentSectionId
        }))
      });
    }
    
    res.json({
      success: true,
      template: newTemplate,
      message: 'Template duplicated successfully'
    });
  } catch (error) {
    console.error('Duplicate template error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Update template
app.put('/api/templates/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, description, workflow, selectedProcess, emailAddress, structure } = req.body;
    
    console.log(`🔧 更新模板 - 模板ID: ${id}`);
    console.log(`📝 更新数据:`, { name, description, workflow, selectedProcess, emailAddress, structure: structure ? 'Yes' : 'No' });
    
    // Prepare update data
    const updateData = {
      updatedAt: new Date()
    };
    
    // Only update fields that are provided
    if (name !== undefined) updateData.name = name;
    if (description !== undefined) updateData.description = description;
    if (workflow !== undefined) updateData.workflow = workflow;
    if (selectedProcess !== undefined) updateData.selectedProcess = selectedProcess;
    if (emailAddress !== undefined) updateData.emailAddress = emailAddress;
    if (structure !== undefined) updateData.structure = structure;
    
    // Update template
    const updatedTemplate = await prisma.rs_report_templates.update({
      where: { id },
      data: updateData
    });
    
    console.log(`✅ 模板更新成功: ${updatedTemplate.name}`);
    if (structure) {
      console.log(`📊 模板结构已更新: ${JSON.stringify(structure).substring(0, 100)}...`);
    }
    
    res.json({
      success: true,
      template: updatedTemplate,
      message: 'Template updated successfully'
    });
  } catch (error) {
    console.error('Update template error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Save template configuration
app.post('/api/templates/:id/config', async (req, res) => {
  try {
    const { id } = req.params;
    const { configs, templateStructure } = req.body;
    
    console.log(`🔧 保存模板配置 - 模板ID: ${id}`);
    console.log(`📝 收到配置数量: ${configs ? configs.length : 0}`);
    console.log(`🏗️ 收到模板结构更新: ${templateStructure ? 'Yes' : 'No'}`);
    
    if (configs && configs.length > 0) {
      console.log('前5个配置示例:');
      configs.slice(0, 5).forEach((config, index) => {
        console.log(`  ${index + 1}. ${config.sectionName}: ${config.dataType} = ${config.value || config.sqlQuery || '空'}`);
      });
    }
    
    // Delete existing configs
    const deleteResult = await prisma.rs_template_configs.deleteMany({
      where: { templateId: id }
    });
    console.log(`🗑️ 删除了 ${deleteResult.count} 个旧配置`);
    
    // Get all valid data source IDs
    const validDataSources = await prisma.rs_data_sources.findMany({
      select: { id: true }
    });
    const validDataSourceIds = new Set(validDataSources.map(ds => ds.id));

    // Create new configs with validated dataSourceId
    const newConfigs = await Promise.all(
      configs.map(config => {
        // Validate dataSourceId exists in database, otherwise set to null
        const validatedDataSourceId = config.dataSourceId && validDataSourceIds.has(config.dataSourceId) 
          ? config.dataSourceId 
          : null;

        return prisma.rs_template_configs.create({
          data: {
            id: uuidv4(),
            templateId: id,
            sectionId: config.sectionId,
            sectionName: config.sectionName,
            dataType: config.dataType,
            value: config.value || '',
            sqlQuery: config.sqlQuery || '',
            dataSourceId: validatedDataSourceId,
            columnIndex: config.columnIndex || null,
            parentSectionId: config.parentSectionId || null,
            createdAt: new Date(),
            updatedAt: new Date()
          }
        })
      })
    );
    
    console.log(`✅ 成功保存 ${newConfigs.length} 个新配置`);
    
    // Save template structure if provided
    if (templateStructure) {
      console.log(`🏗️ 更新模板结构...`);
      
      // Update the database structure
      await prisma.rs_report_templates.update({
        where: { id: id },
        data: {
          structure: templateStructure,
          updatedAt: new Date()
        }
      });
      
      // Update the cache file
      const cacheDir = path.join(__dirname, 'cache');
      const cacheFilePath = path.join(cacheDir, `${id}.json`);
      
      // Ensure cache directory exists
      if (!fs.existsSync(cacheDir)) {
        fs.mkdirSync(cacheDir, { recursive: true });
      }
      
      // Write updated structure to cache
      fs.writeFileSync(cacheFilePath, JSON.stringify(templateStructure, null, 2));
      console.log(`💾 已更新缓存文件: ${cacheFilePath}`);
      console.log(`📊 结构信息: sections=${templateStructure.sections ? templateStructure.sections.length : 0}`);
    }
    
    res.json({
      success: true,
      configs: newConfigs
    });
  } catch (error) {
    console.error('❌ 保存配置失败:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// Save template HTML content for TinyMCE editor
app.post('/api/templates/:id/save-html', async (req, res) => {
  try {
    const { id } = req.params;
    const { html } = req.body;
    
    if (!html) {
      return res.status(400).json({ error: 'HTML content is required' });
    }
    
    console.log(`💾 Saving HTML content for template ${id}, length: ${html.length}`);
    
    // Update cache file with new HTML content
    const cacheDir = path.join(__dirname, 'cache');
    const cacheFilePath = path.join(cacheDir, `${id}.json`);
    
    // Ensure cache directory exists
    if (!fs.existsSync(cacheDir)) {
      fs.mkdirSync(cacheDir, { recursive: true });
    }
    
    // Load existing cache or create new one
    let cacheData = {};
    if (fs.existsSync(cacheFilePath)) {
      try {
        cacheData = JSON.parse(fs.readFileSync(cacheFilePath, 'utf8'));
      } catch (error) {
        console.error('Error reading cache file:', error);
      }
    }
    
    // Update HTML content
    cacheData.htmlContent = html;
    cacheData.lastUpdated = new Date().toISOString();
    
    // Save updated cache
    fs.writeFileSync(cacheFilePath, JSON.stringify(cacheData, null, 2));
    
    console.log(`✅ HTML content saved to cache for template ${id}`);
    
    res.json({ 
      success: true,
      message: 'HTML content saved successfully'
    });
  } catch (error) {
    console.error('Error saving HTML content:', error);
    res.status(500).json({ error: error.message });
  }
});

// Data sources management
app.get('/api/datasources', async (req, res) => {
  try {
    const dataSources = await prisma.rs_data_sources.findMany({
      where: { active: true }
    });
    res.json(dataSources);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/datasources', async (req, res) => {
  try {
    const dataSource = await prisma.rs_data_sources.create({
      data: req.body
    });
    res.json(dataSource);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/datasources/test', async (req, res) => {
  try {
    const { host, port, database, username, password, sqlQuery, dataSourceId } = req.body;
    
    // If we have a dataSourceId, try to return mock data first
    if (dataSourceId) {
      try {
        const dataSource = await prisma.rs_data_sources.findUnique({
          where: { id: dataSourceId }
        });
        
        if (dataSource) {
          const { mockSampleData } = require('./scripts/seed-mock-data');
          const sampleData = mockSampleData[dataSource.name];
          
          if (sampleData) {
            console.log(`🎭 返回Mock数据 for ${dataSource.name}:`, sampleData.length, '条记录');
            return res.json({
              success: true,
              message: `Mock数据查询成功 - ${dataSource.name}`,
              rowCount: sampleData.length,
              sample: sampleData.slice(0, 5),
              isMockData: true
            });
          }
        }
      } catch (mockError) {
        console.warn('Mock数据获取失败，尝试真实连接:', mockError.message);
      }
    }
    
    // Fallback to real database connection
    const { Client } = require('pg');
    
    const client = new Client({
      host,
      port,
      database,
      user: username,
      password
    });
    
    await client.connect();
    
    if (sqlQuery) {
      const result = await client.query(sqlQuery);
      await client.end();
      res.json({ 
        success: true, 
        message: 'Connection and query successful',
        rowCount: result.rowCount,
        sample: result.rows.slice(0, 5),
        isMockData: false
      });
    } else {
      await client.end();
      res.json({ success: true, message: 'Connection successful', isMockData: false });
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Scheduled tasks
app.get('/api/tasks', async (req, res) => {
  try {
    const tasks = await prisma.rs_scheduled_tasks.findMany({
      include: {
        rs_report_templates: true
      }
    });
    res.json(tasks);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/tasks', async (req, res) => {
  try {
    const task = await prisma.rs_scheduled_tasks.create({
      data: req.body
    });
    
    // Schedule the task
    if (task.active) {
      taskScheduler.scheduleTask(task);
    }
    
    res.json(task);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.put('/api/tasks/:id/toggle', async (req, res) => {
  try {
    const task = await prisma.rs_scheduled_tasks.findUnique({
      where: { id: req.params.id }
    });
    
    const updatedTask = await prisma.rs_scheduled_tasks.update({
      where: { id: req.params.id },
      data: { active: !task.active }
    });
    
    if (updatedTask.active) {
      taskScheduler.scheduleTask(updatedTask);
    } else {
      taskScheduler.cancelTask(updatedTask.id);
    }
    
    res.json(updatedTask);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Generate report manually
app.post('/api/reports/generate/:templateId', async (req, res) => {
  try {
    const { templateId } = req.params;
    const report = await reportGenerator.generateReport(templateId);
    res.json(report);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get generated reports
app.get('/api/reports', async (req, res) => {
  try {
    const reports = await prisma.rs_generated_reports.findMany({
      include: {
        rs_report_templates: true,
        rs_scheduled_tasks: true
      },
      orderBy: { generatedAt: 'desc' }
    });
    res.json(reports);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Download report
app.get('/api/reports/download/:id', async (req, res) => {
  try {
    const report = await prisma.rs_generated_reports.findUnique({
      where: { id: req.params.id }
    });
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }
    
    res.download(report.filePath, report.reportName);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Fix template descriptions (修复模板描述乱码)
app.post('/api/templates/fix-descriptions', async (req, res) => {
  try {
    const templates = await prisma.rs_report_templates.findMany();
    let fixedCount = 0;
    
    for (const template of templates) {
      if (template.originalFileName) {
        const hasGarbledText = /[��]/.test(template.originalFileName) || 
                              /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/.test(template.originalFileName) ||
                              /[åèæã]/g.test(template.originalFileName) || // 常见的UTF-8乱码字符
                              template.originalFileName.includes('ã') ||
                              template.originalFileName.includes('æ') ||
                              /[^\x00-\x7F\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u3040-\u309f\u30a0-\u30ff]/.test(template.originalFileName);
        
        if (hasGarbledText && template.name) {
          const cleanFileName = template.name + '.docx';
          await prisma.rs_report_templates.update({
            where: { id: template.id },
            data: { originalFileName: cleanFileName }
          });
          fixedCount++;
          console.log(`Fixed template ${template.id}: ${template.originalFileName} -> ${cleanFileName}`);
        }
      }
    }
    
    res.json({
      success: true,
      message: `Fixed ${fixedCount} templates with garbled descriptions`,
      fixedCount
    });
  } catch (error) {
    console.error('Fix descriptions error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Re-parse template structure (重新解析模板结构)
app.post('/api/templates/:id/reparse', async (req, res) => {
  try {
    const { id } = req.params;
    
    // 获取模板信息
    const template = await prisma.rs_report_templates.findUnique({
      where: { id }
    });
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // 重新解析Word文档
    const structure = await wordParser.parseTemplate(template.filePath);
    
    // 更新数据库中的结构信息
    const updatedTemplate = await prisma.rs_report_templates.update({
      where: { id },
      data: { structure }
    });
    
    console.log('Re-parsed template structure:', JSON.stringify(structure, null, 2));
    
    res.json({
      success: true,
      template: updatedTemplate,
      message: 'Template structure re-parsed successfully'
    });
  } catch (error) {
    console.error('Re-parse template error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get all data sources
app.get('/api/data-sources', async (req, res) => {
  try {
    const dataSources = await prisma.rs_data_sources.findMany({
      where: { active: true },
      orderBy: { name: 'asc' }
    });
    res.json(dataSources);
  } catch (error) {
    console.error('Error fetching data sources:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get mock sample data for a data source (for testing)
app.get('/api/datasources/:id/sample', async (req, res) => {
  try {
    const { id } = req.params;
    
    // Get data source info
    const dataSource = await prisma.rs_data_sources.findUnique({
      where: { id }
    });
    
    if (!dataSource) {
      return res.status(404).json({ error: 'Data source not found' });
    }
    
    // Return mock sample data
    const { mockSampleData } = require('./scripts/seed-mock-data');
    const sampleData = mockSampleData[dataSource.name];
    
    if (!sampleData) {
      return res.status(404).json({ error: 'No sample data available for this data source' });
    }
    
    res.json({
      success: true,
      dataSource: {
        id: dataSource.id,
        name: dataSource.name,
        type: dataSource.type
      },
      sampleData,
      count: sampleData.length,
      isMockData: true
    });
  } catch (error) {
    console.error('Error fetching sample data:', error);
    res.status(500).json({ error: error.message });
  }
});

// Import Word document to HTML for TinyMCE editor
app.post('/api/templates/import-word', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    
    if (!file) {
      return res.status(400).json({ error: '请上传Word文档' });
    }
    
    // 验证文档格式
    await htmlWordConverter.validateWordDocument(file.path);
    
    // 转换为HTML
    const result = await htmlWordConverter.wordToHtml(file.path);
    
    // 处理表格使其可编辑
    let processedHtml = htmlWordConverter.makeTablesEditable(result.html);
    
    // 处理合并单元格
    processedHtml = htmlWordConverter.processMergedCells(processedHtml);
    
    // 清理Word特有的标记
    processedHtml = htmlWordConverter.cleanWordHtml(processedHtml);
    
    // 提取表格结构
    const tableStructures = htmlWordConverter.extractTableStructure(processedHtml);
    
    // 删除临时文件
    await fsPromises.unlink(file.path).catch(console.error);
    
    res.json({
      success: true,
      htmlContent: processedHtml,
      tables: tableStructures,
      messages: result.messages
    });
  } catch (error) {
    console.error('Word import error:', error);
    // 删除临时文件
    if (req.file) {
      await fsPromises.unlink(req.file.path).catch(console.error);
    }
    res.status(500).json({ 
      error: error.message || '文档导入失败'
    });
  }
});

// Export HTML to Word document with dataset support
app.post('/api/templates/:id/export-word', async (req, res) => {
  try {
    const { id } = req.params;
    const { htmlContent } = req.body;

    if (!htmlContent) {
      return res.status(400).json({ error: '缺少HTML内容' });
    }

    // 获取模板信息
    const template = await prisma.rs_report_templates.findUnique({
      where: { id }
    });

    if (!template) {
      return res.status(404).json({ error: '模板不存在' });
    }

    // 处理HTML中的数据集占位符
    console.log('📊 Processing datasets for template:', id);
    let processedHtml = htmlContent;

    // 查找所有数据集占位符 {{dataset:id:name}}
    const datasetPattern = /\{\{dataset:(\d+):([^}]+)\}\}/g;
    const matches = [...processedHtml.matchAll(datasetPattern)];

    for (const match of matches) {
      const [placeholder, datasetId, datasetName] = match;
      console.log(`Processing dataset: ${datasetName} (ID: ${datasetId})`);

      try {
        // 获取数据集并执行查询
        const dataset = datasetStore.getDataset(datasetId);
        if (dataset) {
          const rows = await prisma.$queryRawUnsafe(dataset.sqlQuery);
          const dataResult = {
            type: dataset.type,
            fields: dataset.fields,
            data: dataset.type === 'single' ? rows[0] || {} : rows
          };

          const dataHtml = formatDatasetAsHtml(dataResult);
          processedHtml = processedHtml.replace(placeholder, dataHtml);
        } else {
          console.warn(`Dataset not found: ${datasetId}`);
          processedHtml = processedHtml.replace(placeholder, `[数据集未找到: ${datasetName}]`);
        }
      } catch (error) {
        console.error(`Error processing dataset ${datasetId}:`, error);
        processedHtml = processedHtml.replace(placeholder, `[数据集错误: ${error.message}]`);
      }
    }

    // 生成输出文件路径
    const outputFileName = `${template.name}_${Date.now()}.docx`;
    const outputPath = path.join(process.env.REPORTS_DIR, outputFileName);

    // 转换HTML为Word，传递prisma实例用于数据查询
    await htmlWordConverter.htmlToWord(processedHtml, outputPath, prisma);

    // 发送文件
    res.download(outputPath, outputFileName, (err) => {
      if (err) {
        console.error('File download error:', err);
      }
      // 删除临时文件
      fsPromises.unlink(outputPath).catch(console.error);
    });
  } catch (error) {
    console.error('Word export error:', error);
    res.status(500).json({
      error: error.message || '文档导出失败'
    });
  }
});

// Helper function to format dataset as HTML
function formatDatasetAsHtml(datasetResult) {
  if (!datasetResult) return '';

  switch (datasetResult.type) {
    case 'text':
      return datasetResult.value;

    case 'single':
      const values = datasetResult.fields.map(field =>
        `<strong>${field}:</strong> ${datasetResult.data[field] || ''}`
      );
      return values.join('<br>');

    case 'list':
      let html = '<table border="1" style="border-collapse: collapse; width: 100%;">';
      // Add header
      html += '<tr>';
      datasetResult.fields.forEach(field => {
        html += `<th style="padding: 8px; background-color: #f2f2f2;">${field}</th>`;
      });
      html += '</tr>';
      // Add data rows
      datasetResult.data.forEach(row => {
        html += '<tr>';
        datasetResult.fields.forEach(field => {
          html += `<td style="padding: 8px;">${row[field] || ''}</td>`;
        });
        html += '</tr>';
      });
      html += '</table>';
      return html;

    case 'error':
      return `<span style="color: red;">Error: ${datasetResult.message}</span>`;

    default:
      return '';
  }
}

// Save template with HTML content
app.post('/api/templates/:id/save-html', async (req, res) => {
  try {
    const { id } = req.params;
    const { html, htmlContent, structure } = req.body;
    const content = html || htmlContent; // 兼容两种字段名
    
    // 更新模板的HTML内容
    const template = await prisma.rs_report_templates.update({
      where: { id },
      data: {
        htmlContent: content,
        structure: structure ? JSON.stringify(structure) : null,
        updatedAt: new Date()
      }
    });
    
    res.json({
      success: true,
      template
    });
  } catch (error) {
    console.error('Save template HTML error:', error);
    res.status(500).json({ 
      error: error.message || '保存失败'
    });
  }
});

// Get data source fields (for dataset configuration)
app.get('/api/datasources/:id/fields', async (req, res) => {
  try {
    const dataSourceId = parseInt(req.params.id);
    console.log('Getting fields for data source:', dataSourceId);

    // 根据数据源ID返回不同的字段集合
    let mockFields = [];
    
    switch (dataSourceId) {
      case 1: // 活跃接口清单数据源
        mockFields = [
          { label: '路径', value: 'path', type: 'string' },
          { label: '接口模式', value: 'interface_schema', type: 'string' },
          { label: '域名', value: 'domain', type: 'string' },
          { label: '通过来源', value: 'pass_source', type: 'string' },
          { label: '方法', value: 'method', type: 'string' },
          { label: '请求时间', value: 'request_time', type: 'date' },
          { label: '响应状态', value: 'response_status', type: 'number' },
          { label: '接口类型', value: 'interface_type', type: 'string' },
          { label: '安全级别', value: 'security_level', type: 'string' }
        ];
        break;
        
      case 2: // 审计数据库
        mockFields = [
          { label: '审计名称', value: 'audit_name', type: 'string' },
          { label: '审计时间', value: 'audit_date', type: 'date' },
          { label: '风险等级', value: 'risk_level', type: 'string' },
          { label: '责任部门', value: 'department', type: 'string' },
          { label: '整改状态', value: 'rectification_status', type: 'string' },
          { label: '发现问题数', value: 'issue_count', type: 'number' },
          { label: '已整改数', value: 'resolved_count', type: 'number' },
          { label: '整改率', value: 'rectification_rate', type: 'percentage' },
          { label: '风险评分', value: 'risk_score', type: 'number' },
          { label: '审计人员', value: 'auditor', type: 'string' }
        ];
        break;
        
      case 3: // 安全事件库
        mockFields = [
          { label: '事件ID', value: 'incident_id', type: 'string' },
          { label: '事件标题', value: 'incident_title', type: 'string' },
          { label: '发生时间', value: 'occurrence_time', type: 'date' },
          { label: '事件类型', value: 'incident_type', type: 'string' },
          { label: '严重程度', value: 'severity', type: 'string' },
          { label: '影响范围', value: 'impact_scope', type: 'string' },
          { label: '处理状态', value: 'handling_status', type: 'string' },
          { label: '负责人', value: 'responsible_person', type: 'string' },
          { label: '源IP', value: 'source_ip', type: 'string' },
          { label: '目标IP', value: 'target_ip', type: 'string' }
        ];
        break;
        
      default: // 默认字段集合
        mockFields = [
          { label: '路径', value: 'path', type: 'string' },
          { label: '接口模式', value: 'interface_schema', type: 'string' },
          { label: '域名', value: 'domain', type: 'string' },
          { label: '通过来源', value: 'pass_source', type: 'string' },
          { label: '方法', value: 'method', type: 'string' },
          { label: '审计名称', value: 'audit_name', type: 'string' },
          { label: '审计时间', value: 'audit_date', type: 'date' },
          { label: '风险等级', value: 'risk_level', type: 'string' }
        ];
    }

    console.log(`✅ Retrieved ${mockFields.length} fields for data source ${dataSourceId}`);

    res.json({
      success: true,
      fields: mockFields
    });

  } catch (error) {
    console.error('Get data source fields error:', error);
    res.status(500).json({ 
      error: error.message || '获取数据源字段失败'
    });
  }
});

// Dataset Configuration APIs
// 配置单元格数据集
app.post('/api/templates/:templateId/cells/:cellId/dataset', async (req, res) => {
  try {
    const { templateId, cellId } = req.params;
    const config = req.body;

    console.log('📊 Configuring dataset for cell:', cellId, 'in template:', templateId);

    datasetConfig.addConfig(parseInt(templateId), cellId, config);

    res.json({
      success: true,
      message: '数据集配置已保存'
    });
  } catch (error) {
    console.error('Dataset config error:', error);
    res.status(500).json({ error: error.message || '配置保存失败' });
  }
});

// 获取单元格数据集配置
app.get('/api/templates/:templateId/cells/:cellId/dataset', async (req, res) => {
  try {
    const { templateId, cellId } = req.params;
    const config = datasetConfig.getConfig(parseInt(templateId), cellId);

    res.json({
      success: true,
      config: config || null
    });
  } catch (error) {
    console.error('Get dataset config error:', error);
    res.status(500).json({ error: error.message || '获取配置失败' });
  }
});

// 获取模板所有数据集配置
app.get('/api/templates/:templateId/datasets', async (req, res) => {
  try {
    const { templateId } = req.params;
    const configs = datasetConfig.getTemplateConfigs(parseInt(templateId));

    res.json({
      success: true,
      configs: configs
    });
  } catch (error) {
    console.error('Get template datasets error:', error);
    res.status(500).json({ error: error.message || '获取配置失败' });
  }
});

// 获取数据集实际数据（用于预览）
app.get('/api/templates/:templateId/cells/:cellId/dataset-data', async (req, res) => {
  try {
    const { templateId, cellId } = req.params;
    const data = await datasetService.executeDatasetQuery(parseInt(templateId), cellId);

    res.json({
      success: true,
      data: data
    });
  } catch (error) {
    console.error('Get dataset data error:', error);
    res.status(500).json({ error: error.message || '获取数据失败' });
  }
});

// ============ 数据集管理 API ============

// 获取所有数据集
app.get('/api/datasets', async (req, res) => {
  try {
    const datasets = datasetStore.getAllDatasets();
    res.json({
      success: true,
      datasets: datasets
    });
  } catch (error) {
    console.error('Get datasets error:', error);
    res.status(500).json({ error: error.message || '获取数据集失败' });
  }
});

// 创建数据集
app.post('/api/datasets', async (req, res) => {
  try {
    const { name, description, type, sqlQuery, fields } = req.body;

    if (!name || !sqlQuery || !fields || fields.length === 0) {
      return res.status(400).json({ error: '缺少必要参数' });
    }

    const dataset = datasetStore.addDataset({
      name,
      description,
      type,
      sqlQuery,
      fields
    });

    res.json({
      success: true,
      dataset: dataset
    });
  } catch (error) {
    console.error('Create dataset error:', error);
    res.status(500).json({ error: error.message || '创建数据集失败' });
  }
});

// 更新数据集
app.put('/api/datasets/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updates = req.body;

    const dataset = datasetStore.updateDataset(id, updates);
    if (!dataset) {
      return res.status(404).json({ error: '数据集不存在' });
    }

    res.json({
      success: true,
      dataset: dataset
    });
  } catch (error) {
    console.error('Update dataset error:', error);
    res.status(500).json({ error: error.message || '更新数据集失败' });
  }
});

// 删除数据集
app.delete('/api/datasets/:id', async (req, res) => {
  try {
    const { id } = req.params;

    const success = datasetStore.deleteDataset(id);
    if (!success) {
      return res.status(404).json({ error: '数据集不存在' });
    }

    res.json({
      success: true,
      message: '删除成功'
    });
  } catch (error) {
    console.error('Delete dataset error:', error);
    res.status(500).json({ error: error.message || '删除数据集失败' });
  }
});

// 预览数据集数据
app.post('/api/datasets/preview', async (req, res) => {
  try {
    const { sqlQuery, type, fields } = req.body;

    if (!sqlQuery) {
      return res.status(400).json({ error: '缺少SQL查询' });
    }

    // 执行查询
    const rows = await prisma.$queryRawUnsafe(sqlQuery);

    const result = {
      type: type || 'list',
      fields: fields || (rows[0] ? Object.keys(rows[0]) : []),
      data: type === 'single' ? rows[0] || {} : rows
    };

    res.json({
      success: true,
      result: result
    });
  } catch (error) {
    console.error('Preview dataset error:', error);
    res.json({
      success: true,
      result: {
        type: 'error',
        message: error.message
      }
    });
  }
});

// 执行数据集查询 - 连接真实PostgreSQL数据库
app.get('/api/datasets/:id/execute', async (req, res) => {
  try {
    const { id } = req.params;

    const dataset = datasetStore.getDataset(id);
    if (!dataset) {
      return res.status(404).json({ error: '数据集不存在' });
    }

    console.log(`📊 执行数据集查询: ${dataset.name} (ID: ${id})`);
    console.log(`SQL: ${dataset.sqlQuery}`);

    let result;

    try {
      // 尝试执行真实数据库查询
      const rows = await prisma.$queryRawUnsafe(dataset.sqlQuery);
      console.log(`✅ 查询成功，返回 ${rows.length} 条记录`);

      result = {
        type: dataset.type,
        fields: dataset.fields,
        data: dataset.type === 'single' ? rows[0] || {} : rows,
        source: 'database'
      };
    } catch (dbError) {
      console.warn(`⚠️ 数据库查询失败: ${dbError.message}`);
      console.log('📦 使用模拟数据作为备选');

      // 如果数据库查询失败，使用模拟数据
      const mockData = getMockDataForDataset(dataset);
      result = {
        type: dataset.type,
        fields: dataset.fields,
        data: mockData,
        source: 'mock',
        message: '使用模拟数据（数据库连接失败）'
      };
    }

    res.json({
      success: true,
      result: result
    });
  } catch (error) {
    console.error('Execute dataset error:', error);
    res.json({
      success: true,
      result: {
        type: 'error',
        message: error.message
      }
    });
  }
});

// 获取数据集的模拟数据
function getMockDataForDataset(dataset) {
  // 根据数据集名称返回相应的模拟数据
  const mockDataMap = {
    '模板列表': [
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
  return dataset.type === 'single' ? data[0] || {} : data;
}

// Process template variables (for dynamic export)
app.post('/api/templates/:id/process-variables', async (req, res) => {
  try {
    const templateId = parseInt(req.params.id);
    const { html } = req.body;

    console.log('🔄 Processing template variables for template:', templateId);

    if (!html) {
      return res.status(400).json({ error: 'HTML content is required' });
    }

    // Create context for variable processing
    const context = {
      currentUser: '系统管理员',
      department: '信息安全室',
      systemVersion: 'v1.0.0',
      templateId: templateId
    };

    // Process all dynamic variables
    const processedHtml = await variableProcessor.processVariables(html, context);

    console.log('✅ Template variables processed successfully');

    res.json({
      success: true,
      processedHtml: processedHtml
    });

  } catch (error) {
    console.error('Process variables error:', error);
    res.status(500).json({ 
      error: error.message || '变量处理失败'
    });
  }
});

// Initialize
const PORT = process.env.PORT || 3001;

const startServer = async () => {
  try {
    await createDirs();
    await taskScheduler.loadActiveTasks();
    
    app.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
    });
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
};

startServer();

// Graceful shutdown
process.on('SIGINT', async () => {
  console.log('Shutting down gracefully...');
  await prisma.$disconnect();
  process.exit(0);
});