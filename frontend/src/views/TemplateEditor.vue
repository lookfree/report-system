<template>
  <div class="template-editor-page">
    <!-- 顶部工具栏 -->
    <div class="editor-toolbar">
      <div class="toolbar-group">
        <el-button type="primary" @click="saveTemplate">
          <el-icon><DocumentChecked /></el-icon>
          保存模板
        </el-button>
        <el-button @click="previewTemplate">
          <el-icon><View /></el-icon>
          预览
        </el-button>
      </div>
      
      <el-divider direction="vertical" />
      
      <!-- 格式化工具组 -->
      <div class="toolbar-group">
        <el-tooltip content="粗体" placement="bottom">
          <el-button size="small" @click="formatText('bold')">
            <strong>B</strong>
          </el-button>
        </el-tooltip>
        <el-tooltip content="斜体" placement="bottom">
          <el-button size="small" @click="formatText('italic')">
            <em>I</em>
          </el-button>
        </el-tooltip>
        <el-tooltip content="下划线" placement="bottom">
          <el-button size="small" @click="formatText('underline')">
            <u>U</u>
          </el-button>
        </el-tooltip>
      </div>
      
      <el-divider direction="vertical" />
      
      <!-- 对齐工具组 -->
      <div class="toolbar-group">
        <el-tooltip content="左对齐" placement="bottom">
          <el-button size="small" @click="alignText('left')">
            <el-icon><Operation /></el-icon>
          </el-button>
        </el-tooltip>
        <el-tooltip content="居中对齐" placement="bottom">
          <el-button size="small" @click="alignText('center')">
            <el-icon><Aim /></el-icon>
          </el-button>
        </el-tooltip>
        <el-tooltip content="右对齐" placement="bottom">
          <el-button size="small" @click="alignText('right')">
            <el-icon><Right /></el-icon>
          </el-button>
        </el-tooltip>
        <el-tooltip content="两端对齐" placement="bottom">
          <el-button size="small" @click="alignText('justify')">
            <el-icon><Grid /></el-icon>
          </el-button>
        </el-tooltip>
      </div>
      
      <el-divider direction="vertical" />
      
      <!-- 缩进工具组 -->
      <div class="toolbar-group">
        <el-tooltip content="减少缩进" placement="bottom">
          <el-button size="small" @click="adjustIndent('decrease')">
            <el-icon><Back /></el-icon>
          </el-button>
        </el-tooltip>
        <el-tooltip content="增加缩进" placement="bottom">
          <el-button size="small" @click="adjustIndent('increase')">
            <el-icon><Right /></el-icon>
          </el-button>
        </el-tooltip>
      </div>
      
      <el-divider direction="vertical" />
      
      <!-- 字体大小控制组 -->
      <div class="toolbar-group font-size-group">
        <el-tooltip content="减小字体" placement="bottom">
          <el-button size="small" @click="adjustFontSize('decrease')">
            A⁻
          </el-button>
        </el-tooltip>
        <el-select 
          v-model="currentFontSize" 
          @change="changeFontSize" 
          size="small" 
          style="width: 80px;"
          placeholder="字号"
        >
          <el-option label="9pt" value="9pt" />
          <el-option label="10pt" value="10pt" />
          <el-option label="10.5pt" value="10.5pt" />
          <el-option label="11pt" value="11pt" />
          <el-option label="12pt" value="12pt" />
          <el-option label="14pt" value="14pt" />
          <el-option label="16pt" value="16pt" />
          <el-option label="18pt" value="18pt" />
          <el-option label="20pt" value="20pt" />
          <el-option label="24pt" value="24pt" />
        </el-select>
        <el-tooltip content="增大字体" placement="bottom">
          <el-button size="small" @click="adjustFontSize('increase')">
            A⁺
          </el-button>
        </el-tooltip>
      </div>
      
      <el-divider direction="vertical" />
      
      <div class="toolbar-group">
        <el-button @click="insertData">
          <el-icon><Connection /></el-icon>
          插入数据
        </el-button>
      </div>
      
      <el-divider direction="vertical" />
      
      <div class="toolbar-group">
        <el-button @click="exportToWord">
          <el-icon><Download /></el-icon>
          导出Word
        </el-button>
      </div>
    </div>

    <!-- Word 编辑器 -->
    <div class="editor-container">
      <div 
        id="word-editor" 
        contenteditable="true"
        @input="handleContentChange"
        style="height: calc(100vh - 200px); border: 1px solid #ddd; padding: 40px 60px; background: white; overflow-y: auto; line-height: 1.6; max-width: 21cm; margin: 0 auto; box-shadow: 0 0 10px rgba(0,0,0,0.05); font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', '微软雅黑', 'SimSun', '宋体', serif; font-size: 12pt;"
      ></div>
    </div>

    <!-- 数据插入配置对话框 -->
    <el-dialog v-model="showFieldDialog" title="插入数据" width="800px">
      <el-form :model="fieldForm" label-width="100px">
        <el-form-item label="插入类型">
          <el-radio-group v-model="fieldForm.insertType">
            <el-radio value="FIELD">展示字段</el-radio>
            <el-radio value="TABLE">数据表格</el-radio>
          </el-radio-group>
        </el-form-item>
        
        <!-- 展示字段配置 -->
        <template v-if="fieldForm.insertType === 'FIELD'">
          <el-form-item label="字段名称">
            <el-input v-model="fieldForm.name" placeholder="例如：审计单位" />
          </el-form-item>
          <el-form-item label="数据类型">
            <el-select v-model="fieldForm.dataType" style="width: 100%">
              <el-option label="固定值" value="FIXED" />
              <el-option label="日期变量" value="DATE" />
            </el-select>
          </el-form-item>
        </template>
        
        <!-- 数据表格配置 -->
        <template v-if="fieldForm.insertType === 'TABLE'">
          <el-form-item label="表格标题">
            <el-input v-model="fieldForm.tableTitle" placeholder="例如：审计结果汇总表" />
          </el-form-item>
          <el-form-item label="数据源">
            <el-select v-model="fieldForm.dataSourceId" style="width: 100%" @change="onDataSourceChange">
              <el-option 
                v-for="ds in dataSources" 
                :key="ds.id" 
                :label="ds.name" 
                :value="ds.id" 
              />
            </el-select>
          </el-form-item>
        </template>
        <el-form-item label="默认值" v-if="fieldForm.dataType === 'FIXED'">
          <el-input v-model="fieldForm.defaultValue" />
        </el-form-item>
        <el-form-item label="日期格式" v-if="fieldForm.dataType === 'DATE'">
          <el-select v-model="fieldForm.dateFormat" style="width: 100%">
            <el-option label="年-月-日 (2025-09-11)" value="YYYY-MM-DD" />
            <el-option label="年年月月日 (2025年09月11日)" value="YYYY年MM月DD日" />
            <el-option label="月份 (2025年9月)" value="YYYY年M月" />
            <el-option label="年份 (2025年)" value="YYYY年" />
            <el-option label="当前月 (9月)" value="M月" />
            <el-option label="上个月 (8月)" value="PREV_MONTH" />
            <el-option label="下个月 (10月)" value="NEXT_MONTH" />
          </el-select>
        </el-form-item>
        <el-form-item label="数据源" v-if="fieldForm.dataType === 'DATASET'">
          <el-select v-model="fieldForm.dataSourceId" style="width: 100%" @change="onDataSourceChange">
            <el-option 
              v-for="ds in dataSources" 
              :key="ds.id" 
              :label="ds.name" 
              :value="ds.id" 
            />
          </el-select>
        </el-form-item>
        <el-form-item label="数据集字段" v-if="fieldForm.dataType === 'DATASET' && fieldForm.dataSourceId">
          <el-select v-model="fieldForm.datasetField" style="width: 100%">
            <el-option 
              v-for="field in datasetFields" 
              :key="field.value" 
              :label="field.label" 
              :value="field.value" 
            />
          </el-select>
        </el-form-item>
        
        <!-- 数据集表格列配置 -->
        <template v-if="fieldForm.insertType === 'TABLE' && fieldForm.dataType === 'DATASET' && fieldForm.dataSourceId">
          <el-form-item label="表格列配置">
            <div style="max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 10px;">
              <el-checkbox-group v-model="fieldForm.selectedColumns">
                <div v-for="field in datasetFields" :key="field.value" style="margin-bottom: 8px;">
                  <el-checkbox :value="field.value" style="width: 100%;">
                    <span>{{ field.label }} ({{ field.value }})</span>
                  </el-checkbox>
                </div>
              </el-checkbox-group>
            </div>
          </el-form-item>
          
          <!-- 显示预览数据 -->
          <el-form-item label="数据预览">
            <div style="max-height: 200px; overflow: auto; border: 1px solid #ddd; padding: 10px; background: #f9f9f9;">
              <el-table :data="previewData" size="small" style="width: 100%">
                <el-table-column 
                  v-for="field in datasetFields.filter(f => fieldForm.selectedColumns.includes(f.value))" 
                  :key="field.value"
                  :prop="field.value"
                  :label="field.label"
                  min-width="120"
                />
              </el-table>
            </div>
          </el-form-item>
        </template>
      </el-form>
      <template #footer>
        <el-button @click="showFieldDialog = false">取消</el-button>
        <el-button type="primary" @click="confirmInsertField">确定</el-button>
      </template>
    </el-dialog>

    <!-- 动态表格配置对话框 -->
    <el-dialog v-model="showTableDialog" title="配置动态表格" width="900px">
      <div class="table-builder">
        <el-form :model="tableForm" label-width="100px">
          <el-form-item label="表格标题">
            <el-input v-model="tableForm.title" placeholder="例如：审计结果汇总表" />
          </el-form-item>
          <el-form-item label="数据模式">
            <el-select v-model="tableForm.dataMode" style="width: 100%" @change="onDataModeChange">
              <el-option label="SQL查询" value="SQL" />
              <el-option label="数据集字段" value="DATASET" />
            </el-select>
          </el-form-item>
          <el-form-item label="数据源">
            <el-select 
              v-model="tableForm.dataSourceId" 
              style="width: 100%" 
              @change="onDataSourceChange"
            >
              <el-option 
                v-for="ds in dataSources" 
                :key="ds.id" 
                :label="ds.name" 
                :value="ds.id" 
              />
            </el-select>
          </el-form-item>
          <el-form-item label="SQL查询" v-if="tableForm.dataMode === 'SQL'">
            <el-input 
              v-model="tableForm.sqlQuery" 
              type="textarea" 
              :rows="4"
              placeholder="SELECT * FROM table WHERE ..." 
            />
          </el-form-item>
          
          <!-- 数据集字段配置 -->
          <div v-if="tableForm.dataMode === 'DATASET'" class="dataset-config">
            <el-form-item label="表格列配置">
              <div class="columns-config">
                <div 
                  v-for="(col, index) in tableForm.columns" 
                  :key="index" 
                  class="column-item"
                >
                  <el-input 
                    v-model="col.title" 
                    placeholder="列标题" 
                    style="width: 120px; margin-right: 10px;"
                  />
                  <el-select 
                    v-model="col.field" 
                    placeholder="选择字段"
                    style="width: 150px; margin-right: 10px;"
                  >
                    <el-option 
                      v-for="field in availableFields" 
                      :key="field.value" 
                      :label="field.label" 
                      :value="field.value"
                    />
                  </el-select>
                  <el-button 
                    type="danger" 
                    size="small" 
                    @click="removeColumn(index)"
                    :disabled="tableForm.columns.length === 1"
                  >
                    删除
                  </el-button>
                </div>
                <el-button 
                  type="primary" 
                  size="small" 
                  @click="addColumn"
                  style="margin-top: 10px;"
                >
                  添加列
                </el-button>
              </div>
            </el-form-item>
          </div>
        </el-form>
        
        <!-- 表格预览 -->
        <div class="table-preview" v-if="tableForm.title">
          <h4>表格预览</h4>
          <table class="config-table">
            <thead>
              <tr>
                <th v-for="col in tableForm.columns" :key="col.title">
                  {{ col.title || '未命名列' }}
                </th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td v-for="col in tableForm.columns" :key="col.title">
                  {{ getFieldDisplay(col.field) }}
                </td>
              </tr>
              <tr>
                <td v-for="col in tableForm.columns" :key="col.title">
                  动态数据...
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
      
      <template #footer>
        <el-button @click="showTableDialog = false">取消</el-button>
        <el-button 
          type="primary" 
          @click="confirmInsertTable"
          :disabled="!canInsertTable"
        >
          插入表格
        </el-button>
      </template>
    </el-dialog>

    <!-- Word导入对话框 -->
    <el-dialog v-model="showImportDialog" title="导入Word文档" width="600px">
      <el-upload
        ref="uploadRef"
        class="upload-demo"
        drag
        :action="uploadUrl"
        :on-success="handleUploadSuccess"
        :on-error="handleUploadError"
        :before-upload="beforeUpload"
        accept=".doc,.docx"
        :limit="1"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">
          将Word文档拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            支持 .doc 和 .docx 格式，文件大小不超过 10MB
          </div>
        </template>
      </el-upload>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, computed } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { 
  DocumentChecked, 
  View, 
  Grid, 
  Connection, 
  Upload, 
  Download,
  UploadFilled,
  Operation,
  Aim,
  Right,
  Back
} from '@element-plus/icons-vue'
import { useRoute, useRouter } from 'vue-router'
import api from '@/api'

const route = useRoute()
const router = useRouter()

// 编辑器内容
const content = ref('')
const templateId = ref(route.params.id || null)
const templateName = ref('')

// 编辑器工具栏状态
const currentFontSize = ref('12pt')

// 对话框状态
const showFieldDialog = ref(false)
const showTableDialog = ref(false)
const showImportDialog = ref(false)

// 数据源
const dataSources = ref([])
const datasetFields = ref([]) // 数据集字段列表
const previewData = ref([]) // 预览数据

// 上传配置
const uploadUrl = computed(() => `/api/templates/import-word`)
const uploadRef = ref()

// 字段表单
const fieldForm = reactive({
  insertType: 'FIELD', // FIELD 或 TABLE
  name: '',
  tableTitle: '',
  dataType: 'FIXED',
  dataSourceId: null,
  sqlQuery: '',
  defaultValue: '',
  dateFormat: 'YYYY-MM-DD',
  systemVariable: 'CURRENT_USER',
  datasetField: '',
  selectedColumns: []
})

// 表格表单
const tableForm = reactive({
  title: '',
  dataSourceId: null,
  sqlQuery: '',
  dataMode: 'SQL', // 'SQL' 或 'DATASET'
  columns: [
    { title: '列1', field: '' },
    { title: '列2', field: '' },
    { title: '列3', field: '' }
  ]
})

// 可用字段选项（从后端数据源API获取）
const availableFields = ref([])

// 存储光标位置
let savedRange = null

// 计算属性：是否可以插入表格
const canInsertTable = computed(() => {
  if (!tableForm.title) return false
  if (tableForm.dataMode === 'SQL') {
    return !!(tableForm.dataSourceId && tableForm.sqlQuery)
  } else {
    return tableForm.columns.some(col => col.title && col.field)
  }
})

// 计算属性：格式化字段显示
const getFieldDisplay = (field) => {
  return field ? `{{${field}}}` : '示例数据'
}

// 处理内容变化
const handleContentChange = (event) => {
  content.value = event.target.innerHTML
}

// 初始化编辑器
const initEditor = () => {
  console.log('Initializing simple HTML editor')
  
  // 加载模板内容
  setTimeout(() => {
    loadTemplate()
  }, 100)
}

// 加载模板
const loadTemplate = async () => {
  if (!templateId.value) {
    console.log('No template ID, skipping load')
    return
  }
  
  try {
    console.log('Loading template:', templateId.value)
    const template = await api.getTemplate(templateId.value)
    console.log('Template loaded:', template)
    templateName.value = template.name
    
    let htmlContent = ''
    
    // 如果有HTML内容，加载到编辑器
    if (template.htmlContent) {
      htmlContent = template.htmlContent
      console.log('Found HTML content:', htmlContent.substring(0, 100) + '...')
    } else if (template.structure) {
      // 从结构转换为HTML
      htmlContent = convertStructureToHtml(template.structure)
      console.log('Converted structure to HTML:', htmlContent.substring(0, 100) + '...')
    } else {
      console.log('No content found in template')
      htmlContent = '<p>模板内容为空，请上传Word文档或手动编辑。</p>'
    }
    
    // 直接设置DOM内容，避免光标跳动
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      editorElement.innerHTML = htmlContent
    }
    content.value = htmlContent
    console.log('Template content loaded successfully')
  } catch (error) {
    console.error('Failed to load template:', error)
    ElMessage.error('加载模板失败: ' + (error.response?.data?.error || error.message))
  }
}

// 保存模板
const saveTemplate = async () => {
  try {
    if (!templateId.value) {
      ElMessage.error('模板ID不存在，无法保存')
      return
    }
    
    // 从 HTML 编辑器获取内容
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value
    
    await api.saveTemplateHtml(templateId.value, htmlContent)
    ElMessage.success('模板保存成功')
  } catch (error) {
    ElMessage.error('保存失败: ' + (error.response?.data?.error || error.message))
  }
}

// 预览模板
const previewTemplate = async () => {
  try {
    console.log('🔍 Starting template preview with variable processing...')
    
    // 从编辑器获取内容
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value
    
    if (!templateId.value) {
      ElMessage.warning('请先保存模板后再预览')
      return
    }
    
    // 调用后端API处理动态变量
    const processedResponse = await api.processTemplateVariables(templateId.value, htmlContent)
    
    let processedHtml = htmlContent
    if (processedResponse && processedResponse.processedHtml) {
      processedHtml = processedResponse.processedHtml
      console.log('✅ Variables processed for preview')
    } else {
      console.log('⚠️ No variable processing, using original content')
    }
    
    const previewWindow = window.open('', '_blank')
    previewWindow.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>模板预览 - ${templateName.value || '未命名模板'}</title>
        <style>
          @page {
            size: A4;
            margin: 2.54cm;
          }
          body { 
            font-family: "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimSun", "宋体", "Noto Sans SC", "Source Han Sans SC", Arial, sans-serif !important;
            line-height: 1.6;
            padding: 0;
            max-width: 21cm;
            margin: 0 auto;
            color: #000;
            background: white;
            text-rendering: optimizeLegibility;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
          }
          table { 
            border-collapse: collapse; 
            width: 100%;
            margin: 15px 0;
            border: 1px solid #000;
            table-layout: auto;
          }
          table td, table th { 
            border: 1px solid #000; 
            padding: 8px 12px;
            vertical-align: middle;
            line-height: 1.5;
            text-align: left;
          }
          /* 两列表格的第一列 */
          table.two-column td:first-child {
            width: 20%;
            min-width: 100px;
            white-space: nowrap;
          }
          /* 普通表格不特别处理第一列 */
          h1 {
            font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
            font-size: 18pt;
            font-weight: bold;
            text-align: center;
            margin: 24pt 0 12pt 0;
            letter-spacing: 1px;
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
          p {
            font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
            font-size: 12pt;
            margin: 6pt 0;
            text-indent: 2em;
            text-align: justify;
            line-height: 1.75;
            text-justify: inter-ideograph;
          }
          .dynamic-table h4 {
            margin: 15px 0 10px 0;
            font-size: 14pt;
            font-weight: bold;
          }
          @media print {
            body { margin: 0; }
          }
        </style>
      </head>
      <body>
        <div style="text-align: right; color: #888; font-size: 12px; margin-bottom: 20px;">
          预览时间: ${new Date().toLocaleString('zh-CN')}
        </div>
        ${processedHtml}
      </body>
      </html>
    `)
    previewWindow.document.close()
    
    ElMessage.success('预览已打开，变量已处理为实际值')
    
  } catch (error) {
    console.error('Preview error:', error)
    ElMessage.error('预览失败: ' + (error.response?.data?.error || error.message))
    
    // 如果变量处理失败，至少显示原始内容
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value
    
    const previewWindow = window.open('', '_blank')
    previewWindow.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>模板预览（原始内容）</title>
        <style>
          body { font-family: 'Microsoft YaHei', Arial; padding: 20px; }
          .dynamic-field { background-color: #e3f2fd; padding: 2px 6px; border-radius: 3px; border: 1px dashed #2196f3; }
        </style>
      </head>
      <body>${htmlContent}</body>
      </html>
    `)
    previewWindow.document.close()
  }
}

// 保存当前选区
const saveCurrentSelection = () => {
  const selection = window.getSelection()
  if (selection.rangeCount > 0) {
    savedRange = selection.getRangeAt(0).cloneRange()
  } else {
    savedRange = null
  }
}

// 恢复光标位置
const restoreSelection = () => {
  if (savedRange) {
    const selection = window.getSelection()
    selection.removeAllRanges()
    try {
      selection.addRange(savedRange)
      return savedRange
    } catch (e) {
      console.warn('无法恢复选区，使用默认位置')
      return null
    }
  }
  return null
}

// 插入数据
const insertData = () => {
  // 保存当前光标位置
  saveCurrentSelection()
  
  showFieldDialog.value = true
  fieldForm.insertType = 'FIELD'
  fieldForm.name = ''
  fieldForm.tableTitle = ''
  fieldForm.dataType = 'FIXED'
  fieldForm.dataSourceId = null
  fieldForm.sqlQuery = ''
  fieldForm.defaultValue = ''
  fieldForm.dateFormat = 'YYYY-MM-DD'
  fieldForm.systemVariable = 'CURRENT_USER'
  fieldForm.datasetField = ''
  fieldForm.selectedColumns = []
  datasetFields.value = []
  previewData.value = []
}

// 数据源变化处理
const onDataSourceChange = async () => {
  if (!fieldForm.dataSourceId) {
    datasetFields.value = []
    previewData.value = []
    fieldForm.selectedColumns = []
    return
  }
  
  try {
    const response = await api.getDataSourceFields(fieldForm.dataSourceId)
    datasetFields.value = response.fields || []
    
    // 生成mock预览数据
    generatePreviewData()
  } catch (error) {
    console.error('获取数据源字段失败:', error)
    ElMessage.error('获取数据源字段失败')
    datasetFields.value = []
    previewData.value = []
  }
}

// 生成预览数据
const generatePreviewData = () => {
  if (datasetFields.value.length === 0) {
    previewData.value = []
    return
  }
  
  // 生成3行mock数据
  const mockData = []
  for (let i = 0; i < 3; i++) {
    const row = {}
    datasetFields.value.forEach(field => {
      // 根据字段类型生成不同的mock数据
      switch (field.type) {
        case 'date':
          row[field.value] = `2025-09-${11 + i}`
          break
        case 'string':
        default:
          if (field.value.includes('path')) {
            row[field.value] = `/system/${['backup', 'email', 'deploy'][i]}...`
          } else if (field.value.includes('domain') || field.value.includes('schema')) {
            row[field.value] = ['https', 'https', 'https'][i]
          } else if (field.value.includes('ip') || field.value.includes('address')) {
            row[field.value] = `192.168.2.${195 + i}`
          } else if (field.value.includes('source')) {
            row[field.value] = ['旁路', '旁路', '网关'][i]
          } else if (field.value.includes('method')) {
            row[field.value] = 'post'
          } else {
            row[field.value] = `示例数据${i + 1}`
          }
          break
      }
    })
    mockData.push(row)
  }
  previewData.value = mockData
}

// 确认插入字段
const confirmInsertField = () => {
  if (fieldForm.insertType === 'FIELD') {
    insertFieldElement()
  } else if (fieldForm.insertType === 'TABLE') {
    insertTableElement()
  }
}

// 插入字段元素
const insertFieldElement = () => {
  // 构建字段的数据属性
  let dataAttrs = `data-field-name="${fieldForm.name}" data-field-type="${fieldForm.dataType}"`
  
  if (fieldForm.dataType === 'FIXED') {
    dataAttrs += ` data-default-value="${fieldForm.defaultValue || ''}"`
  } else if (fieldForm.dataType === 'DATE') {
    dataAttrs += ` data-date-format="${fieldForm.dateFormat}"`
  }
  
  const displayName = fieldForm.name || (fieldForm.dataType === 'DATASET' ? fieldForm.datasetField : '字段')
  const fieldHtml = `<span class="dynamic-field" ${dataAttrs} contenteditable="false">{{${displayName}}}</span>`
  
  // 插入到编辑器当前光标位置
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()
    
    // 尝试恢复保存的光标位置
    let range = restoreSelection()
    
    if (!range) {
      // 如果没有保存的位置，获取当前选区
      const selection = window.getSelection()
      if (selection.rangeCount > 0) {
        range = selection.getRangeAt(0)
      } else {
        // 如果没有选区，创建一个在编辑器末尾的range
        range = document.createRange()
        range.selectNodeContents(editorElement)
        range.collapse(false)
      }
    }
    
    // 确保range在编辑器内
    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    // 创建字段元素并插入
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = fieldHtml
    const fieldElement = tempDiv.firstElementChild
    
    // 删除选中内容（如果有的话）
    range.deleteContents()
    
    // 插入字段元素
    range.insertNode(fieldElement)
    
    // 将光标移动到插入的字段后面
    range.setStartAfter(fieldElement)
    range.collapse(true)
    selection.removeAllRanges()
    selection.addRange(range)
    
    // 更新响应式数据
    content.value = editorElement.innerHTML
  }
  
  showFieldDialog.value = false
  ElMessage.success('展示字段已插入到光标位置')
}

// 插入表格元素 (只支持数据集)
const insertTableElement = () => {
  // 表格固定使用DATASET类型
  fieldForm.dataType = 'DATASET'
  let tableHtml = ''
  
  if (fieldForm.dataType === 'DATASET' && fieldForm.selectedColumns.length > 0) {
    // 数据集表格
    const columnsConfig = fieldForm.selectedColumns.map(colValue => {
      const field = datasetFields.value.find(f => f.value === colValue)
      return { title: field?.label || colValue, field: colValue }
    })
    
    tableHtml = `
      <div class="dynamic-table" 
           data-table-title="${fieldForm.tableTitle || '数据表格'}"
           data-source-id="${fieldForm.dataSourceId || ''}"
           data-mode="DATASET"
           data-columns='${JSON.stringify(columnsConfig)}'>
        <h4>${fieldForm.tableTitle || '数据表格'}</h4>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">
          <thead>
            <tr style="background-color: #f5f5f5;">
              ${columnsConfig.map(col => `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${col.title}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${previewData.value.slice(0, 3).map(row => `
              <tr>
                ${columnsConfig.map(col => `<td style="border: 1px solid #ddd; padding: 8px;">${row[col.field] || ''}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `
  } else {
    // SQL查询表格
    tableHtml = `
      <div class="dynamic-table" 
           data-table-title="${fieldForm.tableTitle || '数据表格'}"
           data-source-id="${fieldForm.dataSourceId || ''}"
           data-sql="${fieldForm.sqlQuery || ''}"
           data-mode="SQL">
        <h4>${fieldForm.tableTitle || '数据表格'}</h4>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">
          <thead>
            <tr style="background-color: #f5f5f5;">
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">列1</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">列2</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">列3</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="border: 1px solid #ddd; padding: 8px;">示例数据1</td>
              <td style="border: 1px solid #ddd; padding: 8px;">示例数据2</td>
              <td style="border: 1px solid #ddd; padding: 8px;">示例数据3</td>
            </tr>
          </tbody>
        </table>
      </div>
    `
  }
  
  // 插入到编辑器当前光标位置
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()
    
    // 获取当前选区
    const selection = window.getSelection()
    let range
    
    if (selection.rangeCount > 0) {
      range = selection.getRangeAt(0)
    } else {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    // 确保range在编辑器内
    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    // 创建表格元素
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = tableHtml
    const tableElement = tempDiv.firstElementChild
    
    // 插入表格元素
    range.deleteContents()
    range.insertNode(tableElement)
    
    // 将光标移动到表格后面
    range.setStartAfter(tableElement)
    range.collapse(true)
    selection.removeAllRanges()
    selection.addRange(range)
    
    // 更新响应式数据
    content.value = editorElement.innerHTML
  }
  
  showFieldDialog.value = false
  ElMessage.success('数据表格已插入到光标位置')
}

// 插入动态表格
const insertDynamicTable = () => {
  showTableDialog.value = true
  tableForm.title = ''
  tableForm.dataSourceId = null
  tableForm.sqlQuery = ''
  tableForm.dataMode = 'SQL'
  tableForm.columns = [
    { title: '列1', field: '' },
    { title: '列2', field: '' },
    { title: '列3', field: '' }
  ]
}

// 添加列
const addColumn = () => {
  tableForm.columns.push({ 
    title: `列${tableForm.columns.length + 1}`, 
    field: '' 
  })
}

// 删除列
const removeColumn = (index) => {
  if (tableForm.columns.length > 1) {
    tableForm.columns.splice(index, 1)
  }
}

// 数据模式变化
const onDataModeChange = (mode) => {
  console.log('Data mode changed to:', mode)
  if (mode === 'DATASET') {
    // 重置列配置
    tableForm.columns = [
      { title: '列1', field: '' },
      { title: '列2', field: '' },
      { title: '列3', field: '' }
    ]
    // 获取可用字段
    if (tableForm.dataSourceId) {
      loadDataSourceFields(tableForm.dataSourceId)
    }
  }
}

// 数据源变化时获取字段
const loadDataSourceFields = async (dataSourceId) => {
  if (!dataSourceId) {
    availableFields.value = []
    return
  }
  
  try {
    console.log('🔍 Fetching fields for data source:', dataSourceId)
    // 从后端API获取数据源的字段信息
    const response = await api.getDataSourceFields(dataSourceId)
    
    if (response && response.fields) {
      availableFields.value = response.fields
      console.log('✅ Loaded', response.fields.length, 'fields from backend')
    } else {
      availableFields.value = []
      console.warn('⚠️ No fields returned from backend')
    }
  } catch (error) {
    console.error('获取字段失败:', error)
    ElMessage.error('获取数据源字段失败: ' + (error.response?.data?.error || error.message))
    availableFields.value = []
  }
}

// 确认插入表格
const confirmInsertTable = () => {
  // 构建数据属性
  let dataAttrs = `data-table-title="${tableForm.title}" data-source-id="${tableForm.dataSourceId || ''}" data-mode="${tableForm.dataMode}"`
  
  if (tableForm.dataMode === 'SQL') {
    dataAttrs += ` data-sql="${tableForm.sqlQuery || ''}"`
  } else if (tableForm.dataMode === 'DATASET') {
    dataAttrs += ` data-columns='${JSON.stringify(tableForm.columns)}'`
  }
  
  // 生成表头
  let headersHtml = ''
  tableForm.columns.forEach(col => {
    headersHtml += `<th style="border: 1px solid #ddd; padding: 8px; background-color: #f5f5f5;">${col.title || '未命名列'}</th>`
  })
  
  // 生成示例数据行
  let exampleRowHtml = ''
  tableForm.columns.forEach(col => {
    const cellContent = col.field ? ('{{' + col.field + '}}') : '示例数据'
    exampleRowHtml += `<td style="border: 1px solid #ddd; padding: 8px;">${cellContent}</td>`
  })
  
  let tableHtml = `<div class="dynamic-table" ${dataAttrs}>
    <h4>${tableForm.title}</h4>
    <table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">
      <thead>
        <tr>${headersHtml}</tr>
      </thead>
      <tbody>
        <tr>${exampleRowHtml}</tr>
        <tr>
          ${tableForm.columns.map(() => '<td style="border: 1px solid #ddd; padding: 8px;">动态数据...</td>').join('')}
        </tr>
      </tbody>
    </table>
  </div>`
  
  // 插入表格到编辑器当前光标位置
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()
    
    // 获取当前选区
    const selection = window.getSelection()
    let range
    
    if (selection.rangeCount > 0) {
      range = selection.getRangeAt(0)
    } else {
      // 如果没有选区，创建一个在编辑器末尾的range
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    // 确保range在编辑器内
    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    // 创建表格元素并插入
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = tableHtml
    const tableElement = tempDiv.firstElementChild
    
    // 删除选中内容（如果有的话）
    range.deleteContents()
    
    // 插入表格元素
    range.insertNode(tableElement)
    
    // 将光标移动到插入的表格后面
    range.setStartAfter(tableElement)
    range.collapse(true)
    selection.removeAllRanges()
    selection.addRange(range)
    
    // 更新响应式数据
    content.value = editorElement.innerHTML
  }
  
  showTableDialog.value = false
  ElMessage.success('动态表格已插入')
}

// 导入Word文档
const importFromWord = () => {
  showImportDialog.value = true
}

// 处理文件上传前
const beforeUpload = (file) => {
  const isDoc = file.type === 'application/msword' || 
                file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  const isLt10M = file.size / 1024 / 1024 < 10
  
  if (!isDoc) {
    ElMessage.error('只能上传Word文档!')
    return false
  }
  if (!isLt10M) {
    ElMessage.error('文件大小不能超过10MB!')
    return false
  }
  return true
}

// 处理上传成功
const handleUploadSuccess = (response) => {
  if (response.htmlContent) {
    content.value = response.htmlContent
    // 更新HTML编辑器内容
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      editorElement.innerHTML = response.htmlContent
    }
    ElMessage.success('Word文档导入成功')
    showImportDialog.value = false
  } else {
    ElMessage.error('导入失败：无法解析文档内容')
  }
}

// 处理上传失败
const handleUploadError = (error) => {
  ElMessage.error('文档上传失败: ' + error.message)
}

// 导出为Word
const exportToWord = async () => {
  try {
    // 从 HTML 编辑器获取内容
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value
    
    console.log('🚀 开始导出Word文档，处理动态变量...')
    
    // 调用后端API处理动态变量
    const processedResponse = await api.processTemplateVariables(templateId.value, htmlContent)
    
    if (processedResponse && processedResponse.processedHtml) {
      console.log('✅ 动态变量处理完成')
      htmlContent = processedResponse.processedHtml
    } else {
      console.log('⚠️ 未找到处理后的HTML，使用原始内容')
    }
    
    // 导出处理后的HTML为Word文档
    const response = await api.exportHtmlToWord(templateId.value, htmlContent)
    
    // 创建下载链接
    const blob = new Blob([response], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    const url = window.URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `${templateName.value || '模板'}_${new Date().toISOString().slice(0, 10)}.docx`
    link.click()
    window.URL.revokeObjectURL(url)
    
    ElMessage.success('导出成功，动态变量已更新为最新数据')
  } catch (error) {
    console.error('导出失败:', error)
    ElMessage.error('导出失败: ' + (error.response?.data?.error || error.message))
  }
}

// 将结构转换为HTML
const convertStructureToHtml = (structure) => {
  // 实现结构到HTML的转换
  let html = ''
  
  // 检查是否需要生成封面
  if (structure.title && structure.title.includes('API') && structure.title.includes('报')) {
    html += generateCoverHtml(structure.title)
  } else if (structure.title) {
    html += `<h1>${structure.title}</h1>`
  }
  
  if (structure.sections) {
    structure.sections.forEach(section => {
      html += `<h2>${section.title}</h2>`
      if (section.hasTable && section.tableStructure) {
        html += generateTableHtml(section.tableStructure)
      }
      if (section.contentPreview) {
        html += `<p>${section.contentPreview}</p>`
      }
    })
  }
  
  return html
}

// 生成封面HTML
const generateCoverHtml = (title) => {
  const date = new Date().toLocaleDateString('zh-CN', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  })
  
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
    ">
      <!-- 页面边框 -->
      <div style="
        position: absolute;
        top: 60px;
        left: 60px;
        right: 60px;
        bottom: 60px;
        border: 2px solid #666;
        border-style: none;
      ">
        <!-- 左上角装饰 -->
        <div style="
          position: absolute;
          top: 0;
          left: 0;
          width: 60px;
          height: 60px;
          border-top: 2px solid #666;
          border-left: 2px solid #666;
        "></div>
        <!-- 右上角装饰 -->
        <div style="
          position: absolute;
          top: 0;
          right: 0;
          width: 60px;
          height: 60px;
          border-top: 2px solid #666;
          border-right: 2px solid #666;
        "></div>
        <!-- 左下角装饰 -->
        <div style="
          position: absolute;
          bottom: 0;
          left: 0;
          width: 60px;
          height: 60px;
          border-bottom: 2px solid #666;
          border-left: 2px solid #666;
        "></div>
        <!-- 右下角装饰 -->
        <div style="
          position: absolute;
          bottom: 0;
          right: 0;
          width: 60px;
          height: 60px;
          border-bottom: 2px solid #666;
          border-right: 2px solid #666;
        "></div>
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
        ">省公司 API 接口</h1>
        <h1 style="
          font-size: 42pt;
          font-weight: bold;
          color: #1e88e5;
          font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;
          letter-spacing: 8px;
          margin: 0;
          text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
        ">专题审计日报</h1>
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
        ">（${date}）</p>
      </div>
    </div>
  `
}

// 生成表格HTML
const generateTableHtml = (tableStructure) => {
  let html = '<table class="dynamic-table"><thead><tr>'
  
  tableStructure.headers.forEach(header => {
    html += `<th>${header.name}</th>`
  })
  
  html += '</tr></thead><tbody>'
  
  // 添加示例行
  for (let i = 0; i < 3; i++) {
    html += '<tr>'
    tableStructure.headers.forEach(() => {
      html += '<td class="editable-cell">&nbsp;</td>'
    })
    html += '</tr>'
  }
  
  html += '</tbody></table>'
  return html
}

// 加载数据源
const loadDataSources = async () => {
  try {
    const result = await api.getDataSources()
    dataSources.value = result || []
  } catch (error) {
    console.error('加载数据源失败:', error)
    dataSources.value = []
  }
}

// 格式化文本函数
const formatText = (command) => {
  document.execCommand(command, false, null)
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    content.value = editorElement.innerHTML
  }
}

// 文本对齐函数
const alignText = (alignment) => {
  let command = ''
  switch (alignment) {
    case 'left':
      command = 'justifyLeft'
      break
    case 'center':
      command = 'justifyCenter'
      break
    case 'right':
      command = 'justifyRight'
      break
    case 'justify':
      command = 'justifyFull'
      break
  }
  
  if (command) {
    document.execCommand(command, false, null)
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      content.value = editorElement.innerHTML
    }
  }
}

// 调整缩进函数
const adjustIndent = (direction) => {
  const command = direction === 'increase' ? 'indent' : 'outdent'
  document.execCommand(command, false, null)
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    content.value = editorElement.innerHTML
  }
}

// 更改字体大小函数
const changeFontSize = (size) => {
  const selection = window.getSelection()
  if (selection.rangeCount === 0) {
    ElMessage.warning('请先选中要修改的文字')
    return
  }
  
  const range = selection.getRangeAt(0)
  
  if (range.collapsed) {
    ElMessage.warning('请先选中要修改的文字')
    return
  }
  
  try {
    // 检查是否已经在span标签内
    const selectedText = range.toString()
    if (!selectedText) {
      ElMessage.warning('请先选中要修改的文字')
      return
    }
    
    // 创建新的span元素包装选中的内容
    const span = document.createElement('span')
    span.style.fontSize = size
    span.style.fontFamily = 'inherit' // 保持原有字体
    
    // 尝试包围选中内容
    try {
      range.surroundContents(span)
    } catch (e) {
      // 如果无法直接包围（比如跨越多个元素），则提取内容后插入
      const contents = range.extractContents()
      span.appendChild(contents)
      range.insertNode(span)
    }
    
    // 清除选择并将光标放在修改后的内容后
    selection.removeAllRanges()
    const newRange = document.createRange()
    newRange.setStartAfter(span)
    newRange.collapse(true)
    selection.addRange(newRange)
    
    // 更新内容
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      content.value = editorElement.innerHTML
    }
    
    ElMessage.success(`字体大小已设置为 ${size}`)
    
  } catch (error) {
    console.error('设置字体大小失败:', error)
    ElMessage.error('设置字体大小失败，请重试')
  }
}

// 字体大小增减函数
const adjustFontSize = (direction) => {
  const selection = window.getSelection()
  if (selection.rangeCount === 0 || selection.getRangeAt(0).collapsed) {
    ElMessage.warning('请先选中要修改的文字')
    return
  }
  
  const fontSizes = ['9pt', '10pt', '10.5pt', '11pt', '12pt', '14pt', '16pt', '18pt', '20pt', '24pt']
  let newSize = '12pt' // 默认大小
  
  // 尝试获取当前选中文字的字体大小
  const range = selection.getRangeAt(0)
  const parentElement = range.commonAncestorContainer.nodeType === Node.TEXT_NODE 
    ? range.commonAncestorContainer.parentElement 
    : range.commonAncestorContainer
  
  const currentStyle = window.getComputedStyle(parentElement)
  const currentFontSize = currentStyle.fontSize
  
  // 将像素值转换为pt值（粗略转换）
  const pxToPt = (px) => {
    const pxValue = parseFloat(px)
    const ptValue = Math.round(pxValue * 0.75) // 1pt ≈ 1.33px
    return `${ptValue}pt`
  }
  
  let currentPt = currentFontSize.includes('px') ? pxToPt(currentFontSize) : '12pt'
  
  // 在字体大小数组中找到当前大小的索引
  let currentIndex = fontSizes.indexOf(currentPt)
  if (currentIndex === -1) {
    // 如果找不到确切匹配，找最接近的
    const currentNum = parseFloat(currentPt)
    currentIndex = fontSizes.findIndex(size => parseFloat(size) >= currentNum)
    if (currentIndex === -1) currentIndex = fontSizes.length - 1
  }
  
  // 增大或减小字体
  if (direction === 'increase' && currentIndex < fontSizes.length - 1) {
    newSize = fontSizes[currentIndex + 1]
  } else if (direction === 'decrease' && currentIndex > 0) {
    newSize = fontSizes[currentIndex - 1]
  } else {
    ElMessage.info(direction === 'increase' ? '字体已是最大' : '字体已是最小')
    return
  }
  
  // 应用新的字体大小
  currentFontSize.value = newSize
  changeFontSize(newSize)
}

// 增强表格编辑功能
const enhanceTableEditing = () => {
  const editorElement = document.getElementById('word-editor')
  if (!editorElement) return
  
  // 为所有表格添加编辑增强
  const tables = editorElement.querySelectorAll('table')
  tables.forEach(table => {
    // 添加表格工具栏
    if (!table.previousElementSibling?.classList.contains('table-toolbar')) {
      const toolbar = document.createElement('div')
      toolbar.className = 'table-toolbar'
      toolbar.innerHTML = `
        <div class="table-controls">
          <button class="table-btn" onclick="addTableRow(this)" title="添加行">
            <span>+行</span>
          </button>
          <button class="table-btn" onclick="addTableColumn(this)" title="添加列">
            <span>+列</span>
          </button>
          <button class="table-btn" onclick="deleteTableRow(this)" title="删除行">
            <span>-行</span>
          </button>
          <button class="table-btn" onclick="deleteTableColumn(this)" title="删除列">
            <span>-列</span>
          </button>
        </div>
      `
      table.parentNode.insertBefore(toolbar, table)
    }
    
    // 为表格单元格添加右键菜单支持
    const cells = table.querySelectorAll('td, th')
    cells.forEach(cell => {
      cell.addEventListener('contextmenu', showCellContextMenu)
    })
  })
}

// 显示单元格右键菜单
const showCellContextMenu = (event) => {
  event.preventDefault()
  
  // 这里可以添加右键菜单的实现
  // 暂时使用简单的提示
  ElMessage.info('右键菜单功能开发中，可使用工具栏进行表格操作')
}

onMounted(() => {
  console.log('Component mounted, template ID:', templateId.value)
  loadDataSources()
  // 初始化编辑器
  initEditor()
  
  // 延迟增强表格编辑功能
  setTimeout(() => {
    enhanceTableEditing()
  }, 1000)
})

// 全局表格操作函数（用于table-toolbar按钮）
window.addTableRow = function(button) {
  const table = button.closest('.table-toolbar').nextElementSibling
  if (table && table.tagName === 'TABLE') {
    const tbody = table.querySelector('tbody') || table
    const lastRow = tbody.lastElementChild
    if (lastRow) {
      const newRow = lastRow.cloneNode(true)
      // 清空新行的内容
      const cells = newRow.querySelectorAll('td, th')
      cells.forEach(cell => {
        cell.innerHTML = '&nbsp;'
      })
      tbody.appendChild(newRow)
    }
  }
}

window.addTableColumn = function(button) {
  const table = button.closest('.table-toolbar').nextElementSibling
  if (table && table.tagName === 'TABLE') {
    const rows = table.querySelectorAll('tr')
    rows.forEach(row => {
      const lastCell = row.lastElementChild
      if (lastCell) {
        const newCell = document.createElement(lastCell.tagName.toLowerCase())
        newCell.innerHTML = '&nbsp;'
        // 复制样式
        newCell.style.cssText = lastCell.style.cssText
        row.appendChild(newCell)
      }
    })
  }
}

window.deleteTableRow = function(button) {
  const table = button.closest('.table-toolbar').nextElementSibling
  if (table && table.tagName === 'TABLE') {
    const tbody = table.querySelector('tbody') || table
    const rows = tbody.querySelectorAll('tr')
    if (rows.length > 1) { // 保留至少一行
      tbody.removeChild(rows[rows.length - 1])
    }
  }
}

window.deleteTableColumn = function(button) {
  const table = button.closest('.table-toolbar').nextElementSibling
  if (table && table.tagName === 'TABLE') {
    const rows = table.querySelectorAll('tr')
    let maxCells = 0
    rows.forEach(row => {
      maxCells = Math.max(maxCells, row.children.length)
    })
    
    if (maxCells > 1) { // 保留至少一列
      rows.forEach(row => {
        if (row.children.length > 0) {
          row.removeChild(row.lastElementChild)
        }
      })
    }
  }
}
</script>

<style lang="scss">
/* Word 编辑器样式 - 移除scoped以支持动态HTML */
#word-editor {
  /* 基础字体设置 - 优先使用中文字体 */
  font-family: "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimSun", "宋体", "Noto Sans SC", "Source Han Sans SC", "PingFang SC", "Hiragino Sans GB", "Heiti SC", "WenQuanYi Micro Hei", Arial, sans-serif !important;
  font-size: 12pt;
  line-height: 1.6;
  text-rendering: optimizeLegibility;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  
  /* Word表格样式 */
  table, .word-table {
    border-collapse: collapse;
    width: 100%;
    margin: 15px 0;
    border: 1px solid #000;
    table-layout: auto;
    word-wrap: break-word;
    font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
    
    td, th {
      border: 1px solid #000;
      padding: 8px 12px;
      vertical-align: middle;
      text-align: left;
      line-height: 1.5;
      min-height: 30px;
    }
    
    th {
      background-color: transparent;
      font-weight: normal;
      text-align: left;
    }
    
    /* 两列表格特殊处理 */
    &.two-column-table {
      td:first-child {
        width: 20%;
        min-width: 100px;
        white-space: nowrap;
        padding-right: 15px;
      }
      td:last-child {
        width: 80%;
      }
    }
    
    /* 处理合并单元格 */
    td[colspan], th[colspan] {
      text-align: center;
    }
    
    td[rowspan], th[rowspan] {
      vertical-align: middle;
    }
    
    /* 合并单元格样式 */
    td[colspan], th[colspan] {
      text-align: center;
      font-weight: bold;
    }
    
    td[rowspan], th[rowspan] {
      vertical-align: middle;
    }
    
    /* 复杂表格处理 */
    &[data-table-complex="true"] {
      margin: 15px 0;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* 表格内的段落 */
    p {
      margin: 2px 0;
      line-height: 1.4;
    }
    
    /* 表格内的列表 */
    ul, ol {
      margin: 2px 0;
      padding-left: 15px;
    }
    
    /* 长表格处理 */
    &.page-break-table {
      page-break-inside: auto;
      
      thead {
        display: table-header-group;
      }
      
      tbody {
        display: table-row-group;
      }
      
      tr {
        page-break-inside: avoid;
      }
    }
  }
  
  /* Word标题样式 - 优化中文字体显示 */
  h1 {
    font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
    font-size: 18pt;
    font-weight: bold;
    margin: 24pt 0 12pt 0;
    color: #000;
    text-align: center;
    letter-spacing: 1px;
    line-height: 1.3;
  }
  
  h2 {
    font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
    font-size: 16pt;
    font-weight: bold;
    margin: 18pt 0 6pt 0;
    color: #000;
    line-height: 1.3;
  }
  
  h3 {
    font-family: "SimHei", "黑体", "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", sans-serif !important;
    font-size: 14pt;
    font-weight: bold;
    margin: 12pt 0 6pt 0;
    color: #000;
    line-height: 1.3;
  }
  
  h4 {
    font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
    font-size: 12pt;
    font-weight: bold;
    margin: 10pt 0 6pt 0;
    color: #000;
    line-height: 1.3;
  }
  
  /* Word段落样式 - 优化中文字体显示 */
  p {
    font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;
    font-size: 12pt;
    margin: 6pt 0;
    line-height: 1.75;
    text-align: justify;
    text-indent: 2em;
    color: #000;
    text-justify: inter-ideograph;
  }
  
  /* Word列表样式 */
  ul, ol {
    margin: 8px 0;
    padding-left: 20px;
    
    li {
      margin: 4px 0;
      line-height: 1.6;
    }
  }
  
  /* 强调文本样式 */
  strong, b {
    font-weight: bold;
  }
  
  em, i {
    font-style: italic;
  }
  
  u {
    text-decoration: underline;
  }
  
  /* Word字体样式 */
  .font-songti {
    font-family: "宋体", "SimSun", serif;
  }
  
  .font-heiti {
    font-family: "黑体", "SimHei", sans-serif;
  }
  
  .font-yahei {
    font-family: "微软雅黑", "Microsoft YaHei", sans-serif;
  }
  
  /* Word对齐样式 */
  .text-left {
    text-align: left;
  }
  
  .text-center {
    text-align: center;
  }
  
  .text-right {
    text-align: right;
  }
  
  .text-justify {
    text-align: justify;
  }
  
  /* Word间距样式 */
  .line-height-1 {
    line-height: 1;
  }
  
  .line-height-15 {
    line-height: 1.5;
  }
  
  .line-height-2 {
    line-height: 2;
  }
  
  /* 图片样式 */
  img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 10px auto;
  }
  
  /* 页面布局 */
  .page-break {
    page-break-before: always;
    break-before: page;
  }
  
  /* 表格工具栏样式 */
  .table-toolbar {
    background: #f8f9fa;
    padding: 6px 8px;
    border: 1px solid #e9ecef;
    border-bottom: none;
    border-radius: 4px 4px 0 0;
    
    .table-controls {
      display: flex;
      gap: 6px;
    }
    
    .table-btn {
      background: #fff;
      border: 1px solid #dee2e6;
      border-radius: 3px;
      padding: 4px 8px;
      font-size: 11px;
      cursor: pointer;
      color: #495057;
      
      &:hover {
        background-color: #e9ecef;
        border-color: #adb5bd;
      }
      
      &:active {
        background-color: #dee2e6;
        transform: translateY(1px);
      }
    }
  }
  
  /* 增强选中效果 */
  ::selection {
    background-color: #b3d4fc;
    color: #000;
  }
  
  /* 表格编辑增强 */
  table:hover .table-toolbar {
    opacity: 1;
  }
  
  /* 可编辑区域光标样式 */
  [contenteditable="true"] {
    cursor: text;
  }
  
  /* 表格单元格编辑增强 */
  td:focus, th:focus {
    outline: 2px solid #007bff;
    outline-offset: -1px;
    background-color: rgba(0, 123, 255, 0.05);
  }
}
</style>

<style scoped lang="scss">
.template-editor-page {
  height: 100vh;
  display: flex;
  flex-direction: column;
  background: #f5f5f5;
  
  .editor-toolbar {
    background: white;
    padding: 12px 20px;
    border-bottom: 1px solid #e0e0e0;
    display: flex;
    align-items: center;
    gap: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    flex-wrap: wrap;
    
    .toolbar-group {
      display: flex;
      align-items: center;
      gap: 4px;
    }
    
    .el-button {
      border-radius: 4px;
      
      &.is-plain {
        background: #f9f9f9;
      }
      
      &:hover {
        background-color: #e6f7ff;
        border-color: #91d5ff;
      }
    }
    
    .el-divider--vertical {
      height: 24px;
      margin: 0 8px;
    }
    
    .font-size-group {
      .el-button {
        font-weight: bold;
        font-size: 12px;
        min-width: 32px;
        
        &:first-child {
          border-radius: 4px 0 0 4px;
        }
        
        &:last-child {
          border-radius: 0 4px 4px 0;
        }
      }
      
      .el-select {
        margin: 0 -1px;
        
        .el-input {
          .el-input__wrapper {
            border-radius: 0;
            border-left: none;
            border-right: none;
          }
        }
      }
    }
  }
  
  .editor-container {
    flex: 1;
    padding: 20px;
    overflow: auto;
    
    :deep(.ql-container) {
      border-radius: 0 0 8px 8px;
      font-family: 'Microsoft YaHei', '微软雅黑', Arial, sans-serif;
      font-size: 14px;
      line-height: 1.6;
    }
    
    :deep(.ql-toolbar) {
      border-radius: 8px 8px 0 0;
      border-bottom: 1px solid #ccc;
    }
    
    :deep(.ql-editor) {
      min-height: 500px;
    }
    
    :deep(.ql-snow) {
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    :deep(.dynamic-field) {
      background-color: #e3f2fd;
      padding: 2px 6px;
      border-radius: 3px;
      border: 1px dashed #2196f3;
      display: inline-block;
      min-width: 60px;
      cursor: pointer;
    }
    
    :deep(.dynamic-table) {
      border: 2px solid #4caf50;
      position: relative;
      margin: 10px 0;
    }
  }
}

.table-builder {
  .dataset-config {
    margin: 20px 0;
    
    .columns-config {
      .column-item {
        display: flex;
        align-items: center;
        margin-bottom: 10px;
        padding: 10px;
        border: 1px solid #e4e7ed;
        border-radius: 4px;
        background: #fafafa;
        
        &:hover {
          background: #f5f7fa;
        }
      }
    }
  }
  
  .table-preview {
    margin-top: 20px;
    
    h4 {
      margin-bottom: 12px;
      color: #303133;
    }
    
    .config-table {
      width: 100%;
      border-collapse: collapse;
      border: 1px solid #ddd;
      
      th, td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: center;
      }
      
      th {
        background: #f5f5f5;
        font-weight: bold;
      }
      
      td {
        color: #666;
        font-style: italic;
      }
    }
  }
  
  .upload-demo {
    :deep(.el-upload-dragger) {
      padding: 40px;
      
      .el-icon--upload {
        font-size: 67px;
        color: #c0c4cc;
        margin-bottom: 16px;
      }
    }
  }
}
</style>