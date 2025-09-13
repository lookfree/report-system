<template>
  <div class="template-editor-page">
    <!-- 顶部工具栏 -->
    <EditorToolbar 
      :current-font-size="currentFontSize"
      @save="saveTemplate"
      @preview="previewTemplate"
      @format-text="formatText"
      @align-text="alignText"
      @adjust-indent="adjustIndent"
      @change-font-size="changeFontSize"
      @adjust-font-size="adjustFontSize"
      @insert-data="insertData"
      @export-word="exportToWord"
      @save-selection="saveSelection"
    />

    <!-- Word 编辑器 -->
    <div class="editor-container">
      <div 
        id="word-editor" 
        contenteditable="true"
        @input="handleContentChange"
        @mouseup="handleTextSelection"
        @keyup="handleTextSelection"
        @keydown="handleKeyDown"
        class="word-editor"
      ></div>
    </div>

    <!-- 插入数据对话框 -->
    <InsertDataDialog
      v-model:visible="showFieldDialog"
      :data-sources="dataSources"
      :dataset-fields="datasetFields"
      :preview-data="previewData"
      @confirm="confirmInsertField"
      @data-source-change="onDataSourceChange"
    />

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
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, computed } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { useRoute, useRouter } from 'vue-router'
import api from '@/api'

// 导入组件和composables
import EditorToolbar from '@/components/EditorToolbar.vue'
import InsertDataDialog from '@/components/InsertDataDialog.vue'
import { useTextFormatting } from '@/composables/useTextFormatting'
import { useTableEditing } from '@/composables/useTableEditing'

const route = useRoute()
const router = useRouter()

// 使用composables
const { 
  currentFontSize,
  formatText,
  alignText,
  adjustIndent,
  changeFontSize,
  adjustFontSize,
  saveSelection,
  restoreSelection
} = useTextFormatting()

const { enhanceTableEditing, setupGlobalTableFunctions } = useTableEditing()

// 基础数据
const content = ref('')
const templateId = ref(route.params.id || null)
const templateName = ref('')

// 对话框状态
const showFieldDialog = ref(false)
const showTableDialog = ref(false)

// 数据源
const dataSources = ref([])
const datasetFields = ref([])
const previewData = ref([])
const availableFields = ref([])

// 表格表单
const tableForm = reactive({
  title: '',
  dataSourceId: null,
  sqlQuery: '',
  dataMode: 'SQL',
  columns: [
    { title: '列1', field: '' },
    { title: '列2', field: '' },
    { title: '列3', field: '' }
  ]
})

// 计算属性：是否可以插入表格
const canInsertTable = computed(() => {
  if (!tableForm.title) return false
  if (tableForm.dataMode === 'SQL') {
    return !!(tableForm.dataSourceId && tableForm.sqlQuery)
  } else {
    return tableForm.columns.some(col => col.title && col.field)
  }
})

// 格式化字段显示
const getFieldDisplay = (field) => {
  return field ? `{{${field}}}` : '示例数据'
}

// 处理内容变化
const handleContentChange = (event) => {
  content.value = event.target.innerHTML
}

// 处理文字选择，自动保存选区
const handleTextSelection = () => {
  // 延迟执行，确保选区已完全建立
  setTimeout(() => {
    const selection = window.getSelection()
    if (selection.rangeCount > 0 && !selection.getRangeAt(0).collapsed) {
      saveSelection()
    }
  }, 10)
}

// 处理键盘快捷键
const handleKeyDown = (event) => {
  // 获取当前光标所在的元素
  const selection = window.getSelection()
  if (selection.rangeCount === 0) return
  
  let node = selection.getRangeAt(0).startContainer
  while (node && node.nodeType !== 1) {
    node = node.parentNode
  }
  
  // 检查是否在表格单元格内
  while (node && !['TD', 'TH'].includes(node.tagName)) {
    node = node.parentNode
  }
  
  if (node && ['TD', 'TH'].includes(node.tagName)) {
    // Tab键：增加缩进（文字右移）
    if (event.key === 'Tab' && !event.shiftKey) {
      event.preventDefault()
      const currentPadding = parseInt(window.getComputedStyle(node).paddingLeft) || 4
      node.style.paddingLeft = (currentPadding + 8) + 'px'
      ElMessage.success('文字右移')
    }
    // Shift+Tab：减少缩进（文字左移）
    else if (event.key === 'Tab' && event.shiftKey) {
      event.preventDefault()
      const currentPadding = parseInt(window.getComputedStyle(node).paddingLeft) || 4
      node.style.paddingLeft = Math.max(2, currentPadding - 8) + 'px'
      ElMessage.success('文字左移')
    }
  }
}

// 初始化编辑器
const initEditor = () => {
  console.log('Initializing editor')
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
    
    if (template.htmlContent) {
      htmlContent = template.htmlContent
      console.log('Found HTML content:', htmlContent.substring(0, 100) + '...')
    } else if (template.structure) {
      htmlContent = convertStructureToHtml(template.structure)
      console.log('Converted structure to HTML:', htmlContent.substring(0, 100) + '...')
    } else {
      console.log('No content found in template')
      htmlContent = '<p>模板内容为空，请上传Word文档或手动编辑。</p>'
    }
    
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
    
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value
    
    if (!templateId.value) {
      ElMessage.warning('请先保存模板后再预览')
      return
    }
    
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
            padding: 2px 4px;
            vertical-align: middle;
            line-height: 1.5;
            text-align: left;
          }
          table.two-column td:first-child {
            width: 20%;
            min-width: 100px;
            white-space: nowrap;
          }
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
  }
}

// 导出为Word
const exportToWord = async () => {
  try {
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value
    
    console.log('🚀 开始导出Word文档，处理动态变量...')
    
    const processedResponse = await api.processTemplateVariables(templateId.value, htmlContent)
    
    if (processedResponse && processedResponse.processedHtml) {
      console.log('✅ 动态变量处理完成')
      htmlContent = processedResponse.processedHtml
    } else {
      console.log('⚠️ 未找到处理后的HTML，使用原始内容')
    }
    
    const response = await api.exportHtmlToWord(templateId.value, htmlContent)
    
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

// 插入数据
const insertData = () => {
  showFieldDialog.value = true
}

// 确认插入字段
const confirmInsertField = (fieldFormData) => {
  if (fieldFormData.insertType === 'FIELD') {
    insertFieldElement(fieldFormData)
  } else if (fieldFormData.insertType === 'TABLE') {
    insertTableElement(fieldFormData)
  }
  showFieldDialog.value = false
}

// 插入字段元素
const insertFieldElement = (fieldForm) => {
  let dataAttrs = `data-field-name="${fieldForm.name}" data-field-type="${fieldForm.dataType}"`
  
  if (fieldForm.dataType === 'FIXED') {
    dataAttrs += ` data-default-value="${fieldForm.defaultValue || ''}"`
  } else if (fieldForm.dataType === 'DATE') {
    dataAttrs += ` data-date-format="${fieldForm.dateFormat}"`
  }
  
  const displayName = fieldForm.name || (fieldForm.dataType === 'DATASET' ? fieldForm.datasetField : '字段')
  const fieldHtml = `<span class="dynamic-field" ${dataAttrs} contenteditable="false">{{${displayName}}}</span>`
  
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()
    
    const selection = window.getSelection()
    let range
    
    if (selection.rangeCount > 0) {
      range = selection.getRangeAt(0)
    } else {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = fieldHtml
    const fieldElement = tempDiv.firstElementChild
    
    range.deleteContents()
    range.insertNode(fieldElement)
    
    range.setStartAfter(fieldElement)
    range.collapse(true)
    selection.removeAllRanges()
    selection.addRange(range)
    
    content.value = editorElement.innerHTML
  }
  
  ElMessage.success('展示字段已插入到光标位置')
}

// 插入表格元素
const insertTableElement = (fieldForm) => {
  fieldForm.dataType = 'DATASET'
  let tableHtml = ''
  
  if (fieldForm.dataType === 'DATASET' && fieldForm.selectedColumns.length > 0) {
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
  }
  
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()
    
    const selection = window.getSelection()
    let range
    
    if (selection.rangeCount > 0) {
      range = selection.getRangeAt(0)
    } else {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = tableHtml
    const tableElement = tempDiv.firstElementChild
    
    range.deleteContents()
    range.insertNode(tableElement)
    
    range.setStartAfter(tableElement)
    range.collapse(true)
    selection.removeAllRanges()
    selection.addRange(range)
    
    content.value = editorElement.innerHTML
  }
  
  ElMessage.success('数据表格已插入到光标位置')
}

// 数据源变化处理
const onDataSourceChange = async (dataSourceId) => {
  if (!dataSourceId) {
    datasetFields.value = []
    previewData.value = []
    return
  }
  
  try {
    const response = await api.getDataSourceFields(dataSourceId)
    datasetFields.value = response.fields || []
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
  
  const mockData = []
  for (let i = 0; i < 3; i++) {
    const row = {}
    datasetFields.value.forEach(field => {
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
    tableForm.columns = [
      { title: '列1', field: '' },
      { title: '列2', field: '' },
      { title: '列3', field: '' }
    ]
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
  let dataAttrs = `data-table-title="${tableForm.title}" data-source-id="${tableForm.dataSourceId || ''}" data-mode="${tableForm.dataMode}"`
  
  if (tableForm.dataMode === 'SQL') {
    dataAttrs += ` data-sql="${tableForm.sqlQuery || ''}"`
  } else if (tableForm.dataMode === 'DATASET') {
    dataAttrs += ` data-columns='${JSON.stringify(tableForm.columns)}'`
  }
  
  let headersHtml = ''
  tableForm.columns.forEach(col => {
    headersHtml += `<th style="border: 1px solid #ddd; padding: 8px; background-color: #f5f5f5;">${col.title || '未命名列'}</th>`
  })
  
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
  
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()
    
    const selection = window.getSelection()
    let range
    
    if (selection.rangeCount > 0) {
      range = selection.getRangeAt(0)
    } else {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }
    
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = tableHtml
    const tableElement = tempDiv.firstElementChild
    
    range.deleteContents()
    range.insertNode(tableElement)
    
    range.setStartAfter(tableElement)
    range.collapse(true)
    selection.removeAllRanges()
    selection.addRange(range)
    
    content.value = editorElement.innerHTML
  }
  
  showTableDialog.value = false
  ElMessage.success('动态表格已插入')
}

// 将结构转换为HTML
const convertStructureToHtml = (structure) => {
  let html = ''
  
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
      <div style="
        position: absolute;
        top: 60px;
        left: 60px;
        right: 60px;
        bottom: 60px;
        border: 2px solid #666;
        border-style: none;
      ">
        <div style="
          position: absolute;
          top: 0;
          left: 0;
          width: 60px;
          height: 60px;
          border-top: 2px solid #666;
          border-left: 2px solid #666;
        "></div>
        <div style="
          position: absolute;
          top: 0;
          right: 0;
          width: 60px;
          height: 60px;
          border-top: 2px solid #666;
          border-right: 2px solid #666;
        "></div>
        <div style="
          position: absolute;
          bottom: 0;
          left: 0;
          width: 60px;
          height: 60px;
          border-bottom: 2px solid #666;
          border-left: 2px solid #666;
        "></div>
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

onMounted(() => {
  console.log('Component mounted, template ID:', templateId.value)
  loadDataSources()
  initEditor()
  
  // 设置全局表格函数并增强表格编辑
  setupGlobalTableFunctions()
  setTimeout(() => {
    enhanceTableEditing()
    // 确保表格按钮已更新
    setupGlobalTableFunctions()
  }, 1000)
})
</script>

<style lang="scss">
/* Word 编辑器样式 */
.word-editor {
  height: calc(100vh - 200px);
  border: 1px solid #ddd;
  padding: 40px 60px;
  background: white;
  overflow-y: auto;
  line-height: 1.6;
  max-width: 21cm;
  margin: 0 auto;
  box-shadow: 0 0 10px rgba(0,0,0,0.05);
  font-family: "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimSun", "宋体", "Noto Sans SC", "Source Han Sans SC", "PingFang SC", "Hiragino Sans GB", "Heiti SC", "WenQuanYi Micro Hei", Arial, sans-serif !important;
  font-size: 12pt;
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
      padding: 2px 4px;
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
    
    td[colspan], th[colspan] {
      text-align: center;
      font-weight: bold;
    }
    
    td[rowspan], th[rowspan] {
      vertical-align: middle;
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
  
  /* 表格单元格编辑增强 */
  td:focus, th:focus {
    outline: 2px solid #007bff;
    outline-offset: -1px;
    background-color: rgba(0, 123, 255, 0.05);
  }
  
  /* 动态字段样式 */
  .dynamic-field {
    background-color: #e3f2fd;
    padding: 2px 6px;
    border-radius: 3px;
    border: 1px dashed #2196f3;
    display: inline-block;
    min-width: 60px;
    cursor: pointer;
  }
  
  .dynamic-table {
    border: 2px solid #4caf50;
    position: relative;
    margin: 10px 0;
  }
}
</style>

<style scoped lang="scss">
.template-editor-page {
  height: 100vh;
  display: flex;
  flex-direction: column;
  background: #f5f5f5;
  
  .editor-container {
    flex: 1;
    padding: 20px;
    overflow: auto;
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
}
</style>