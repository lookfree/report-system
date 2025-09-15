<template>
  <div class="template-editor-page">
    <!-- 顶部工具栏 -->
    <div class="editor-toolbar">
      <div class="toolbar-group">
        <el-button type="primary" @click="saveTemplate" :loading="saving">
          <el-icon><DocumentChecked /></el-icon>
          {{ saving ? '保存中...' : '保存模板' }}
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

      <!-- 自动保存状态 -->
      <div class="auto-save-status" v-if="autoSaveEnabled">
        <el-icon v-if="autoSaving" class="is-loading"><Loading /></el-icon>
        <span v-if="lastSaveTime">
          {{ autoSaving ? '自动保存中...' : `已自动保存 ${formatSaveTime(lastSaveTime)}` }}
        </span>
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
        <el-form-item label="展示内容">
          <el-radio-group v-model="fieldForm.insertType">
            <el-radio value="FIELD">仅文字</el-radio>
            <el-radio value="DATASET">数据集</el-radio>
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
        
        <!-- 数据集配置 -->
        <template v-if="fieldForm.insertType === 'DATASET'">
          <el-form-item label="数据集">
            <el-select
              v-model="fieldForm.datasetId"
              style="width: 100%"
              @change="onConfiguredDatasetChange"
              placeholder="请选择已配置的数据集"
            >
              <el-option
                v-for="dataset in configuredDatasets"
                :key="dataset.id"
                :label="dataset.name"
                :value="dataset.id"
              >
                <div>
                  <div>{{ dataset.name }}</div>
                  <div style="font-size: 12px; color: #999;">{{ dataset.description }}</div>
                </div>
              </el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="数据展示方式" v-if="selectedDataset">
            <el-radio-group v-model="fieldForm.displayMode">
              <el-radio value="SINGLE">单条（仅显示第一条数据）</el-radio>
              <el-radio value="LIST" v-if="selectedDataset.type === 'list'">列表（当前sheet页）</el-radio>
            </el-radio-group>
          </el-form-item>

          <!-- 字段选择 -->
          <el-form-item label="展示字段" v-if="selectedDataset && datasetPreview">
            <div style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; border-radius: 4px;">
              <!-- 单条模式时使用单选 -->
              <el-radio-group v-if="fieldForm.displayMode === 'SINGLE'" v-model="fieldForm.selectedField">
                <div v-for="field in selectedDataset.fields" :key="field" style="margin-bottom: 8px;">
                  <el-radio :value="field" style="width: 100%;">
                    <span style="font-weight: 500;">{{ field }}</span>
                    <span style="font-size: 12px; color: #999; margin-left: 10px;">（仅显示第一条数据的此字段）</span>
                  </el-radio>
                </div>
              </el-radio-group>
              <!-- 列表模式时使用多选 -->
              <el-checkbox-group v-else-if="fieldForm.displayMode === 'LIST'" v-model="fieldForm.displayFields">
                <div v-for="field in selectedDataset.fields" :key="field" style="margin-bottom: 8px;">
                  <el-checkbox :value="field" style="width: 100%;">
                    <span style="font-weight: 500;">{{ field }}</span>
                  </el-checkbox>
                </div>
              </el-checkbox-group>
            </div>
          </el-form-item>

          <!-- 字段值预览 -->
          <el-form-item label="预览数据" v-if="datasetPreview && ((fieldForm.displayMode === 'SINGLE' && fieldForm.selectedField) || (fieldForm.displayMode === 'LIST' && fieldForm.displayFields && fieldForm.displayFields.length > 0))">
            <div style="max-height: 200px; overflow-y: auto; background: #f8f9fa; border: 1px solid #ddd; border-radius: 4px;">
              <table style="width: 100%; border-collapse: collapse;">
                <thead>
                  <tr style="background: #e9ecef;">
                    <th v-for="fieldName in (fieldForm.displayMode === 'SINGLE' ? [fieldForm.selectedField] : fieldForm.displayFields)" :key="fieldName"
                        style="padding: 8px; border: 1px solid #ddd; font-weight: 500; text-align: left;">
                      {{ fieldName }}
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, index) in (datasetPreview.data ? (fieldForm.displayMode === 'SINGLE' ? [Array.isArray(datasetPreview.data) ? datasetPreview.data[0] : datasetPreview.data] : datasetPreview.data.slice(0, 4)) : [])"
                      :key="index"
                      style="background: white;">
                    <td v-for="fieldName in (fieldForm.displayMode === 'SINGLE' ? [fieldForm.selectedField] : fieldForm.displayFields)" :key="fieldName"
                        style="padding: 8px; border: 1px solid #ddd;">
                      {{ row[fieldName] }}
                    </td>
                  </tr>
                </tbody>
              </table>
              <div v-if="selectedDataset.type === 'list' && datasetPreview.data && datasetPreview.data.length > 4"
                   style="padding: 8px; color: #666; text-align: center; font-size: 12px;">
                共 {{ datasetPreview.data.length }} 条数据，仅显示前 4 条预览
              </div>
            </div>
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
        <el-button type="danger" @click="clearCurrentCellDataset" v-if="currentEditingCell && fieldForm.insertType === 'DATASET'">清除数据集</el-button>
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
import { ref, reactive, onMounted, onUnmounted, computed, watch, nextTick } from 'vue'
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
  Back,
  Loading,
  Coin
} from '@element-plus/icons-vue'
import { useRoute, useRouter } from 'vue-router'
import api from '@/api'

const route = useRoute()
const router = useRouter()

// 编辑器内容
const content = ref('')
const templateId = ref(route.params.id || null)
const templateName = ref('')

// 保存状态
const saving = ref(false)
const autoSaving = ref(false)
const autoSaveEnabled = ref(true)
const lastSaveTime = ref(null)
const hasUnsavedChanges = ref(false)
let autoSaveTimer = null

// 编辑器工具栏状态
const currentFontSize = ref('12pt')
let savedSelection = null // 保存的选区

// 表格合并相关状态
const mergedCells = ref({}) // 存储合并单元格信息
const selectedCells = ref({}) // 选中的单元格
const dragStart = ref(null) // 拖拽开始位置
const dragEnd = ref(null) // 拖拽结束位置
const mergeMode = ref(false) // 是否处于合并模式

// 单元格数据集相关状态
const currentEditingCell = ref(null) // 当前正在编辑的单元格

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
  insertType: 'FIELD', // FIELD 或 DATASET
  name: '',
  tableTitle: '',
  dataType: 'FIXED',
  dataSourceId: null,
  sqlQuery: '',
  defaultValue: '',
  dateFormat: 'YYYY-MM-DD',
  systemVariable: 'CURRENT_USER',
  datasetField: '',
  selectedColumns: [],
  // 数据集相关
  datasetId: null, // 配置的数据集ID
  datasetName: '', // 兼容旧的mock数据集
  dataStructure: 'SINGLE', // SINGLE 或 LIST
  displayMode: 'SINGLE', // SINGLE（单条） 或 LIST（列表）
  sheetConfig: 'CURRENT', // CURRENT 或 SEPARATE
  displayFields: [],
  selectedField: '' // 单条数据时选择的单个字段
})

// 数据集预览结果
const datasetPreview = ref(null)
const previewLoading = ref(false)

// 配置的数据集列表（从API获取）
const configuredDatasets = ref([])

// 选中的数据集
const selectedDataset = computed(() => {
  return configuredDatasets.value.find(d => d.id === fieldForm.datasetId)
})

// Mock数据集
const mockDatasets = ref([
  {
    name: 'asset_source_distribution',
    displayName: '资产来源分布',
    fields: [
      { name: 'total', displayName: '总数', type: 'number' },
      { name: 'passSource', displayName: '来源', type: 'string' },
      { name: 'count', displayName: '数量', type: 'number' }
    ],
    data: [
      { total: 828, passSource: '网关', count: 615 },
      { total: 1343, passSource: '旁路', count: 950 },
      { total: 8, passSource: '1', count: 6 },
      { total: 10, passSource: '2', count: 8 }
    ]
  },
  {
    name: 'security_audit_results',
    displayName: '安全审计结果',
    fields: [
      { name: 'auditUnit', displayName: '审计单位', type: 'string' },
      { name: 'riskLevel', displayName: '风险等级', type: 'string' },
      { name: 'issueCount', displayName: '问题数量', type: 'number' },
      { name: 'score', displayName: '评分', type: 'number' }
    ],
    data: [
      { auditUnit: '信息技术部', riskLevel: '中等', issueCount: 15, score: 85 },
      { auditUnit: '财务部', riskLevel: '低', issueCount: 3, score: 95 },
      { auditUnit: '人事部', riskLevel: '高', issueCount: 28, score: 68 },
      { auditUnit: '运营部', riskLevel: '中等', issueCount: 12, score: 88 }
    ]
  },
  {
    name: 'interface_vulnerability',
    displayName: '接口-定级备案单元数据分布钻取',
    fields: [
      { name: 'method', displayName: '方法', type: 'string' },
      { name: 'vulnerabilityCount', displayName: '漏洞数量', type: 'number' },
      { name: 'severityLevel', displayName: '严重程度', type: 'string' }
    ],
    data: [
      { method: 'post', vulnerabilityCount: 25, severityLevel: '高' },
      { method: 'get', vulnerabilityCount: 18, severityLevel: '中' },
      { method: 'put', vulnerabilityCount: 7, severityLevel: '低' },
      { method: 'delete', vulnerabilityCount: 12, severityLevel: '中' }
    ]
  }
])

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

// 计算属性：当前选中数据集的字段
const currentDatasetFields = computed(() => {
  const dataset = mockDatasets.value.find(d => d.name === fieldForm.datasetName)
  return dataset ? dataset.fields : []
})

// 计算属性：当前选中数据集的数据
const currentDatasetData = computed(() => {
  const dataset = mockDatasets.value.find(d => d.name === fieldForm.datasetName)
  return dataset ? dataset.data : []
})

// 获取字段显示名称
const getFieldDisplayName = (fieldName) => {
  const field = currentDatasetFields.value.find(f => f.name === fieldName)
  return field ? field.displayName : fieldName
}

// 监听字段对话框打开
watch(() => showFieldDialog.value, async (newVal) => {
  if (newVal) {
    // 每次对话框打开时加载配置的数据集
    await loadConfiguredDatasets()
  }
})

// 监听数据结构变化
watch(() => fieldForm.dataStructure, (newVal) => {
  if (fieldForm.datasetName) {
    const dataset = mockDatasets.value.find(d => d.name === fieldForm.datasetName)
    if (dataset && dataset.fields.length > 0) {
      if (newVal === 'SINGLE') {
        // 切换到单条时，清空多选字段，设置单选字段
        fieldForm.displayFields = []
        fieldForm.selectedField = dataset.fields[0].name
      } else {
        // 切换到列表时，清空单选字段，设置多选字段
        fieldForm.selectedField = ''
        fieldForm.displayFields = dataset.fields.slice(0, Math.min(3, dataset.fields.length)).map(f => f.name)
      }
    }
  }
})

// 计算属性：格式化字段显示
const getFieldDisplay = (field) => {
  return field ? `{{${field}}}` : '示例数据'
}

// 处理内容变化
const handleContentChange = (event) => {
  content.value = event.target.innerHTML
  hasUnsavedChanges.value = true

  // 触发自动保存
  if (autoSaveEnabled.value) {
    clearTimeout(autoSaveTimer)
    autoSaveTimer = setTimeout(() => {
      autoSave()
    }, 3000) // 3秒后自动保存
  }
}

// 格式化保存时间
const formatSaveTime = (time) => {
  if (!time) return ''
  const now = new Date()
  const saveTime = new Date(time)
  const diff = Math.floor((now - saveTime) / 1000)

  if (diff < 60) return '刚刚'
  if (diff < 3600) return `${Math.floor(diff / 60)}分钟前`
  return saveTime.toLocaleTimeString('zh-CN')
}

// 自动保存
const autoSave = async () => {
  if (!hasUnsavedChanges.value || saving.value) return

  try {
    autoSaving.value = true

    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value

    await api.saveTemplateHtml(templateId.value, htmlContent)

    hasUnsavedChanges.value = false
    lastSaveTime.value = new Date()
    console.log('自动保存成功')
  } catch (error) {
    console.error('自动保存失败:', error)
  } finally {
    autoSaving.value = false
  }
}

// 初始化编辑器
const initEditor = () => {
  console.log('Initializing simple HTML editor')

  // 监听编辑器的选区变化，自动保存有效选区
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    // 监听鼠标松开事件（选择文本完成）
    editorElement.addEventListener('mouseup', () => {
      setTimeout(() => {
        const selection = window.getSelection()
        if (selection.rangeCount > 0 && !selection.getRangeAt(0).collapsed) {
          // 只有当有实际选中内容时才保存
          savedSelection = selection.getRangeAt(0).cloneRange()
          console.log('已保存选区')
        }
      }, 10)
    })

    // 监听键盘事件（通过键盘选择文本）
    editorElement.addEventListener('keyup', (e) => {
      // Shift + 方向键选择文本
      if (e.shiftKey) {
        const selection = window.getSelection()
        if (selection.rangeCount > 0 && !selection.getRangeAt(0).collapsed) {
          savedSelection = selection.getRangeAt(0).cloneRange()
          console.log('已保存选区（键盘）')
        }
      }
    })
  }

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

    // 加载内容后处理表格
    setTimeout(() => {
      // 移除所有表格工具栏（包含行列增减按钮）
      const toolbars = editorElement.querySelectorAll('.table-toolbar')
      toolbars.forEach(toolbar => {
        toolbar.remove()
      })

      // 增强表格编辑功能
      enhanceTableEditing()
    }, 100)
  } catch (error) {
    console.error('Failed to load template:', error)
    ElMessage.error('加载模板失败: ' + (error.response?.data?.error || error.message))
  }
}

// 保存模板
const saveTemplate = async () => {
  if (saving.value) return

  try {
    saving.value = true

    if (!templateId.value) {
      ElMessage.error('模板ID不存在，无法保存')
      return
    }

    // 从 HTML 编辑器获取内容
    const editorElement = document.getElementById('word-editor')
    let htmlContent = editorElement ? editorElement.innerHTML : content.value

    await api.saveTemplateHtml(templateId.value, htmlContent)

    hasUnsavedChanges.value = false
    lastSaveTime.value = new Date()
    ElMessage.success('模板保存成功')
  } catch (error) {
    ElMessage.error('保存失败: ' + (error.response?.data?.error || error.message))
  } finally {
    saving.value = false
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

  // 加载配置的数据集
  loadConfiguredDatasets()

  showFieldDialog.value = true
  fieldForm.insertType = 'FIELD'
  fieldForm.name = ''
  fieldForm.tableTitle = ''
  fieldForm.dataType = 'FIXED'
  fieldForm.dataSourceId = null
  fieldForm.sqlQuery = ''
  fieldForm.defaultValue = ''
  fieldForm.dateFormat = 'YYYY-MM-DD'
  // 清空数据集相关字段
  fieldForm.datasetId = null
  fieldForm.datasetName = ''
  fieldForm.displayFields = []
  fieldForm.selectedField = ''
  datasetPreview.value = null
  fieldForm.systemVariable = 'CURRENT_USER'
  fieldForm.datasetField = ''
  fieldForm.selectedColumns = []
  // 数据集相关字段重置
  fieldForm.datasetName = ''
  fieldForm.dataStructure = 'SINGLE'
  fieldForm.sheetConfig = 'CURRENT'
  fieldForm.displayFields = []
  fieldForm.selectedField = ''
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

// 加载配置的数据集列表
const loadConfiguredDatasets = async () => {
  try {
    const response = await fetch('/api/datasets')
    const data = await response.json()
    if (data.success) {
      configuredDatasets.value = data.datasets || []
    }
  } catch (error) {
    console.error('Failed to load datasets:', error)
  }
}

// 配置的数据集选择变化处理
const onConfiguredDatasetChange = async () => {
  // 清空之前选择的字段
  fieldForm.displayFields = []
  fieldForm.selectedField = ''
  datasetPreview.value = null

  if (!fieldForm.datasetId) return

  // 获取数据集预览
  previewLoading.value = true
  try {
    const response = await fetch(`/api/datasets/${fieldForm.datasetId}/execute`)
    const data = await response.json()

    if (data.success && data.result) {
      datasetPreview.value = data.result

      // 根据数据类型设置默认展示模式和选中的字段
      if (selectedDataset.value) {
        // 设置默认展示模式
        if (selectedDataset.value.type === 'single') {
          // 单条数据集默认为单条模式
          fieldForm.displayMode = 'SINGLE'
          fieldForm.selectedField = selectedDataset.value.fields[0] || ''
        } else {
          // 列表数据集默认为单条模式（用于统计场景）
          fieldForm.displayMode = 'SINGLE'
          fieldForm.selectedField = selectedDataset.value.fields[0] || ''
          // 如果需要列表模式，可以手动切换
          // fieldForm.displayFields = selectedDataset.value.fields.slice(0, Math.min(3, selectedDataset.value.fields.length))
        }
      }
    }
  } catch (error) {
    console.error('Failed to preview dataset:', error)
    ElMessage.error('获取数据集预览失败')
  } finally {
    previewLoading.value = false
  }
}

// 旧的Mock数据集变化处理（保留兼容）
const onDatasetChange = () => {
  // 清空之前选择的字段
  fieldForm.displayFields = []
  fieldForm.selectedField = ''

  if (fieldForm.datasetName) {
    const dataset = mockDatasets.value.find(d => d.name === fieldForm.datasetName)
    if (dataset && dataset.fields.length > 0) {
      if (fieldForm.dataStructure === 'SINGLE') {
        // 单条数据时默认选中第一个字段
        fieldForm.selectedField = dataset.fields[0].name
      } else {
        // 列表数据时默认选中前几个字段
        fieldForm.displayFields = dataset.fields.slice(0, Math.min(3, dataset.fields.length)).map(f => f.name)
      }
    }
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

// 确认插入数据
const confirmInsertField = () => {
  if (fieldForm.insertType === 'FIELD') {
    insertFieldElement()
  } else if (fieldForm.insertType === 'DATASET') {
    insertDatasetElement()
  }
}

// 插入数据元素
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
    
    // 插入数据元素
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

// 插入数据集元素
const insertDatasetElement = () => {
  // 检查是否是在表格单元格中插入
  const isTableCell = currentEditingCell.value !== null

  // 优先使用配置的数据集
  if (fieldForm.datasetId) {
    if (!selectedDataset.value) {
      ElMessage.warning('请选择数据集')
      return
    }

    // 根据展示模式验证字段选择
    if (fieldForm.displayMode === 'SINGLE' && !fieldForm.selectedField) {
      ElMessage.warning('请选择要展示的字段')
      return
    }

    if (fieldForm.displayMode === 'LIST' && fieldForm.displayFields.length === 0) {
      ElMessage.warning('请选择要展示的字段')
      return
    }

    // 如果是表格单元格插入
    if (isTableCell) {
      const cell = currentEditingCell.value

      if (fieldForm.displayMode === 'SINGLE') {
        // 单条模式 - 在光标位置插入数据集字段占位符（支持多个字段混合文本）
        const placeholder = `<span class="dataset-placeholder-inline" data-dataset-id="${selectedDataset.value.id}" data-dataset-name="${selectedDataset.value.name}" data-field-name="${fieldForm.selectedField}" data-data-type="single" data-display-mode="SINGLE" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 2px 6px; border-radius: 3px; font-weight: 500; font-size: 0.9em; display: inline-block; margin: 0 2px; cursor: pointer; user-select: none;" title="双击删除或按Delete键删除">📊${fieldForm.selectedField}</span>`

        // 获取当前光标位置并插入占位符
        const selection = window.getSelection()
        if (selection.rangeCount > 0) {
          const range = selection.getRangeAt(0)

          // 如果是在单元格内，确保range在单元格内
          if (cell.contains(range.commonAncestorContainer)) {
            // 在光标位置插入占位符
            const placeholderElement = document.createElement('span')
            placeholderElement.innerHTML = placeholder
            range.insertNode(placeholderElement.firstChild)

            // 移动光标到占位符后面
            range.setStartAfter(placeholderElement.firstChild)
            range.collapse(true)
            selection.removeAllRanges()
            selection.addRange(range)
          } else {
            // 如果光标不在单元格内，追加到单元格末尾
            cell.innerHTML += placeholder
          }
        } else {
          // 没有选区时，追加到单元格末尾
          cell.innerHTML += placeholder
        }
      } else if (fieldForm.displayMode === 'LIST') {
        // 列表模式 - 扩展到表格的多个单元格
        const table = cell.closest('table')
        const row = cell.parentElement
        const startCellIndex = cell.cellIndex
        const startRowIndex = row.rowIndex

        // 标记起始单元格为数据集占位符
        cell.innerHTML = `<div class="dataset-placeholder-start" data-dataset-id="${selectedDataset.value.id}" data-dataset-name="${selectedDataset.value.name}" data-display-fields="${fieldForm.displayFields.join(',')}" data-data-type="list-start" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 3px 6px; border-radius: 4px; font-weight: 500; font-size: 12px;">📊 ${selectedDataset.value.name}</div>`

        // 为接下来的单元格添加占位符标记
        const fieldsCount = fieldForm.displayFields.length

        // 填充当前行剩余的单元格
        for (let i = 1; i < fieldsCount && (startCellIndex + i) < row.cells.length; i++) {
          const nextCell = row.cells[startCellIndex + i]
          nextCell.innerHTML = `<div class="dataset-placeholder-field" data-field-index="${i}" style="background: linear-gradient(135deg, #e0e7ff 0%, #f0f4ff 100%); color: #667eea; padding: 3px 6px; border-radius: 2px; font-size: 11px; text-align: center;">${fieldForm.displayFields[i]}</div>`
        }

        // 如果需要，在下面的行继续填充
        if (datasetPreview.value && datasetPreview.value.data) {
          const previewData = Array.isArray(datasetPreview.value.data) ?
            datasetPreview.value.data.slice(0, 3) : [datasetPreview.value.data]

          for (let rowIdx = 0; rowIdx < previewData.length && (startRowIndex + rowIdx + 1) < table.rows.length; rowIdx++) {
            const nextRow = table.rows[startRowIndex + rowIdx + 1]
            for (let colIdx = 0; colIdx < fieldsCount && (startCellIndex + colIdx) < nextRow.cells.length; colIdx++) {
              const dataCell = nextRow.cells[startCellIndex + colIdx]
              dataCell.innerHTML = `<div class="dataset-placeholder-data" data-row="${rowIdx}" data-col="${colIdx}" style="color: #999; font-style: italic; font-size: 12px;">{{data}}</div>`
            }
          }
        }
      }

      // 更新编辑器内容
      const editorElement = document.getElementById('word-editor')
      content.value = editorElement.innerHTML

      // 清空当前编辑单元格引用
      currentEditingCell.value = null

      showFieldDialog.value = false
      ElMessage.success('数据集已插入')
      return
    }

    // 生成带数据集名称的表格HTML（非单元格插入）
    let datasetHtml = ''

    if (selectedDataset.value.type === 'single') {
      // 单条数据 - 显示字段名和占位符
      datasetHtml = `<span class="dynamic-field dataset-placeholder" data-dataset-id="${selectedDataset.value.id}" data-dataset-name="${selectedDataset.value.name}" data-field-name="${fieldForm.selectedField}" data-data-type="single" contenteditable="false" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 3px 10px; border-radius: 4px; font-weight: 500; display: inline-block; margin: 0 2px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">📊 ${selectedDataset.value.name}:${fieldForm.selectedField}</span>`
    } else {
      // 列表数据 - 生成表格显示
      const headerRow = fieldForm.displayFields.map(field =>
        `<th style="border: 1px solid #ddd; padding: 8px; background: #f5f5f5; font-weight: bold;">${field}</th>`
      ).join('')

      // 如果有预览数据，显示几行示例
      let dataRows = ''
      if (datasetPreview.value && datasetPreview.value.data) {
        const previewData = Array.isArray(datasetPreview.value.data) ?
          datasetPreview.value.data.slice(0, 3) :
          [datasetPreview.value.data]

        dataRows = previewData.map(row => {
          const cells = fieldForm.displayFields.map(field =>
            `<td style="border: 1px solid #ddd; padding: 8px;">${row[field] || '-'}</td>`
          ).join('')
          return `<tr>${cells}</tr>`
        }).join('')
      } else {
        // 没有预览数据时显示占位符
        const cells = fieldForm.displayFields.map(() =>
          `<td style="border: 1px solid #ddd; padding: 8px; color: #999;">...</td>`
        ).join('')
        dataRows = `<tr>${cells}</tr>`
      }

      datasetHtml = `
        <div class="dynamic-table"
             data-dataset-id="${selectedDataset.value.id}"
             data-dataset-name="${selectedDataset.value.name}"
             data-display-fields="${fieldForm.displayFields.join(',')}"
             data-data-type="list"
             contenteditable="false"
             style="margin: 10px 0;">
          <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 10px; border-radius: 6px 6px 0 0; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            📊 数据集: ${selectedDataset.value.name}
            <span style="float: right; font-size: 12px; opacity: 0.9; font-weight: normal;">
              占位符: {{dataset:${selectedDataset.value.id}}}
            </span>
          </div>
          <table style="width: 100%; border-collapse: collapse; border: 2px solid #667eea;">
            <thead>
              <tr style="background: #f8f9ff;">${headerRow}</tr>
            </thead>
            <tbody>
              ${dataRows}
            </tbody>
          </table>
          <div style="text-align: center; color: #666; font-size: 12px; padding: 8px; background: linear-gradient(to bottom, #f8f9ff, #ffffff); border: 2px solid #667eea; border-top: none; border-radius: 0 0 6px 6px; font-style: italic;">
            🔄 ${selectedDataset.value.description || '数据将在导出时动态获取'}
          </div>
        </div>
      `
    }

    // 插入到编辑器
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      editorElement.focus()

      // 获取当前选区或恢复之前保存的位置
      let range = restoreSelection()

      if (!range) {
        const selection = window.getSelection()
        if (selection.rangeCount > 0) {
          range = selection.getRangeAt(0)
        } else {
          // 如果没有选区，在编辑器末尾插入
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

      // 创建临时容器解析HTML
      const tempDiv = document.createElement('div')
      tempDiv.innerHTML = datasetHtml
      const element = tempDiv.firstElementChild

      // 插入元素
      range.deleteContents()
      range.insertNode(element)

      // 在元素后插入空格以便继续编辑
      const space = document.createTextNode(' ')
      range.setStartAfter(element)
      range.insertNode(space)
      range.setStartAfter(space)
      range.collapse(true)

      const selection = window.getSelection()
      selection.removeAllRanges()
      selection.addRange(range)

      // 更新内容
      content.value = editorElement.innerHTML
    }

    showFieldDialog.value = false
    currentEditingCell.value = null // 清空单元格引用
    ElMessage.success('数据集已插入')
    return
  }

  // 兼容旧的mock数据集
  if (!fieldForm.datasetName) {
    ElMessage.warning('请选择数据集')
    return
  }

  if (fieldForm.dataStructure === 'SINGLE' && !fieldForm.selectedField) {
    ElMessage.warning('请选择要展示的字段')
    return
  }

  if (fieldForm.dataStructure === 'LIST' && fieldForm.displayFields.length === 0) {
    ElMessage.warning('请选择要展示的字段')
    return
  }

  const dataset = mockDatasets.value.find(d => d.name === fieldForm.datasetName)
  if (!dataset) {
    ElMessage.error('数据集不存在')
    return
  }

  let datasetHtml = ''

  if (fieldForm.dataStructure === 'SINGLE') {
    // 单条数据 - 生成单个字段变量
    const field = dataset.fields.find(f => f.name === fieldForm.selectedField)
    const displayName = field ? field.displayName : fieldForm.selectedField
    datasetHtml = `<span class="dynamic-field dataset-placeholder" data-dataset-name="${fieldForm.datasetName}" data-field-name="${fieldForm.selectedField}" data-data-structure="SINGLE" contenteditable="false" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 3px 10px; border-radius: 4px; font-weight: 500; display: inline-block; margin: 0 2px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">📊 {{${displayName}}}</span>`
  } else {
    // 列表数据 - 生成表格
    const headerRow = fieldForm.displayFields.map(fieldName => {
      const field = dataset.fields.find(f => f.name === fieldName)
      return field ? field.displayName : fieldName
    }).join('</th><th style="border: 1px solid #ddd; padding: 8px; background: #f5f5f5; font-weight: bold;">')

    const dataRows = dataset.data.slice(0, 3).map(row => {
      const cells = fieldForm.displayFields.map(fieldName => row[fieldName] || '').join('</td><td style="border: 1px solid #ddd; padding: 8px;">')
      return `<tr><td style="border: 1px solid #ddd; padding: 8px;">${cells}</td></tr>`
    }).join('')

    datasetHtml = `
      <div class="dynamic-table"
           data-dataset-name="${fieldForm.datasetName}"
           data-display-fields="${fieldForm.displayFields.join(',')}"
           data-data-structure="LIST"
           data-sheet-config="${fieldForm.sheetConfig}"
           contenteditable="false"
           style="margin: 10px 0;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 10px; border-radius: 6px 6px 0 0; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
          📊 数据集: ${dataset.displayName}
          <span style="float: right; font-size: 12px; opacity: 0.9; font-weight: normal;">
            共 ${dataset.data.length} 条数据
          </span>
        </div>
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #667eea;">
          <thead>
            <tr style="background: #f8f9ff;">
              <th style="border: 1px solid #ddd; padding: 8px; background: #f8f9ff; font-weight: bold;">${headerRow}</th>
            </tr>
          </thead>
          <tbody>
            ${dataRows}
          </tbody>
        </table>
        <div style="text-align: center; color: #666; font-size: 12px; padding: 8px; background: linear-gradient(to bottom, #f8f9ff, #ffffff); border: 2px solid #667eea; border-top: none; border-radius: 0 0 6px 6px; font-style: italic;">
          🔄 数据将在导出时动态获取 | 仅显示预览
        </div>
      </div>
    `
  }

  // 插入到编辑器当前光标位置
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    editorElement.focus()

    // 尝试恢复保存的光标位置
    let range = restoreSelection()

    if (!range) {
      const selection = window.getSelection()
      if (selection.rangeCount > 0) {
        range = selection.getRangeAt(0)
      } else {
        range = document.createRange()
        range.selectNodeContents(editorElement)
        range.collapse(false)
      }
    }

    if (!editorElement.contains(range.commonAncestorContainer)) {
      range = document.createRange()
      range.selectNodeContents(editorElement)
      range.collapse(false)
    }

    // 创建临时容器解析HTML
    const tempDiv = document.createElement('div')
    tempDiv.innerHTML = datasetHtml
    const elements = Array.from(tempDiv.childNodes)

    // 插入所有元素
    range.deleteContents()
    elements.forEach(element => {
      range.insertNode(element)
      range.setStartAfter(element)
      range.setEndAfter(element)
    })

    // 更新选择位置到最后插入的元素后面
    const selection = window.getSelection()
    selection.removeAllRanges()
    selection.addRange(range)

    // 更新内容
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true
  }

  showFieldDialog.value = false
  ElMessage.success(`数据集已插入到光标位置 (${fieldForm.dataStructure === 'SINGLE' ? '单条' : '列表'}模式)`)
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
  // 保存并恢复选区
  saveTextSelection()
  restoreTextSelection()

  document.execCommand(command, false, null)
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    content.value = editorElement.innerHTML
  }
}

// 文本对齐函数
const alignText = (alignment) => {
  // 保存并恢复选区
  saveTextSelection()
  restoreTextSelection()

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
  // 保存并恢复选区
  saveTextSelection()
  restoreTextSelection()

  const command = direction === 'increase' ? 'indent' : 'outdent'
  document.execCommand(command, false, null)
  const editorElement = document.getElementById('word-editor')
  if (editorElement) {
    content.value = editorElement.innerHTML
  }
}

// 保存文本选区
const saveTextSelection = () => {
  const selection = window.getSelection()
  if (selection.rangeCount > 0) {
    savedSelection = selection.getRangeAt(0).cloneRange()
  }
}

// 恢复文本选区
const restoreTextSelection = () => {
  if (savedSelection) {
    const selection = window.getSelection()
    selection.removeAllRanges()
    selection.addRange(savedSelection)
  }
}

// 更改字体大小函数
const changeFontSize = (size) => {
  // 检查是否有保存的选区
  if (!savedSelection || savedSelection.collapsed) {
    ElMessage.warning('请先选中要修改的文字')
    return
  }

  // 恢复保存的选区
  const selection = window.getSelection()
  selection.removeAllRanges()
  selection.addRange(savedSelection)

  const range = selection.getRangeAt(0)
  
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
  // 先保存选区，防止点击按钮时失去选区
  saveTextSelection()
  // 立即恢复选区
  restoreTextSelection()

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

// 处理单元格双击事件
const handleCellDoubleClick = (e) => {
  e.preventDefault()
  e.stopPropagation()

  const cell = e.currentTarget
  currentEditingCell.value = cell

  // 检查单元格是否已经有数据集（包括单条数据和列表数据的起始标记）
  const existingDataset = cell.querySelector('.dataset-placeholder, .dataset-placeholder-start')

  // 保存当前光标位置
  const selection = window.getSelection()
  const range = document.createRange()
  range.selectNodeContents(cell)
  selection.removeAllRanges()
  selection.addRange(range)
  savedSelection = range

  // 重置表单
  fieldForm.insertType = 'DATASET'
  fieldForm.datasetId = null
  fieldForm.selectedField = ''
  fieldForm.displayFields = []

  // 如果单元格已有数据集，回显选择
  if (existingDataset) {
    const datasetId = existingDataset.getAttribute('data-dataset-id')
    const fieldName = existingDataset.getAttribute('data-field-name')
    const dataType = existingDataset.getAttribute('data-data-type')
    const displayMode = existingDataset.getAttribute('data-display-mode')
    const displayFields = existingDataset.getAttribute('data-display-fields')

    // 先加载数据集，然后回显选择
    loadConfiguredDatasets().then(() => {
      // 回显数据集选择
      fieldForm.datasetId = datasetId

      // 等待Vue更新后设置字段
      nextTick(() => {
        // 设置展示模式
        if (displayMode) {
          fieldForm.displayMode = displayMode
        } else if (dataType === 'single') {
          fieldForm.displayMode = 'SINGLE'
        } else if (dataType === 'list' || dataType === 'list-start') {
          fieldForm.displayMode = 'LIST'
        }

        // 设置字段选择
        if (fieldForm.displayMode === 'SINGLE' && fieldName) {
          fieldForm.selectedField = fieldName
          fieldForm.displayFields = []
        } else if (fieldForm.displayMode === 'LIST' && displayFields) {
          fieldForm.displayFields = displayFields.split(',')
          fieldForm.selectedField = ''
        }

        // 触发数据集选择变化以加载预览
        if (datasetId) {
          onConfiguredDatasetChange()
        }
      })
    })
  } else {
    // 清空选择
    fieldForm.datasetId = null
    fieldForm.selectedField = ''
    fieldForm.displayFields = []
    datasetPreview.value = null
  }

  // 加载配置的数据集
  loadConfiguredDatasets()

  // 打开插入数据对话框
  showFieldDialog.value = true
}

// 清除当前单元格的数据集
const clearCurrentCellDataset = () => {
  if (!currentEditingCell.value) return

  // 清除单元格中的所有数据集占位符
  const placeholders = currentEditingCell.value.querySelectorAll('.dataset-placeholder, .dataset-placeholder-inline, .dataset-placeholder-start')
  placeholders.forEach(p => p.remove())

  // 如果单元格为空，添加默认文本
  if (!currentEditingCell.value.textContent.trim()) {
    currentEditingCell.value.textContent = ' '
  }

  // 更新内容
  const editorElement = document.getElementById('word-editor')
  content.value = editorElement.innerHTML
  hasUnsavedChanges.value = true

  showFieldDialog.value = false
  ElMessage.success('已清除单元格数据集')
}

// 处理内联数据集占位符的交互
const handleInlineDatasetInteraction = () => {
  const editorElement = document.getElementById('word-editor')
  if (!editorElement) return

  // 双击删除功能
  editorElement.addEventListener('dblclick', (e) => {
    const target = e.target
    if (target.classList.contains('dataset-placeholder-inline')) {
      e.preventDefault()
      e.stopPropagation()

      ElMessageBox.confirm('确定要删除这个数据集字段吗？', '提示', {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }).then(() => {
        target.remove()
        content.value = editorElement.innerHTML
        hasUnsavedChanges.value = true
        ElMessage.success('已删除数据集字段')
      }).catch(() => {})
    }
  })

  // 键盘删除功能
  editorElement.addEventListener('keydown', (e) => {
    if ((e.key === 'Delete' || e.key === 'Backspace')) {
      const selection = window.getSelection()
      if (!selection.rangeCount) return

      const range = selection.getRangeAt(0)
      let node = range.startContainer

      // 查找最近的内联占位符
      while (node && node !== editorElement) {
        if (node.nodeType === 1 && node.classList && node.classList.contains('dataset-placeholder-inline')) {
          e.preventDefault()
          node.remove()
          content.value = editorElement.innerHTML
          hasUnsavedChanges.value = true
          ElMessage.success('已删除数据集字段')
          return
        }
        node = node.parentNode
      }
    }
  })
}

// 增强表格编辑功能
const enhanceTableEditing = () => {
  const editorElement = document.getElementById('word-editor')
  if (!editorElement) return

  // 为所有表格添加编辑增强
  const tables = editorElement.querySelectorAll('table')
  tables.forEach(table => {
    // 设置表格样式使其可调整大小
    table.style.tableLayout = 'fixed'
    table.style.width = table.style.width || '100%'

    // 为表格单元格添加编辑监听
    const cells = table.querySelectorAll('td, th')
    cells.forEach(cell => {
      // 使单元格可编辑
      if (!cell.hasAttribute('contenteditable')) {
        cell.setAttribute('contenteditable', 'true')
      }

      // 设置单元格样式
      cell.style.position = 'relative'
      cell.style.overflow = 'hidden'

      // 如果单元格没有宽度，设置默认宽度
      if (!cell.style.width) {
        cell.style.width = '150px'
        cell.style.minWidth = '100px'
      }

      // 如果单元格没有高度，设置默认高度
      if (!cell.style.height) {
        cell.style.height = '40px'
        cell.style.minHeight = '30px'
      }

      // 移除旧的事件监听器（如果存在）
      cell.removeEventListener('input', handleCellEdit)
      cell.removeEventListener('blur', handleCellBlur)
      cell.removeEventListener('mousedown', handleCellMouseDown)
      cell.removeEventListener('mouseenter', handleCellMouseEnter)
      cell.removeEventListener('mouseup', handleCellMouseUp)
      cell.removeEventListener('dblclick', handleCellDoubleClick)

      // 添加新的事件监听器
      cell.addEventListener('input', handleCellEdit)
      cell.addEventListener('blur', handleCellBlur)

      // 添加双击事件处理数据集插入
      cell.addEventListener('dblclick', handleCellDoubleClick)

      // 添加拖拽选择事件
      cell.addEventListener('mousedown', handleCellMouseDown)
      cell.addEventListener('mouseenter', handleCellMouseEnter)
      cell.addEventListener('mouseup', handleCellMouseUp)

      // 添加右键菜单支持
      cell.addEventListener('contextmenu', showCellContextMenu)

      // 添加列调整大小功能
      addColumnResizeHandle(cell)

      // 添加列边框拖拽功能
      addColumnBorderResize(cell)

      // 只为每行的第一个单元格添加行调整大小功能
      if (cell.cellIndex === 0) {
        addRowResizeHandle(cell)
      }
    })
  })
}

// 添加列调整大小手柄
const addColumnResizeHandle = (cell) => {
  // 检查是否已经有调整手柄
  if (cell.querySelector('.resize-handle')) return

  // 创建调整手柄（放在右边框）
  const resizeHandle = document.createElement('div')
  resizeHandle.className = 'resize-handle'
  resizeHandle.style.cssText = `
    position: absolute;
    right: -3px;
    top: 0;
    width: 6px;
    height: 100%;
    cursor: col-resize;
    background: transparent;
    z-index: 100;
  `

  // 鼠标悬停时显示视觉效果
  resizeHandle.addEventListener('mouseenter', () => {
    resizeHandle.style.background = 'rgba(33, 150, 243, 0.5)'
    // 高亮整列边框
    const table = cell.closest('table')
    const colIndex = cell.cellIndex
    const rows = table.querySelectorAll('tr')
    rows.forEach(row => {
      if (row.cells[colIndex]) {
        row.cells[colIndex].style.borderRight = '2px solid rgba(33, 150, 243, 0.5)'
      }
    })
  })

  resizeHandle.addEventListener('mouseleave', () => {
    resizeHandle.style.background = 'transparent'
    // 恢复边框
    const table = cell.closest('table')
    const colIndex = cell.cellIndex
    const rows = table.querySelectorAll('tr')
    rows.forEach(row => {
      if (row.cells[colIndex]) {
        row.cells[colIndex].style.borderRight = ''
      }
    })
  })

  // 处理拖拽调整大小
  let isResizing = false
  let startX = 0
  let startWidth = 0
  let nextCellStartWidth = 0
  let nextCell = null

  resizeHandle.addEventListener('mousedown', (e) => {
    e.preventDefault()
    e.stopPropagation()
    isResizing = true
    startX = e.pageX
    startWidth = cell.offsetWidth

    // 获取下一个单元格（用于调整间距）
    const row = cell.parentElement
    nextCell = row.cells[cell.cellIndex + 1] || null
    if (nextCell) {
      nextCellStartWidth = nextCell.offsetWidth
    }

    // 创建拖拽指示线
    const dragLine = document.createElement('div')
    dragLine.id = 'column-drag-line'
    dragLine.style.cssText = `
      position: fixed;
      top: ${cell.getBoundingClientRect().top}px;
      left: ${e.pageX}px;
      width: 2px;
      height: ${cell.closest('table').offsetHeight}px;
      background: #2196f3;
      z-index: 1000;
      pointer-events: none;
    `
    document.body.appendChild(dragLine)

    // 添加临时事件监听器
    const handleMouseMove = (e) => {
      if (!isResizing) return

      // 更新拖拽线位置
      const dragLine = document.getElementById('column-drag-line')
      if (dragLine) {
        dragLine.style.left = e.pageX + 'px'
      }

      const diff = e.pageX - startX
      const newWidth = Math.max(50, startWidth + diff) // 最小宽度50px

      // 同步调整整列的宽度
      const table = cell.closest('table')
      const colIndex = cell.cellIndex
      const rows = table.querySelectorAll('tr')

      rows.forEach(row => {
        if (row.cells[colIndex]) {
          row.cells[colIndex].style.width = newWidth + 'px'
        }
        // 如果有下一列，调整其宽度以保持表格总宽度
        if (nextCell && row.cells[colIndex + 1]) {
          const nextNewWidth = Math.max(50, nextCellStartWidth - diff)
          row.cells[colIndex + 1].style.width = nextNewWidth + 'px'
        }
      })
    }

    const handleMouseUp = () => {
      isResizing = false

      // 移除拖拽线
      const dragLine = document.getElementById('column-drag-line')
      if (dragLine) {
        dragLine.remove()
      }

      document.removeEventListener('mousemove', handleMouseMove)
      document.removeEventListener('mouseup', handleMouseUp)

      // 触发内容更新
      const editorElement = document.getElementById('word-editor')
      content.value = editorElement.innerHTML
      hasUnsavedChanges.value = true
    }

    document.addEventListener('mousemove', handleMouseMove)
    document.addEventListener('mouseup', handleMouseUp)
  })

  cell.appendChild(resizeHandle)
}

// 添加列边框拖拽功能
const addColumnBorderResize = (cell) => {
  // 为单元格添加边框悬停检测
  cell.addEventListener('mousemove', (e) => {
    const rect = cell.getBoundingClientRect()
    const distanceFromRightBorder = rect.right - e.clientX
    const distanceFromLeftBorder = e.clientX - rect.left

    // 检查是否靠近右边框（5px范围内）
    if (distanceFromRightBorder <= 5 && distanceFromRightBorder >= 0) {
      cell.style.cursor = 'col-resize'
      cell.setAttribute('data-resize-border', 'right')
    }
    // 检查是否靠近左边框（5px范围内）且不是第一列
    else if (distanceFromLeftBorder <= 5 && distanceFromLeftBorder >= 0 && cell.cellIndex > 0) {
      cell.style.cursor = 'col-resize'
      cell.setAttribute('data-resize-border', 'left')
    }
    else {
      cell.style.cursor = 'text'
      cell.removeAttribute('data-resize-border')
    }
  })

  // 鼠标离开时重置光标
  cell.addEventListener('mouseleave', () => {
    cell.style.cursor = 'text'
    cell.removeAttribute('data-resize-border')
  })

  // 处理边框拖拽
  cell.addEventListener('mousedown', (e) => {
    const resizeBorder = cell.getAttribute('data-resize-border')
    if (!resizeBorder) return

    e.preventDefault()
    e.stopPropagation()

    let targetCell = cell
    let nextCell = null

    if (resizeBorder === 'right') {
      // 拖拽右边框：调整当前列和下一列
      const row = cell.parentElement
      nextCell = row.cells[cell.cellIndex + 1] || null
    } else if (resizeBorder === 'left') {
      // 拖拽左边框：调整前一列和当前列
      const row = cell.parentElement
      targetCell = row.cells[cell.cellIndex - 1]
      nextCell = cell
    }

    if (!targetCell || !nextCell) return

    const startX = e.pageX
    const targetStartWidth = targetCell.offsetWidth
    const nextStartWidth = nextCell.offsetWidth

    // 创建拖拽指示线
    const dragLine = document.createElement('div')
    dragLine.id = 'border-drag-line'
    dragLine.style.cssText = `
      position: fixed;
      top: ${targetCell.getBoundingClientRect().top}px;
      left: ${e.pageX}px;
      width: 2px;
      height: ${targetCell.closest('table').offsetHeight}px;
      background: #f44336;
      z-index: 1000;
      pointer-events: none;
    `
    document.body.appendChild(dragLine)

    const handleMouseMove = (e) => {
      // 更新拖拽线位置
      const dragLine = document.getElementById('border-drag-line')
      if (dragLine) {
        dragLine.style.left = e.pageX + 'px'
      }

      const diff = e.pageX - startX
      const newTargetWidth = Math.max(50, targetStartWidth + diff)
      const newNextWidth = Math.max(50, nextStartWidth - diff)

      // 同步调整整列的宽度
      const table = targetCell.closest('table')
      const targetColIndex = targetCell.cellIndex
      const nextColIndex = nextCell.cellIndex
      const rows = table.querySelectorAll('tr')

      rows.forEach(row => {
        if (row.cells[targetColIndex]) {
          row.cells[targetColIndex].style.width = newTargetWidth + 'px'
        }
        if (row.cells[nextColIndex]) {
          row.cells[nextColIndex].style.width = newNextWidth + 'px'
        }
      })
    }

    const handleMouseUp = () => {
      // 移除拖拽线
      const dragLine = document.getElementById('border-drag-line')
      if (dragLine) {
        dragLine.remove()
      }

      document.removeEventListener('mousemove', handleMouseMove)
      document.removeEventListener('mouseup', handleMouseUp)

      // 重置光标
      cell.style.cursor = 'text'
      cell.removeAttribute('data-resize-border')

      // 触发内容更新
      const editorElement = document.getElementById('word-editor')
      content.value = editorElement.innerHTML
      hasUnsavedChanges.value = true
    }

    document.addEventListener('mousemove', handleMouseMove)
    document.addEventListener('mouseup', handleMouseUp)
  })
}

// 添加行调整大小手柄
const addRowResizeHandle = (cell) => {
  // 检查是否已经有调整手柄
  if (cell.querySelector('.row-resize-handle')) return

  // 创建调整手柄
  const resizeHandle = document.createElement('div')
  resizeHandle.className = 'row-resize-handle'
  resizeHandle.style.cssText = `
    position: absolute;
    left: 0;
    bottom: 0;
    width: 100%;
    height: 5px;
    cursor: row-resize;
    background: transparent;
    z-index: 1;
  `

  // 鼠标悬停时显示视觉效果
  resizeHandle.addEventListener('mouseenter', () => {
    resizeHandle.style.background = 'rgba(76, 175, 80, 0.3)'
  })

  resizeHandle.addEventListener('mouseleave', () => {
    resizeHandle.style.background = 'transparent'
  })

  // 处理拖拽调整大小
  let isResizing = false
  let startY = 0
  let startHeight = 0

  resizeHandle.addEventListener('mousedown', (e) => {
    e.preventDefault()
    e.stopPropagation()
    isResizing = true
    startY = e.pageY

    const row = cell.parentElement
    startHeight = row.offsetHeight

    // 添加临时事件监听器
    const handleMouseMove = (e) => {
      if (!isResizing) return

      const diff = e.pageY - startY
      const newHeight = Math.max(30, startHeight + diff) // 最小高度30px

      // 设置整行的高度
      const row = cell.parentElement
      row.style.height = newHeight + 'px'

      // 同步设置该行所有单元格的高度
      for (let i = 0; i < row.cells.length; i++) {
        row.cells[i].style.height = newHeight + 'px'
      }
    }

    const handleMouseUp = () => {
      isResizing = false
      document.removeEventListener('mousemove', handleMouseMove)
      document.removeEventListener('mouseup', handleMouseUp)

      // 触发内容更新
      const editorElement = document.getElementById('word-editor')
      content.value = editorElement.innerHTML
      hasUnsavedChanges.value = true
    }

    document.addEventListener('mousemove', handleMouseMove)
    document.addEventListener('mouseup', handleMouseUp)
  })

  cell.appendChild(resizeHandle)
}

// 设置动态字段删除功能
const setupDynamicFieldDeletion = () => {
  const editorElement = document.getElementById('word-editor')
  if (!editorElement) return

  // 添加键盘事件监听器
  const handleKeyDown = (e) => {
    // 检查是否按下了删除键（Backspace 或 Delete）
    if (e.key === 'Backspace' || e.key === 'Delete') {
      const selection = window.getSelection()
      if (!selection.rangeCount) return

      const range = selection.getRangeAt(0)
      const container = range.commonAncestorContainer

      // 查找是否在动态字段附近
      let dynamicField = null

      // 如果选中的是文本节点，检查其父元素
      if (container.nodeType === Node.TEXT_NODE) {
        const parent = container.parentElement
        if (parent && parent.classList.contains('dynamic-field')) {
          dynamicField = parent
        }
      }
      // 如果选中的是元素节点，检查是否是动态字段
      else if (container.nodeType === Node.ELEMENT_NODE && container.classList.contains('dynamic-field')) {
        dynamicField = container
      }
      // 检查选区是否包含动态字段
      else if (container.nodeType === Node.ELEMENT_NODE) {
        const fieldsInRange = container.querySelectorAll('.dynamic-field')
        fieldsInRange.forEach(field => {
          if (range.intersectsNode && range.intersectsNode(field)) {
            dynamicField = field
          }
        })
      }

      // 如果找到动态字段，删除它
      if (dynamicField) {
        e.preventDefault()
        dynamicField.remove()

        // 更新内容并触发自动保存
        content.value = editorElement.innerHTML
        hasUnsavedChanges.value = true

        ElMessage.success('动态字段已删除')
      }
    }
  }

  // 为编辑器添加键盘事件监听
  editorElement.addEventListener('keydown', handleKeyDown)

  // 为动态字段添加点击选择功能
  const enhanceDynamicFields = () => {
    const dynamicFields = editorElement.querySelectorAll('.dynamic-field')
    dynamicFields.forEach(field => {
      // 移除旧的事件监听器
      field.removeEventListener('click', selectDynamicField)
      field.removeEventListener('dblclick', deleteDynamicField)

      // 添加点击选择功能
      field.addEventListener('click', selectDynamicField)
      field.addEventListener('dblclick', deleteDynamicField)

      // 添加删除按钮（悬停时显示）
      field.style.position = 'relative'

      field.addEventListener('mouseenter', () => {
        if (!field.querySelector('.delete-btn')) {
          const deleteBtn = document.createElement('span')
          deleteBtn.className = 'delete-btn'
          deleteBtn.innerHTML = '×'
          deleteBtn.style.cssText = `
            position: absolute;
            right: -8px;
            top: -8px;
            width: 16px;
            height: 16px;
            background: #f56c6c;
            color: white;
            border-radius: 50%;
            font-size: 12px;
            line-height: 16px;
            text-align: center;
            cursor: pointer;
            z-index: 1000;
          `

          deleteBtn.addEventListener('click', (e) => {
            e.stopPropagation()
            field.remove()
            content.value = editorElement.innerHTML
            hasUnsavedChanges.value = true
            ElMessage.success('动态字段已删除')
          })

          field.appendChild(deleteBtn)
        }
      })

      field.addEventListener('mouseleave', () => {
        const deleteBtn = field.querySelector('.delete-btn')
        if (deleteBtn) {
          deleteBtn.remove()
        }
      })
    })
  }

  // 选中动态字段
  const selectDynamicField = (e) => {
    e.stopPropagation()
    const field = e.target

    // 创建选区选中整个动态字段
    const selection = window.getSelection()
    const range = document.createRange()
    range.selectNodeContents(field)
    selection.removeAllRanges()
    selection.addRange(range)
  }

  // 双击删除动态字段
  const deleteDynamicField = (e) => {
    e.stopPropagation()
    e.preventDefault()

    const field = e.target
    field.remove()
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true
    ElMessage.success('动态字段已删除')
  }

  // 初始化现有的动态字段
  enhanceDynamicFields()

  // 监听内容变化，为新添加的动态字段增强功能
  const observer = new MutationObserver(() => {
    enhanceDynamicFields()
  })

  observer.observe(editorElement, {
    childList: true,
    subtree: true
  })
}

// 处理单元格鼠标按下（开始拖拽选择）
const handleCellMouseDown = (event) => {
  const cell = event.target.closest('td, th')
  if (!cell || !mergeMode.value) return

  event.preventDefault()
  const row = cell.parentElement.rowIndex
  const col = cell.cellIndex

  dragStart.value = { row, col }
  dragEnd.value = { row, col }

  // 添加选中样式
  updateSelectedCells()
}

// 处理单元格鼠标进入（拖拽选择中）
const handleCellMouseEnter = (event) => {
  const cell = event.target.closest('td, th')
  if (!cell || !mergeMode.value || !dragStart.value) return

  if (event.buttons === 1) { // 鼠标左键按下状态
    const row = cell.parentElement.rowIndex
    const col = cell.cellIndex
    dragEnd.value = { row, col }
    updateSelectedCells()
  }
}

// 处理单元格鼠标释放（结束拖拽选择）
const handleCellMouseUp = (event) => {
  const cell = event.target.closest('td, th')
  if (!cell || !mergeMode.value || !dragStart.value) return

  const row = cell.parentElement.rowIndex
  const col = cell.cellIndex
  dragEnd.value = { row, col }
  updateSelectedCells()

  // 自动执行合并
  if (dragStart.value && dragEnd.value) {
    const rowDiff = Math.abs(dragEnd.value.row - dragStart.value.row)
    const colDiff = Math.abs(dragEnd.value.col - dragStart.value.col)
    if (rowDiff > 0 || colDiff > 0) {
      // 有多个单元格被选中，执行合并
      mergeCells()
    }
  }
}

// 更新选中的单元格样式
const updateSelectedCells = () => {
  // 清除之前的选中样式
  const cells = document.querySelectorAll('.cell-selecting')
  cells.forEach(cell => {
    cell.classList.remove('cell-selecting')
  })

  if (!dragStart.value || !dragEnd.value) return

  const table = document.querySelector('#word-editor table')
  if (!table) return

  const minRow = Math.min(dragStart.value.row, dragEnd.value.row)
  const maxRow = Math.max(dragStart.value.row, dragEnd.value.row)
  const minCol = Math.min(dragStart.value.col, dragEnd.value.col)
  const maxCol = Math.max(dragStart.value.col, dragEnd.value.col)

  // 添加选中样式
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cell = table.rows[r] && table.rows[r].cells[c]
      if (cell) {
        cell.classList.add('cell-selecting')
      }
    }
  }
}

// 处理单元格编辑
const handleCellEdit = (event) => {
  const editorElement = document.getElementById('word-editor')
  content.value = editorElement.innerHTML
  hasUnsavedChanges.value = true

  // 重置自动保存定时器
  if (autoSaveEnabled.value) {
    clearTimeout(autoSaveTimer)
    autoSaveTimer = setTimeout(() => {
      autoSave()
    }, 3000)
  }
}

// 处理单元格失焦
const handleCellBlur = (event) => {
  // 立即触发一次内容更新
  const editorElement = document.getElementById('word-editor')
  content.value = editorElement.innerHTML
  console.log('表格单元格编辑完成')
}

// 切换合并模式
const toggleMergeMode = () => {
  mergeMode.value = !mergeMode.value
  if (!mergeMode.value) {
    // 退出合并模式时清空选择
    clearCellSelection()
  }
  ElMessage.info(mergeMode.value ? '进入合并模式，拖拽选择单元格' : '退出合并模式')
}

// 合并选中的单元格
const mergeCells = () => {
  if (!dragStart.value || !dragEnd.value) {
    ElMessage.warning('请先拖拽选择要合并的单元格')
    return
  }

  try {
    const table = document.querySelector('#word-editor table')
    if (!table) return

    // 计算选择范围
    const minRow = Math.min(dragStart.value.row, dragEnd.value.row)
    const maxRow = Math.max(dragStart.value.row, dragEnd.value.row)
    const minCol = Math.min(dragStart.value.col, dragEnd.value.col)
    const maxCol = Math.max(dragStart.value.col, dragEnd.value.col)

    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1

    // 获取第一个单元格
    const firstCell = table.rows[minRow].cells[minCol]
    if (!firstCell) return

    // 收集所有单元格的内容
    let mergedContent = ''
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const cell = table.rows[r] && table.rows[r].cells[c]
        if (cell && cell.innerHTML.trim()) {
          if (mergedContent) mergedContent += '<br>'
          mergedContent += cell.innerHTML
        }
      }
    }

    // 设置合并属性
    firstCell.setAttribute('rowspan', rowspan)
    firstCell.setAttribute('colspan', colspan)
    firstCell.innerHTML = mergedContent

    // 隐藏其他单元格
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        if (r !== minRow || c !== minCol) {
          const cell = table.rows[r] && table.rows[r].cells[c]
          if (cell) {
            cell.style.display = 'none'
            cell.setAttribute('data-merged', 'true')
            cell.setAttribute('data-merge-parent', `${minRow}_${minCol}`)
            cell.setAttribute('data-row', r.toString())
            cell.setAttribute('data-col', c.toString())
          }
        }
      }
    }

    // 保存合并信息
    const mergeKey = `${minRow}_${minCol}`
    if (!mergedCells.value.table) {
      mergedCells.value.table = {}
    }
    mergedCells.value.table[mergeKey] = {
      startRow: minRow,
      startCol: minCol,
      rowspan: rowspan,
      colspan: colspan
    }

    // 清除选择
    clearCellSelection()
    // 退出合并模式
    mergeMode.value = false

    // 更新内容
    const editorElement = document.getElementById('word-editor')
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true

    ElMessage.success('单元格已合并')
  } catch (error) {
    console.error('合并单元格失败:', error)
    ElMessage.error('合并单元格失败，请重试')
  }
}

// 拆分单元格
const splitCell = (cell) => {
  if (!cell) {
    ElMessage.warning('请先选择要拆分的单元格')
    return
  }

  const rowSpan = parseInt(cell.getAttribute('rowspan')) || 1
  const colSpan = parseInt(cell.getAttribute('colspan')) || 1

  if (rowSpan === 1 && colSpan === 1) {
    ElMessage.warning('该单元格未被合并，无需拆分')
    return
  }

  try {
    const table = cell.closest('table')
    const rowIndex = cell.parentElement.rowIndex
    const colIndex = cell.cellIndex

    // 保存原单元格内容
    const originalContent = cell.innerHTML

    // 重置合并属性
    cell.removeAttribute('rowspan')
    cell.removeAttribute('colspan')
    cell.innerHTML = originalContent // 保留内容在第一个单元格

    // 恢复被隐藏的单元格
    for (let r = rowIndex; r < rowIndex + rowSpan; r++) {
      for (let c = colIndex; c < colIndex + colSpan; c++) {
        if (r !== rowIndex || c !== colIndex) {
          // 找到所有被隐藏的单元格并显示它们
          const allCells = table.querySelectorAll(`td[data-merge-parent="${rowIndex}_${colIndex}"], th[data-merge-parent="${rowIndex}_${colIndex}"]`)
          allCells.forEach(hiddenCell => {
            const hiddenRow = parseInt(hiddenCell.getAttribute('data-row')) || r
            const hiddenCol = parseInt(hiddenCell.getAttribute('data-col')) || c
            if (hiddenRow === r && hiddenCol === c) {
              hiddenCell.style.display = ''
              hiddenCell.removeAttribute('data-merged')
              hiddenCell.removeAttribute('data-merge-parent')
              hiddenCell.innerHTML = '' // 清空内容
              hiddenCell.setAttribute('contenteditable', 'true')
            }
          })

          // 如果没找到隐藏的单元格，创建新的
          const row = table.rows[r]
          if (row && !row.cells[c]) {
            const newCell = document.createElement('td')
            newCell.innerHTML = ''
            newCell.setAttribute('contenteditable', 'true')

            // 找到正确的插入位置
            let insertBeforeCell = null
            for (let i = c + 1; i < row.cells.length; i++) {
              if (!row.cells[i].hasAttribute('data-merged')) {
                insertBeforeCell = row.cells[i]
                break
              }
            }

            if (insertBeforeCell) {
              row.insertBefore(newCell, insertBeforeCell)
            } else {
              row.appendChild(newCell)
            }
          }
        }
      }
    }

    // 删除合并信息
    const mergeKey = `${rowIndex}_${colIndex}`
    if (mergedCells.value.table && mergedCells.value.table[mergeKey]) {
      delete mergedCells.value.table[mergeKey]
    }

    // 重新增强表格编辑功能
    setTimeout(() => {
      enhanceTableEditing()
    }, 100)

    // 更新内容
    const editorElement = document.getElementById('word-editor')
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true

    ElMessage.success('单元格已拆分')
  } catch (error) {
    console.error('拆分单元格失败:', error)
    ElMessage.error('拆分单元格失败，请重试')
  }
}

// 清除单元格选择
const clearCellSelection = () => {
  // 清除所有选中样式
  const cells = document.querySelectorAll('.cell-selected, .cell-selecting')
  cells.forEach(cell => {
    cell.classList.remove('cell-selected', 'cell-selecting')
  })

  dragStart.value = null
  dragEnd.value = null
  selectedCells.value = {}
}

// 添加表格行
const addTableRow = (cell) => {
  try {
    const table = cell.closest('table')
    const currentRow = cell.parentElement
    const rowIndex = currentRow.rowIndex

    // 创建新行
    const newRow = table.insertRow(rowIndex + 1)

    // 复制当前行的高度
    if (currentRow.style.height) {
      newRow.style.height = currentRow.style.height
    } else {
      newRow.style.height = '40px'
      newRow.style.minHeight = '30px'
    }

    // 添加与当前行相同数量的单元格，并复制宽度
    const cellCount = currentRow.cells.length
    for (let i = 0; i < cellCount; i++) {
      const newCell = newRow.insertCell()
      newCell.innerHTML = ''
      newCell.setAttribute('contenteditable', 'true')

      // 复制当前行单元格的宽度和高度
      const currentCell = currentRow.cells[i]
      if (currentCell.style.width) {
        newCell.style.width = currentCell.style.width
      } else {
        // 设置默认宽度
        newCell.style.width = '150px'
        newCell.style.minWidth = '100px'
      }

      // 设置高度和padding
      newCell.style.height = '40px'
      newCell.style.minHeight = '30px'
      newCell.style.padding = '8px'
      newCell.style.verticalAlign = 'middle'
    }

    // 重新增强表格编辑功能
    setTimeout(() => {
      enhanceTableEditing()
    }, 100)

    // 更新内容
    const editorElement = document.getElementById('word-editor')
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true

    ElMessage.success('已添加新行')
  } catch (error) {
    console.error('添加行失败:', error)
    ElMessage.error('添加行失败，请重试')
  }
}

// 添加表格列
const addTableColumn = (cell) => {
  try {
    const table = cell.closest('table')
    const colIndex = cell.cellIndex

    // 为每一行添加新单元格
    for (let i = 0; i < table.rows.length; i++) {
      const row = table.rows[i]
      const newCell = row.insertCell(colIndex + 1)
      newCell.innerHTML = ''
      newCell.setAttribute('contenteditable', 'true')

      // 设置宽度
      newCell.style.width = '150px'
      newCell.style.minWidth = '100px'

      // 设置高度（从同行的其他单元格复制）
      if (row.cells[colIndex] && row.cells[colIndex].style.height) {
        newCell.style.height = row.cells[colIndex].style.height
      } else {
        newCell.style.height = '40px'
        newCell.style.minHeight = '30px'
      }

      newCell.style.padding = '8px'
      newCell.style.verticalAlign = 'middle'
    }

    // 重新增强表格编辑功能
    setTimeout(() => {
      enhanceTableEditing()
    }, 100)

    // 更新内容
    const editorElement = document.getElementById('word-editor')
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true

    ElMessage.success('已添加新列')
  } catch (error) {
    console.error('添加列失败:', error)
    ElMessage.error('添加列失败，请重试')
  }
}

// 删除表格行
const deleteTableRow = (cell) => {
  try {
    const table = cell.closest('table')
    const rowIndex = cell.parentElement.rowIndex

    // 至少保留一行
    if (table.rows.length <= 1) {
      ElMessage.warning('表格至少需要保留一行')
      return
    }

    table.deleteRow(rowIndex)

    // 更新内容
    const editorElement = document.getElementById('word-editor')
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true

    ElMessage.success('已删除行')
  } catch (error) {
    console.error('删除行失败:', error)
    ElMessage.error('删除行失败，请重试')
  }
}

// 删除表格列
const deleteTableColumn = (cell) => {
  try {
    const table = cell.closest('table')
    const colIndex = cell.cellIndex

    // 至少保留一列
    const firstRow = table.rows[0]
    if (firstRow && firstRow.cells.length <= 1) {
      ElMessage.warning('表格至少需要保留一列')
      return
    }

    // 从每一行删除对应的单元格
    for (let i = 0; i < table.rows.length; i++) {
      const row = table.rows[i]
      if (row.cells[colIndex]) {
        row.deleteCell(colIndex)
      }
    }

    // 更新内容
    const editorElement = document.getElementById('word-editor')
    content.value = editorElement.innerHTML
    hasUnsavedChanges.value = true

    ElMessage.success('已删除列')
  } catch (error) {
    console.error('删除列失败:', error)
    ElMessage.error('删除列失败，请重试')
  }
}

// 显示单元格右键菜单
const showCellContextMenu = (event) => {
  event.preventDefault()

  const cell = event.target.closest('td, th')
  if (!cell) return

  // 创建右键菜单
  const menu = document.createElement('div')
  menu.className = 'cell-context-menu'
  menu.style.cssText = `
    position: fixed;
    left: ${event.clientX}px;
    top: ${event.clientY}px;
    background: white;
    border: 1px solid #ddd;
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    padding: 4px 0;
    z-index: 9999;
    min-width: 150px;
  `

  // 菜单项
  const menuItems = []

  // 表格操作菜单
  menuItems.push(
    {
      label: '➕ 在下方插入行',
      action: () => addTableRow(cell)
    },
    {
      label: '➕ 在右侧插入列',
      action: () => addTableColumn(cell)
    },
    {
      label: '➖ 删除当前行',
      action: () => deleteTableRow(cell),
      style: 'color: #f56c6c;'
    },
    {
      label: '➖ 删除当前列',
      action: () => deleteTableColumn(cell),
      style: 'color: #f56c6c;'
    },
    {
      label: '─────────',
      disabled: true,
      style: 'padding: 2px 20px; color: #ddd; cursor: default;'
    }
  )

  // 检查是否是合并的单元格
  const rowspan = parseInt(cell.getAttribute('rowspan')) || 1
  const colspan = parseInt(cell.getAttribute('colspan')) || 1
  const isMerged = rowspan > 1 || colspan > 1

  if (isMerged) {
    menuItems.push({
      label: '↗️ 拆分单元格',
      action: () => splitCell(cell)
    })
  }

  // 添加合并模式切换
  menuItems.push({
    label: mergeMode.value ? '✖️ 退出合并模式' : '🔗 进入合并模式',
    action: () => toggleMergeMode()
  })

  // 添加菜单项
  menuItems.forEach(item => {
    const menuItem = document.createElement('div')
    menuItem.textContent = item.label
    menuItem.style.cssText = `
      padding: 8px 20px;
      cursor: ${item.disabled ? 'default' : 'pointer'};
      color: ${item.disabled ? '#999' : '#333'};
      font-size: 14px;
    `

    if (!item.disabled) {
      menuItem.onmouseover = () => {
        menuItem.style.backgroundColor = '#f0f0f0'
      }
      menuItem.onmouseout = () => {
        menuItem.style.backgroundColor = 'transparent'
      }
      menuItem.onclick = () => {
        item.action()
        document.body.removeChild(menu)
      }
    }

    menu.appendChild(menuItem)
  })

  // 添加到页面
  document.body.appendChild(menu)

  // 点击其他地方关闭菜单
  const closeMenu = (e) => {
    if (!menu.contains(e.target)) {
      document.body.removeChild(menu)
      document.removeEventListener('click', closeMenu)
    }
  }

  setTimeout(() => {
    document.addEventListener('click', closeMenu)
  }, 0)
}

// 监听页面关闭前的未保存提醒
const beforeUnloadHandler = (e) => {
  if (hasUnsavedChanges.value) {
    const message = '您有未保存的更改，确定要离开吗？'
    e.preventDefault()
    e.returnValue = message
    return message
  }
}

onMounted(() => {
  console.log('Component mounted, template ID:', templateId.value)
  loadDataSources()
  // 初始化编辑器
  initEditor()

  // 处理内联数据集占位符交互
  setTimeout(() => {
    handleInlineDatasetInteraction()
  }, 500)

  // 延迟增强表格编辑功能
  setTimeout(() => {
    enhanceTableEditing()
  }, 1000)

  // 添加键盘事件监听器来处理动态字段删除
  setTimeout(() => {
    setupDynamicFieldDeletion()
  }, 1500)

  // 添加页面离开前的提醒
  window.addEventListener('beforeunload', beforeUnloadHandler)
})

// 组件销毁时清理
onUnmounted(() => {
  window.removeEventListener('beforeunload', beforeUnloadHandler)
  if (autoSaveTimer) {
    clearTimeout(autoSaveTimer)
  }
})

// 移除全局表格操作函数，不再需要添加/删除行列的功能
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
    margin: 10px 0;  /* 减小表格上下间距 */
    border: 1px solid #000;
    table-layout: auto;
    word-wrap: break-word;
    font-family: "SimSun", "宋体", "Noto Serif SC", "Source Han Serif SC", serif !important;

    td, th {
      border: 1px solid #000;
      padding: 4px 6px;  /* 减小单元格内边距，从 8px 12px 改为 4px 6px */
      vertical-align: middle;
      text-align: left;
      line-height: 1.4;  /* 稍微减小行高 */
      min-height: 25px;  /* 减小最小高度 */
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
    margin: 4pt 0;  /* 减小段落间距 */
    line-height: 1.5;  /* 减小行高 */
    text-align: justify;
    text-indent: 1.5em;  /* 减小首行缩进 */
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
  
  /* 增强选中效果 */
  ::selection {
    background-color: #b3d4fc;
    color: #000;
  }

  /* 选中的单元格样式 */
  .cell-selected {
    background-color: #e3f2fd !important;
    outline: 2px solid #2196f3 !important;
    position: relative;
  }

  /* 拖拽选择中的单元格样式 */
  .cell-selecting {
    background: linear-gradient(135deg, #e6f7ff 0%, #d4edda 100%) !important;
    border: 2px dashed #52c41a !important;
    animation: merge-highlight 2s ease-in-out infinite;
  }

  /* 表格样式增强 */
  #word-editor :deep(table) {
    table-layout: fixed !important;
    width: 100% !important;
    border-collapse: collapse;
    border: 1px solid #000 !important;
  }

  #word-editor :deep(td),
  #word-editor :deep(th) {
    position: relative !important;
    min-width: 100px !important;
    overflow: hidden;
  }

  /* 列调整手柄 */
  #word-editor :deep(.resize-handle) {
    position: absolute;
    right: -3px;
    top: 0;
    width: 6px;
    height: 100%;
    cursor: col-resize;
    background: transparent;
    z-index: 100;
    transition: background 0.2s;
  }

  #word-editor :deep(.resize-handle:hover) {
    background: rgba(33, 150, 243, 0.5) !important;
  }

  /* 确保表格单元格有边框且可见 */
  #word-editor :deep(table) {
    border-spacing: 0;
  }

  #word-editor :deep(td),
  #word-editor :deep(th) {
    border: 1px solid #000;
    box-sizing: border-box;
  }

  /* 行调整手柄 */
  #word-editor :deep(.row-resize-handle) {
    position: absolute;
    left: 0;
    bottom: 0;
    width: 100%;
    height: 5px;
    cursor: row-resize;
    background: transparent;
    z-index: 10;
    transition: background 0.2s;
  }

  #word-editor :deep(.row-resize-handle:hover) {
    background: rgba(76, 175, 80, 0.5) !important;
  }

  /* 确保单元格有最小高度 */
  #word-editor :deep(td),
  #word-editor :deep(th) {
    min-height: 30px !important;
    vertical-align: middle;
  }

  @keyframes merge-highlight {
    0% {
      background: rgba(82, 196, 26, 0.05);
    }
    50% {
      background: rgba(82, 196, 26, 0.15);
    }
    100% {
      background: rgba(82, 196, 26, 0.05);
    }
  }

  /* 合并模式下的表格样式 */
  .merge-mode-active td,
  .merge-mode-active th {
    cursor: crosshair !important;
    user-select: none;
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

  /* 强制编辑器中所有表格显示黑色边框 */
  #word-editor table {
    border: 1px solid #000 !important;
    border-collapse: collapse !important;
  }

  #word-editor table td,
  #word-editor table th {
    border: 1px solid #000 !important;
  }

  /* 对动态插入的表格也应用黑色边框 */
  #word-editor .dynamic-table table {
    border: 1px solid #000 !important;
  }

  #word-editor .dynamic-table table td,
  #word-editor .dynamic-table table th {
    border: 1px solid #000 !important;
  }
}

/* 数据集占位符样式 */
#word-editor .dataset-placeholder {
  position: relative;
  cursor: not-allowed;
  user-select: none;
  animation: pulse 2s infinite;
}

#word-editor .dataset-placeholder:hover {
  transform: scale(1.05);
  transition: transform 0.2s;
}

/* 数据集表格样式 */
#word-editor .dynamic-table {
  position: relative;
  margin: 15px 0;
  animation: fadeIn 0.5s;
}

#word-editor .dynamic-table:hover {
  box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
  transition: box-shadow 0.3s;
}

/* 动画效果 */
@keyframes pulse {
  0% {
    box-shadow: 0 2px 4px rgba(102, 126, 234, 0.3);
  }
  50% {
    box-shadow: 0 2px 8px rgba(102, 126, 234, 0.5);
  }
  100% {
    box-shadow: 0 2px 4px rgba(102, 126, 234, 0.3);
  }
}

@keyframes fadeIn {
  from {
    opacity: 0;
    transform: translateY(-10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
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

    .auto-save-status {
      margin-left: auto;
      color: #909399;
      font-size: 12px;
      display: flex;
      align-items: center;
      gap: 8px;

      .is-loading {
        animation: rotate 1s linear infinite;
      }
    }
  }

  @keyframes rotate {
    from {
      transform: rotate(0deg);
    }
    to {
      transform: rotate(360deg);
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
      color: #1976d2;
      font-weight: bold;
      display: inline-block;
      min-width: 60px;
      margin: 0 2px;
      cursor: pointer;
      user-select: none;
      position: relative;
      transition: all 0.2s ease;
    }

    :deep(.dynamic-field:hover) {
      background-color: #bbdefb;
      border-color: #1565c0;
      transform: scale(1.02);
    }

    :deep(.dynamic-field.selected) {
      background-color: #1976d2;
      color: white;
      border-color: #0d47a1;
    }

    :deep(.delete-btn) {
      position: absolute !important;
      right: -8px !important;
      top: -8px !important;
      width: 16px !important;
      height: 16px !important;
      background: #f56c6c !important;
      color: white !important;
      border-radius: 50% !important;
      font-size: 12px !important;
      line-height: 16px !important;
      text-align: center !important;
      cursor: pointer !important;
      z-index: 1000 !important;
      transition: background 0.2s ease !important;
    }

    :deep(.delete-btn:hover) {
      background: #e53e3e !important;
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