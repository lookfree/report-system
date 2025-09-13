<template>
  <div class="template-config-page" v-loading="loading">
    <el-card>
      <template #header>
        <div class="card-header">
          <span>模板配置: {{ template?.name }}</span>
          <div>
            <el-button @click="$router.back()">返回</el-button>
            <el-button type="primary" @click="handleSaveConfig" :loading="saving">
              保存配置
            </el-button>
          </div>
        </div>
      </template>
      
      <div v-if="template">
        <div v-for="section in template.structure.sections" :key="section.id" class="section-config">
          <!-- 默认表格不显示头部信息配置 -->
          <div v-if="section.id !== 'default_header_table'" class="table-header-config" style="background: #f0f0f0; padding: 10px; margin: 10px 0; border-left: 4px solid #1890ff;">
            <div style="display: flex; align-items: center; justify-content: space-between; gap: 15px;">
              <div style="flex: 1; display: flex; align-items: center;">
                <label style="font-weight: bold; margin-right: 10px; white-space: nowrap;">名称:</label>
                <el-input 
                  v-model="tableNames[section.id]" 
                  style="flex: 1; max-width: 400px;" 
                  placeholder="请输入表格名称"
                  size="small"
                />
                <h3 v-if="!tableNames[section.id]" style="margin-left: 15px; margin-right: 0; margin-top: 0; margin-bottom: 0; font-size: 16px; color: #666;">{{ section.title }}</h3>
              </div>
            </div>
            
            <!-- 备注配置 -->
            <div style="margin-top: 15px;">
              <label style="font-weight: bold; margin-right: 10px;">备注:</label>
              <el-input 
                v-model="tableRemarks[section.id]"
                placeholder="请输入备注"
                size="small"
                style="max-width: 400px;"
              />
            </div>
            
            <!-- 统计配置 -->
            <div v-if="section.hasTable" style="margin-top: 15px; display: flex; align-items: center; gap: 10px;">
              <label style="font-weight: bold;">统计:</label>
              <el-switch 
                v-model="statisticsEnabled[section.id]"
                size="small"
              />
            </div>
            
          </div>
          
          <!-- 如果是表格类型，显示动态表格配置 -->
          <div v-if="section.hasTable && section.tableStructure?.headers?.length > 0">
            <el-alert 
              v-if="section.id !== 'default_header_table'"
              type="info" 
              :title="`检测到表格 (${section.tableStructure.rowCount} 行 × ${section.tableStructure.columnCount} 列)`"
              :closable="false"
              style="margin-bottom: 15px;"
            />
            
            <!-- 表格配置 -->
            <div class="table-config">
              <div class="config-header">
                <h4>表格配置</h4>
              </div>
              
              <!-- 表格操作工具栏 - 简化版 -->
              <div class="table-operations-simple">
                <div class="operation-buttons">
                  <el-button 
                    size="small" 
                    type="primary"
                    @click="toggleMergeMode(section.id)"
                    :class="{ 'merge-mode-active': mergeMode[section.id] }">
                    <el-icon><svg viewBox="0 0 1024 1024" width="16" height="16"><path d="M896 128H128c-35.3 0-64 28.7-64 64v640c0 35.3 28.7 64 64 64h768c35.3 0 64-28.7 64-64V192c0-35.3-28.7-64-64-64zM448 832H128V576h320v256zm0-320H128V256h320v256zm448 320H512V576h384v256zm0-320H512V256h384v256z"/></svg></el-icon>
                    {{ mergeMode[section.id] ? '取消' : '合并单元格' }}
                  </el-button>
                  <el-button 
                    size="small" 
                    type="warning" 
                    @click="splitSelectedCell(section.id)"
                    :disabled="!hasSelectedMergedCell(section.id)">
                    <el-icon><svg viewBox="0 0 1024 1024" width="16" height="16"><path d="M896 128H128c-35.3 0-64 28.7-64 64v640c0 35.3 28.7 64 64 64h768c35.3 0 64-28.7 64-64V192c0-35.3-28.7-64-64-64zM192 768V576h256v192H192zm0-256V256h256v256H192zm320 256V576h256v192H512zm0-256V256h256v256H512z"/></svg></el-icon>
                    拆分单元格
                  </el-button>
                </div>
                <div class="operation-tips" v-if="mergeMode[section.id]">
                  <el-icon class="el-icon--primary" style="animation: pulse-border 1.5s ease-in-out infinite;">
                    <svg viewBox="0 0 1024 1024" width="16" height="16">
                      <path d="M512 64C264.6 64 64 264.6 64 512s200.6 448 448 448 448-200.6 448-448S759.4 64 512 64zm0 820c-205.4 0-372-166.6-372-372s166.6-372 372-372 372 166.6 372 372-166.6 372-372 372z"/>
                      <path d="M464 336a48 48 0 1 0 96 0 48 48 0 1 0-96 0zm72 112h-48c-4.4 0-8 3.6-8 8v272c0 4.4 3.6 8 8 8h48c4.4 0 8-3.6 8-8V456c0-4.4-3.6-8-8-8z"/>
                    </svg>
                  </el-icon>
                  <span style="margin-left: 8px; color: #52c41a; font-weight: 500;">
                    🖱️ 按住鼠标拖拽选择要合并的单元格区域
                  </span>
                </div>
              </div>


              <!-- 纯数据表格实现 - 无表头概念 -->
              <div class="modern-table-container">
                <table class="modern-table data-table" :class="{ 'merge-mode-active-table': mergeMode[section.id] }">
                  <tbody>
                    <!-- 所有行都作为数据行，包括原来的表头 -->
                    <tr v-for="(row, rowIndex) in getAllTableRows(section.id)" :key="rowIndex" 
                        :class="{ 'header-row': rowIndex < getHeaderRowCount(section.id) }">
                      <td v-for="(cellData, colIndex) in row" 
                          :key="colIndex"
                          v-show="!isCellHidden(section.id, rowIndex, colIndex)"
                          :id="`cell_${section.id}_${rowIndex}_${colIndex}`"
                          class="table-cell"
                          :class="getCellClass(section.id, rowIndex, colIndex)"
                          :colspan="getCellColspan(section.id, rowIndex, colIndex)"
                          :rowspan="getCellRowspan(section.id, rowIndex, colIndex)"
                          @mousedown="handleCellMouseDown(section.id, rowIndex, colIndex, $event)"
                          @mouseenter="handleCellMouseEnter(section.id, rowIndex, colIndex, $event)"
                          @mouseup="handleCellMouseUp(section.id, rowIndex, colIndex, $event)"
                          @click="handleCellClickEvent(section.id, rowIndex, colIndex, $event)"
                          @dblclick="handleCellDoubleClick(section.id, rowIndex, colIndex, $event)">
                        <input 
                          v-if="editingCell === `${section.id}_${rowIndex}_${colIndex}`"
                          v-model="editingValue"
                          class="cell-input"
                          @blur="confirmEdit(section.id, rowIndex, colIndex)"
                          @keyup.enter="confirmEdit(section.id, rowIndex, colIndex)"
                          @keyup.esc="cancelEdit"
                          ref="cellInput"
                        />
                        <span v-else>
                          {{ getCellDisplayContent(section.id, rowIndex, colIndex, cellData || '') }}
                        </span>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              
              <!-- 统计行 -->
              <div v-if="statisticsEnabled[section.id] && section.id !== 'default_header_table'" class="statistics-container">
                <table class="statistics-table">
                  <thead>
                    <tr>
                      <th colspan="100%" class="stats-header">统计汇总</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td v-for="(header, colIndex) in section.tableStructure?.headers || []"
                          :key="colIndex"
                          :id="`stats_cell_${section.id}_${colIndex}`"
                          class="stats-cell"
                          @click="handleStatsCellClick(section.id, colIndex)">
                        {{ getStatsContent(section.id, colIndex) }}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          
          <!-- 非表格类型的常规配置 -->
          <div v-else>
            <el-table :data="[getConfigForSection(section.id)]" style="width: 100%" border stripe>
              <el-table-column label="数据类型" width="150">
                <template #default="{ row }">
                  <el-select v-model="row.dataType" placeholder="选择类型">
                    <el-option label="固定值" value="FIXED" />
                    <el-option label="手动填写" value="MANUAL" />
                    <el-option label="动态获取" value="DYNAMIC" />
                  </el-select>
                </template>
              </el-table-column>
              
              <el-table-column label="配置内容">
                <template #default="{ row }">
                  <!-- Fixed or Manual value -->
                  <el-input
                    v-if="row.dataType === 'FIXED' || row.dataType === 'MANUAL'"
                    v-model="row.value"
                    :placeholder="row.dataType === 'FIXED' ? '输入固定值' : '运行时手动填写'"
                    type="textarea"
                    :rows="2"
                  />
                  
                  <!-- Dynamic SQL -->
                  <div v-else-if="row.dataType === 'DYNAMIC'">
                    <el-select v-model="row.dataSourceId" placeholder="选择数据源" style="width: 100%; margin-bottom: 10px;">
                      <el-option
                        v-for="ds in dataSources"
                        :key="ds.id"
                        :label="ds.name"
                        :value="ds.id"
                      />
                    </el-select>
                    <el-input
                      v-model="row.sqlQuery"
                      placeholder="输入SQL查询语句"
                      type="textarea"
                      :rows="3"
                    />
                    <el-button
                      v-if="row.dataSourceId && row.sqlQuery"
                      size="small"
                      style="margin-top: 10px;"
                      @click="testQuery(row)"
                    >
                      测试查询
                    </el-button>
                  </div>
                </template>
              </el-table-column>
            </el-table>
          </div>
        </div>
      </div>
    </el-card>
    
    <!-- Query Test Dialog -->
    <el-dialog v-model="showTestDialog" title="查询测试结果" width="600px">
      <div v-loading="testing">
        <el-alert v-if="testResult.error" type="error" :title="testResult.error" show-icon />
        <div v-else-if="testResult.data">
          <el-alert type="success" title="查询成功" show-icon :closable="false" />
          <p style="margin: 10px 0;">返回 {{ testResult.rowCount }} 行数据</p>
          <el-table :data="testResult.sample" size="small" max-height="300" border stripe>
            <el-table-column
              v-for="(val, key) in testResult.sample[0]"
              :key="key"
              :prop="key"
              :label="key"
            >
            </el-table-column>
          </el-table>
        </div>
      </div>
    </el-dialog>

    <!-- 单元格配置弹窗 -->
    <el-dialog v-model="showCellConfigDialog" title="单元格配置" width="800px">
      <el-form :model="cellConfig" label-width="100px">
        <!-- 展示内容 -->
        <el-form-item label="展示内容">
          <el-radio-group v-model="cellConfig.displayType">
            <el-radio label="text">仅文字</el-radio>
            <el-radio label="dataset">数据集</el-radio>
          </el-radio-group>
        </el-form-item>

        <!-- 仅文字模式 -->
        <div v-if="cellConfig.displayType === 'text'">
          <el-form-item label="文字内容">
            <el-input v-model="cellConfig.textContent" placeholder="请输入文字内容" />
          </el-form-item>
        </div>

        <!-- 数据集模式 -->
        <div v-if="cellConfig.displayType === 'dataset'">
          <el-form-item label="数据集">
            <el-select v-model="cellConfig.datasetId" placeholder="请选择" style="width: 200px;">
              <el-option
                v-for="ds in dataSources"
                :key="ds.id"
                :label="ds.name"
                :value="ds.id"
              />
            </el-select>
          </el-form-item>

          <el-form-item label="数据结构">
            <el-radio-group v-model="cellConfig.dataStructure">
              <el-radio label="single">单条</el-radio>
              <el-radio label="list" :disabled="isMergedCell || isHiddenCell">列表</el-radio>
            </el-radio-group>
            <div v-if="isMergedCell || isHiddenCell" class="field-tip">
              <el-text type="warning" size="small">
                <el-icon><Warning /></el-icon>
                合并单元格不支持列表模式
              </el-text>
            </div>
          </el-form-item>

          <!-- 列表模式的sheet页配置 -->
          <div v-if="cellConfig.dataStructure === 'list'">
            <el-form-item label="sheet页配置">
              <el-radio-group v-model="cellConfig.sheetConfig">
                <el-radio label="current">当前sheet页</el-radio>
                <el-radio label="separate">单独sheet页</el-radio>
              </el-radio-group>
            </el-form-item>
          </div>

          <el-form-item label="展示字段" v-if="cellConfig.dataStructure === 'single' || cellConfig.dataStructure === 'list'">
            <el-select 
              v-model="cellConfig.displayFields" 
              multiple
              placeholder="请选择字段（可多选）" 
              style="width: 400px;"
              :max-collapse-tags="3"
              collapse-tags
              collapse-tags-tooltip
            >
              <el-option-group label="用户权限字段">
                <el-option label="数据批次(字符串)" value="data_batch" />
                <el-option label="组织编号(字符串)" value="org_code" />
                <el-option label="组织名称(字符串)" value="org_name" />
                <el-option label="父级机构编号(字符串)" value="parent_org_code" />
                <el-option label="姓名(字符串)" value="person_name" />
                <el-option label="主账号类型(字符串)" value="account_type" />
                <el-option label="电话(字符串)" value="phone" />
                <el-option label="录入时间(字符串)" value="entry_time" />
              </el-option-group>
              <el-option-group label="系统日志字段">
                <el-option label="日志ID(字符串)" value="log_id" />
                <el-option label="日志时间(时间)" value="log_time" />
                <el-option label="日志级别(字符串)" value="log_level" />
                <el-option label="模块名称(字符串)" value="module" />
                <el-option label="日志消息(字符串)" value="message" />
                <el-option label="IP地址(字符串)" value="ip_address" />
                <el-option label="原始日志(字符串)" value="originalLog" />
              </el-option-group>
              <el-option-group label="业务数据字段">
                <el-option label="业务ID(字符串)" value="business_id" />
                <el-option label="客户名称(字符串)" value="customer_name" />
                <el-option label="合同金额(数字)" value="contract_amount" />
                <el-option label="合同日期(日期)" value="contract_date" />
                <el-option label="状态(字符串)" value="status" />
                <el-option label="负责人(字符串)" value="manager" />
                <el-option label="部门(字符串)" value="department" />
              </el-option-group>
              <el-option-group label="财务数据字段">
                <el-option label="账户代码(字符串)" value="account_code" />
                <el-option label="账户名称(字符串)" value="account_name" />
                <el-option label="余额(数字)" value="balance" />
                <el-option label="货币类型(字符串)" value="currency" />
                <el-option label="更新时间(日期)" value="last_update" />
                <el-option label="账户类型(字符串)" value="account_type" />
              </el-option-group>
              <el-option-group label="人事数据字段">
                <el-option label="员工ID(字符串)" value="employee_id" />
                <el-option label="员工姓名(字符串)" value="employee_name" />
                <el-option label="部门名称(字符串)" value="department" />
                <el-option label="职位(字符串)" value="position" />
                <el-option label="入职日期(日期)" value="hire_date" />
                <el-option label="薪资(数字)" value="salary" />
                <el-option label="员工状态(字符串)" value="status" />
              </el-option-group>
            </el-select>
          </el-form-item>

          <el-form-item label="字段预览" v-if="cellConfig.displayFields && cellConfig.displayFields.length > 0">
            <!-- 列表模式：表格预览 -->
            <div class="field-preview-table" v-if="cellConfig.dataStructure === 'list'">
              <el-table 
                :data="getPreviewData()" 
                style="width: 100%; max-height: 200px;" 
                size="small"
                border
                stripe
              >
                <el-table-column 
                  v-for="field in cellConfig.displayFields" 
                  :key="field"
                  :prop="field"
                  :label="getFieldLabel(field)"
                  :width="120"
                  show-overflow-tooltip
                />
              </el-table>
              <p class="preview-note">
                <el-icon><Document /></el-icon>
                已选择 {{ cellConfig.displayFields.length }} 个字段，预览前 3 条数据
              </p>
            </div>
            
            <!-- 单条模式：字段值预览 -->
            <div class="field-preview-single" v-if="cellConfig.dataStructure === 'single'">
              <div class="single-data-preview">
                <div 
                  v-for="field in cellConfig.displayFields" 
                  :key="field"
                  class="field-item"
                >
                  <span class="field-label">{{ getFieldLabel(field) }}：</span>
                  <span class="field-value">{{ getSingleFieldValue(field) }}</span>
                </div>
              </div>
              <p class="preview-note">
                <el-icon><Document /></el-icon>
                已选择 {{ cellConfig.displayFields.length }} 个字段，单条数据预览
              </p>
            </div>
          </el-form-item>
        </div>
      </el-form>

      <template #footer>
        <span class="dialog-footer">
          <el-button @click="showCellConfigDialog = false">取消</el-button>
          <el-button type="primary" @click="confirmCellConfig">确定</el-button>
        </span>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted, nextTick, onUnmounted, watch } from 'vue'
import { useRoute } from 'vue-router'
import { ElMessage } from 'element-plus'
import { Plus, Delete, Document, Warning } from '@element-plus/icons-vue'
import RevoGrid from '@revolist/vue3-datagrid'
import { useTableConfig } from '@/composables/useTableConfig'
import { useTemplateConfig } from '@/composables/useTemplateConfig'
import { useGridHelpers } from '@/composables/useGridHelpers'
import api from '@/api'

const route = useRoute()
const templateId = ref(route.params.id)

// 使用composable

const {
  contentRows,
  getContentRows,
  onContentRowChange
} = useTableConfig()

const {
  template,
  loading,
  saving,
  configs,
  dataSources,
  showTestDialog,
  testing,
  testResult,
  getConfigForSection,
  onColumnConfigChange,
  resetTableData,
  loadDataSources,
  saveConfig,
  testQuery,
  testColumnQuery,
  loadTemplate
} = useTemplateConfig()

const {
  hasMergedHeaders,
  getMergedHeaderGroups,
  getGridColumns,
  getGridData,
  onCellEdit,
  onBeforeEdit,
  onRowHeaderClick,
  onColHeaderClick
} = useGridHelpers()

// 新增的响应式数据
const tableNames = ref({})
const tableDataSources = ref({})
const tableRemarks = ref({})
const statisticsEnabled = ref({})
// 单元格配置存储
const cellConfigurations = ref({})
// 统计行配置存储
const statisticsConfigurations = ref({})
// 选中的单元格
const selectedCells = ref({})
// 合并的单元格信息
const mergedCells = ref({})
// 合并模式状态
const mergeMode = ref({})
// 框选起始位置
const dragStart = ref(null)
// 框选结束位置
const dragEnd = ref(null)


// 单元格配置弹窗相关
const showCellConfigDialog = ref(false)
const selectedCellInfo = ref(null)
const isMergedCell = ref(false)
const isHiddenCell = ref(false) // 是否为被合并的单元格（非主单元格）
const cellConfig = ref({
  displayType: 'text', // 'text' | 'dataset'
  textContent: '', // 仅文字模式的内容
  datasetId: '', // 数据集ID
  dataStructure: 'single', // 'single' | 'list' | 'attachment'
  sheetConfig: 'current', // 'current' | 'separate'
  displayFields: [], // 展示字段数组（支持多选）
  displayField: '', // 保持向后兼容
  // 保持向后兼容
  fillType: '',
  fixedValue: '',
  linkedField: ''
})

// 监听template变化，自动填充表格名称
watch(template, (newTemplate) => {
  if (newTemplate && newTemplate.structure && newTemplate.structure.sections) {
    const names = {}
    newTemplate.structure.sections.forEach(section => {
      if (section.hasTable && section.title) {
        names[section.id] = section.title
      }
    })
    tableNames.value = names
  }
}, { immediate: true })

// 从保存的配置中恢复数据
const restoreConfigData = () => {
  if (!configs.value || configs.value.length === 0) return
  
  // 恢复表格名称
  const tableNameConfigs = configs.value.filter(c => c.sectionId.endsWith('_name'))
  tableNameConfigs.forEach(config => {
    const sectionId = config.parentSectionId
    if (sectionId) {
      tableNames.value[sectionId] = config.value
    }
  })
  
  // 恢复数据源配置
  const dataSourceConfigs = configs.value.filter(c => c.sectionId.endsWith('_datasource'))
  dataSourceConfigs.forEach(config => {
    const sectionId = config.parentSectionId
    if (sectionId) {
      tableDataSources.value[sectionId] = config.value
    }
  })
  
  // 恢复备注配置
  const remarkConfigs = configs.value.filter(c => c.sectionId.endsWith('_remark'))
  remarkConfigs.forEach(config => {
    const sectionId = config.parentSectionId
    if (sectionId) {
      tableRemarks.value[sectionId] = config.value
    }
  })
  
  // 恢复统计开关配置
  const statsConfigs = configs.value.filter(c => c.sectionId.endsWith('_stats'))
  statsConfigs.forEach(config => {
    const sectionId = config.parentSectionId
    if (sectionId) {
      statisticsEnabled.value[sectionId] = config.value === 'true'
    }
  })
  
  // 清空原有单元格配置
  cellConfigurations.value = {}
  
  // 恢复单元格配置
  const cellConfigs = configs.value.filter(c => c.sectionId.startsWith('cell_'))
  cellConfigs.forEach(config => {
    const cellKey = config.sectionId.replace('cell_', '')
    const [sectionId, rowIndex, colIndex] = cellKey.split('_')
    
    // 统一配置格式
    const configData = {
      displayType: config.dataType === 'FIXED' ? 'text' : 'dataset',
      textContent: config.dataType === 'FIXED' ? config.value || '' : '',
      datasetId: config.dataType === 'DYNAMIC' ? (config.dataSourceId || '1') : '',
      dataStructure: 'single',
      sheetConfig: 'current',
      displayFields: [],
      displayField: '',
      
      // 保持兼容性字段
      fillType: config.dataType === 'FIXED' ? 'fixed' : 'auto',
      fixedValue: config.value || '',
      linkedField: ''
    }
    
    // 解析数据集配置
    if (config.dataType === 'DYNAMIC' && config.value && config.value.includes('|')) {
      const [structure, fieldsStr] = config.value.split('|')
      configData.dataStructure = structure || 'single'
      if (fieldsStr && fieldsStr.includes(',')) {
        configData.displayFields = fieldsStr.split(',').map(f => f.trim())
      } else if (fieldsStr) {
        configData.displayFields = [fieldsStr.trim()]
      }
      configData.displayField = configData.displayFields[0] || ''
      configData.linkedField = configData.displayField
    }
    
    cellConfigurations.value[cellKey] = configData
    
    // 页面显示通过 getCellDisplayContent 函数处理，无需手动更新DOM
  })
  
  // 恢复统计行配置
  const statsRowConfigs = configs.value.filter(c => c.sectionId.startsWith('stats_'))
  statsRowConfigs.forEach(config => {
    const cellKey = config.sectionId
    const [, sectionId, colIndex] = cellKey.split('_')
    
    statisticsConfigurations.value[cellKey] = {
      fillType: config.dataType === 'FIXED' ? 'fixed' : 'auto',
      fixedValue: config.value || '',
      linkedField: config.sqlQuery ? config.sqlQuery.replace('SELECT ', '').replace(' FROM table', '') : ''
    }
    
    // 稍后更新统计行显示（需要等待DOM渲染）
    setTimeout(() => {
      const statsElement = document.getElementById(`stats_cell_${sectionId}_${colIndex}`)
      if (statsElement) {
        const displayContent = config.dataType === 'FIXED' ? config.value : '自动汇总'
        statsElement.textContent = displayContent
        
        if (config.dataType === 'FIXED') {
          statsElement.style.backgroundColor = '#e6f7ff'
          statsElement.style.color = '#1890ff'
        } else {
          statsElement.style.backgroundColor = '#f0f9ff'
          statsElement.style.color = '#52c41a'
        }
      }
    }, 100)
  })
  
  // 恢复合并单元格信息
  const mergeConfigs = configs.value.filter(c => c.sectionId.startsWith('merge_'))
  mergedCells.value = {}
  
  console.log('开始恢复合并单元格信息，找到配置:', mergeConfigs.length, '个')
  
  mergeConfigs.forEach(config => {
    try {
      // 解析sectionId，处理多种格式:
      // 1. merge_table_section_0_3_0 (完整格式)
      // 2. merge_table_0_0 (简化格式)
      // 3. merge_table_3_0 (简化格式)
      const parts = config.sectionId.split('_')
      
      let sectionId, mergeKey
      
      if (parts.length >= 3 && parts[0] === 'merge') {
        if (parts[1] === 'table' && parts.length >= 4) {
          // 格式: merge_table_xxx 或 merge_table_section_xxx
          if (parts[2] === 'section') {
            // merge_table_section_0_3_0 -> sectionId: table_section_0, mergeKey: 3_0
            sectionId = `table_section_${parts[3]}`
            mergeKey = parts.slice(4).join('_')
          } else {
            // merge_table_3_0 -> sectionId: table_section_0, mergeKey: 3_0
            // 假设默认是 table_section_0
            sectionId = 'table_section_0'
            mergeKey = parts.slice(2).join('_')
          }
        } else {
          // 标准格式: merge_sectionId_mergeKey
          sectionId = parts[1]
          mergeKey = parts.slice(2).join('_')
        }
        
        // 解析合并信息
        const mergeInfo = JSON.parse(config.value)
        
        if (!mergedCells.value[sectionId]) {
          mergedCells.value[sectionId] = {}
        }
        
        // 为合并区域的所有单元格设置合并信息
        for (let row = mergeInfo.startRow; row < mergeInfo.startRow + mergeInfo.rowspan; row++) {
          for (let col = mergeInfo.startCol; col < mergeInfo.startCol + mergeInfo.colspan; col++) {
            const cellKey = `${row}_${col}`
            mergedCells.value[sectionId][cellKey] = {
              startRow: mergeInfo.startRow,
              startCol: mergeInfo.startCol,
              rowspan: mergeInfo.rowspan,
              colspan: mergeInfo.colspan
            }
          }
        }
        
      } else {
        console.warn('无法解析合并单元格配置格式:', config.sectionId)
      }
    } catch (error) {
      console.warn('恢复合并单元格配置失败:', config.sectionId, error)
    }
  })
  
  // 如果没有保存过的合并配置，从模版结构中生成初始合并配置（仅用于表头行）
  if (mergeConfigs.length === 0 && template.value?.structure?.sections) {
    template.value.structure.sections.forEach(section => {
      if (section.hasTable && section.tableStructure?.headers) {
        initializeHeaderMergedCells(section.id)
      }
    })
  }
  
  // 强制触发Vue重新渲染，确保合并单元格样式生效
  if (Object.keys(mergedCells.value).length > 0) {
    nextTick(() => {
      // 合并单元格样式生效
    })
  }
  
  // 恢复配置数据完成
}

// 恢复表格数据
const restoreTableData = (template, getContentRows) => {
  if (!template?.structure?.sections) return
  
  // 开始恢复表格数据
  
  template.structure.sections.forEach(section => {
    if (section.hasTable && section.tableData) {
      
      // 清空现有数据
      const currentRows = getContentRows(section.id)
      currentRows.splice(0, currentRows.length)
      
      // 获取当前表头列数（这是扩展后的列数）
      const expectedColumnCount = section.tableStructure?.headers?.length || 0
      console.log(`表格 ${section.id} 期望列数: ${expectedColumnCount}`)
      
      // 恢复保存的数据，确保行数据与扩展后的表头匹配
      section.tableData.forEach(savedRow => {
        const cells = savedRow.cells || []
        
        // 确保每行的单元格数量与扩展后的表头数量匹配
        while (cells.length < expectedColumnCount - 1) { // -1 因为第一列通常是名称列
          cells.push('') // 用空字符串填充新增的列
        }
        
        currentRows.push({
          name: savedRow.name || '',
          type: savedRow.type || 'data',
          cells: cells,
          locked: savedRow.locked || false
        })
      })
      
      // 确保表格结构的行数和列数与实际数据一致
      if (section.tableStructure) {
        section.tableStructure.rowCount = currentRows.length
        section.tableStructure.columnCount = expectedColumnCount
      }
      
      console.log(`表格 ${section.id} 数据恢复完成，共 ${currentRows.length} 行，每行 ${expectedColumnCount} 列`)
      console.log(`表格 ${section.id} 结构同步：rowCount=${section.tableStructure?.rowCount}, columnCount=${section.tableStructure?.columnCount}`)
    }
  })
  
  console.log('所有表格数据和结构恢复完成')
}


onMounted(() => {
  loadTemplate(templateId.value, getContentRows, (template, getContentRows) => {
    // RevoGrid会自动渲染，无需手动初始化
    console.log('Template loaded, RevoGrid will render automatically')
    
    // 恢复保存的表格数据
    restoreTableData(template, getContentRows)
    
    // 合并单元格信息将从后台保存的配置中恢复，不再使用模版header结构初始化
    console.log('🚀 表格加载完成，合并单元格信息将从配置数据中恢复')
    
    // 延迟恢复配置数据，确保表格已渲染
    setTimeout(() => {
      restoreConfigData()
    }, 500)
  })
  loadDataSources()
})

onUnmounted(() => {
  // RevoGrid会自动清理，无需手动销毁
  console.log('Component unmounting, RevoGrid will auto cleanup')
})

// 单元格配置相关方法
const openCellConfigDialog = (sectionId, rowIndex, colIndex, cellData) => {
  selectedCellInfo.value = {
    sectionId,
    rowIndex,
    colIndex
  }
  
  // 检查单元格是否被合并
  isMergedCell.value = isCellMerged(sectionId, rowIndex, colIndex)
  // 检查单元格是否被隐藏（被合并的单元格）
  isHiddenCell.value = isCellHidden(sectionId, rowIndex, colIndex)
  
  // 初始化弹窗数据
  if (cellData) {
    console.log('收到的cellData:', cellData)
    
    // 确保 displayFields 始终是数组
    let displayFields = []
    if (cellData.displayFields && Array.isArray(cellData.displayFields)) {
      displayFields = cellData.displayFields
    } else if (cellData.displayField) {
      displayFields = [cellData.displayField]
    } else if (cellData.linkedField) {
      displayFields = [cellData.linkedField]
    }
    
    cellConfig.value = {
      displayType: cellData.displayType || 'text',
      textContent: cellData.textContent || cellData.fixedValue || '',
      datasetId: cellData.datasetId || '1', // 默认数据源
      dataStructure: isHiddenCell.value && cellData.dataStructure === 'list' 
        ? 'single' // 被合并的单元格强制使用单条模式
        : cellData.dataStructure || 'single',
      sheetConfig: cellData.sheetConfig || 'current',
      displayFields: displayFields,
      displayField: displayFields[0] || '', // 向后兼容
      // 向后兼容
      fillType: cellData.fillType || '',
      fixedValue: cellData.fixedValue || '',
      linkedField: cellData.linkedField || ''
    }
    console.log('设置的cellConfig:', cellConfig.value)
    console.log('displayFields数组:', cellConfig.value.displayFields)
  } else {
    cellConfig.value = {
      displayType: 'text',
      textContent: '',
      datasetId: '',
      dataStructure: 'single', // 默认为单条
      sheetConfig: 'current',
      displayFields: [],
      displayField: '', // 向后兼容
      fillType: '',
      fixedValue: '',
      linkedField: ''
    }
  }
  
  showCellConfigDialog.value = true
}

// 调试已移除

const confirmCellConfig = () => {
  if (!selectedCellInfo.value) return
  
  const { sectionId, rowIndex, colIndex, isStats } = selectedCellInfo.value
  
  // 根据是否为统计行决定存储位置和更新元素
  if (isStats) {
    const cellKey = `stats_${sectionId}_${colIndex}`
    statisticsConfigurations.value[cellKey] = { ...cellConfig.value }
    
    // 更新统计行单元格显示
    const statsElement = document.getElementById(`stats_cell_${sectionId}_${colIndex}`)
    if (statsElement) {
      let displayContent = ''
      if (cellConfig.value.fillType === 'fixed') {
        displayContent = cellConfig.value.fixedValue || '统计值'
      } else if (cellConfig.value.fillType === 'auto') {
        displayContent = '动态汇总'
      }
      statsElement.textContent = displayContent
      
      // 设置样式
      if (cellConfig.value.fillType === 'fixed') {
        statsElement.style.backgroundColor = '#e6f7ff'
        statsElement.style.color = '#1890ff'
      } else if (cellConfig.value.fillType === 'auto') {
        statsElement.style.backgroundColor = '#f0f9ff'
        statsElement.style.color = '#52c41a'
      }
    }
  } else {
    const cellKey = `${sectionId}_${rowIndex}_${colIndex}`
    
    // 存储单元格配置
    cellConfigurations.value[cellKey] = { ...cellConfig.value }
    
    // 确保底层数据结构存在，但不直接修改内容
    // 内容显示交给 getCellDisplayContent 函数处理
    const rows = getContentRows(sectionId)
    if (rows && rows[rowIndex]) {
      if (!rows[rowIndex].cells) {
        rows[rowIndex].cells = []
      }
      
      // 确保单元格位置存在，但内容由 getCellDisplayContent 控制显示
      if (rows[rowIndex].cells[colIndex] === undefined) {
        rows[rowIndex].cells[colIndex] = ''
      }
    }
  }
  
  console.log('保存单元格配置:', {
    cellInfo: selectedCellInfo.value,
    config: cellConfig.value,
    isStats
  })
  
  // 如果是普通单元格（非统计行）且有内容配置，检查是否需要自动扩展
  if (!isStats) {
    const cellKey = `${sectionId}_${rowIndex}_${colIndex}`
    const config = cellConfigurations.value[cellKey]
    
    // 检查是否有有效的配置内容
    let hasContent = false
    if (config) {
      if (config.displayType === 'text' && config.textContent) {
        hasContent = config.textContent.trim() !== ''
      } else if (config.displayType === 'dataset' && config.displayField) {
        hasContent = true
      } else if (config.fillType === 'fixed' && config.fixedValue) {
        hasContent = config.fixedValue.trim() !== ''
      } else if (config.fillType === 'auto' && config.linkedField) {
        hasContent = true
      }
    }
    
    console.log(`🔥 操作第${rowIndex + 1}行第${colIndex + 1}列, hasContent=${hasContent}`)
    
    if (hasContent) {
      autoAddColumnIfNeeded(sectionId, colIndex)
      autoAddRowIfNeeded(sectionId, rowIndex)
    }
  }
  
  showCellConfigDialog.value = false
  ElMessage.success('单元格配置保存成功')
}

// 旧的handleCellClick函数已被删除，使用新的框选版本

// 统计行相关方法
const handleStatsCellClick = (sectionId, colIndex) => {
  const cellKey = `stats_${sectionId}_${colIndex}`
  const existingConfig = statisticsConfigurations.value[cellKey]
  openStatsCellConfigDialog(sectionId, colIndex, existingConfig)
}

const openStatsCellConfigDialog = (sectionId, colIndex, cellData) => {
  selectedCellInfo.value = {
    sectionId,
    rowIndex: 'stats',
    colIndex,
    isStats: true
  }
  
  // 初始化弹窗数据
  cellConfig.value = {
    fillType: cellData?.fillType || '',
    fixedValue: cellData?.fixedValue || '',
    linkedField: cellData?.linkedField || ''
  }
  
  showCellConfigDialog.value = true
}

const getStatsContent = (sectionId, colIndex) => {
  const cellKey = `stats_${sectionId}_${colIndex}`
  const config = statisticsConfigurations.value[cellKey]
  
  if (config) {
    if (config.fillType === 'fixed') {
      return config.fixedValue || '统计值'
    } else if (config.fillType === 'auto') {
      return '动态汇总'
    }
  }
  
  return '点击配置'
}

// 获取单元格显示内容
const getCellDisplayContent = (sectionId, rowIndex, colIndex, originalContent) => {
  const cellKey = `${sectionId}_${rowIndex}_${colIndex}`
  const config = cellConfigurations.value[cellKey]
  
  // 如果有配置，根据配置显示内容
  if (config) {
    if (config.displayType === 'text') {
      // 仅文字模式
      return config.textContent || ''
    } else if (config.displayType === 'dataset' && config.displayField) {
      // 数据集模式
      if (config.dataStructure === 'list') {
        return `[列表数据: ${config.displayField}]`
      } else {
        // 单条数据
        return getFieldDisplayValue(config.displayField)
      }
    }
    
    // 向后兼容旧的配置方式
    if (!config.displayType && config.fillType === 'fixed' && config.fixedValue) {
      return config.fixedValue
    } else if (!config.displayType && config.fillType === 'auto' && config.linkedField) {
      return getFieldDisplayValue(config.linkedField)
    }
  }
  
  // 如果没有配置或配置无效，返回原始内容
  return originalContent || ''
}

// 获取单元格样式类
const getCellClass = (sectionId, rowIndex, colIndex) => {
  const cellKey = `${sectionId}_${rowIndex}_${colIndex}`
  const config = cellConfigurations.value[cellKey]
  
  let classes = []
  
  // 配置状态样式
  if (config) {
    if (config.fillType === 'fixed' || config.displayType === 'text') {
      classes.push('cell-configured-fixed')
    } else if (config.fillType === 'auto' || config.displayType === 'dataset') {
      classes.push('cell-configured-auto')
    }
  } else {
    classes.push('cell-not-configured')
  }
  
  // 拖拽选择中的预览效果
  if (dragStart.value && dragEnd.value && mergeMode.value[sectionId]) {
    const minRow = Math.min(dragStart.value.row, dragEnd.value.row)
    const maxRow = Math.max(dragStart.value.row, dragEnd.value.row)
    const minCol = Math.min(dragStart.value.col, dragEnd.value.col)
    const maxCol = Math.max(dragStart.value.col, dragEnd.value.col)
    
    if (rowIndex >= minRow && rowIndex <= maxRow && colIndex >= minCol && colIndex <= maxCol) {
      classes.push('cell-selecting')
    }
  }
  
  // 选中状态样式
  const selected = selectedCells.value[sectionId]
  if (selected && selected.some(cell => cell.row === rowIndex && cell.col === colIndex)) {
    classes.push('cell-selected')
  }
  
  // 合并状态样式
  const mergeKey = `${rowIndex}_${colIndex}`
  if (mergedCells.value[sectionId] && mergedCells.value[sectionId][mergeKey]) {
    classes.push('cell-merged')
  }
  
  return classes.join(' ')
}

// 获取字段显示值
const getFieldDisplayValue = (fieldKey) => {
  const fieldMap = {
    // 用户权限字段
    'data_batch': '2024-12',
    'org_code': 'ORG001',
    'org_name': '总公司',
    'parent_org_code': '',
    'person_name': '张三',
    'account_type': '管理员账户',
    'phone': '13800138000',
    'entry_time': '2024-01-15',
    
    // 系统日志字段
    'log_id': 'LOG001',
    'log_time': '2024-12-10 08:30:00',
    'log_level': 'INFO',
    'module': '用户登录',
    'message': '用户张三成功登录系统',
    'ip_address': '192.168.1.100',
    
    // 业务数据字段
    'business_id': 'BUS001',
    'customer_name': '阿里巴巴集团',
    'contract_amount': '1,500,000.00',
    'contract_date': '2024-01-15',
    'status': '执行中',
    'manager': '张三',
    'department': '销售部',
    
    // 财务数据字段
    'account_code': 'ACC001',
    'account_name': '银行存款',
    'balance': '5,680,000.00',
    'currency': 'CNY',
    'last_update': '2024-12-09',
    
    // 人事数据字段
    'employee_id': 'EMP001',
    'employee_name': '张三',
    'position': '高级工程师',
    'hire_date': '2020-01-15',
    'salary': '25,000.00'
  }
  return fieldMap[fieldKey] || '示例数据'
}

// 处理保存配置
const handleSaveConfig = async () => {
  try {
    saving.value = true
    
    // 准备完整的结构数据，包括表格内容
    const structureWithContent = JSON.parse(JSON.stringify(template.value.structure))
    
    console.log('即将保存的完整结构:', JSON.stringify(structureWithContent, null, 2))
    
    // 将表格内容数据合并到结构中
    if (structureWithContent.sections) {
      structureWithContent.sections.forEach(section => {
        if (section.hasTable) {
          const rows = getContentRows(section.id)
          console.log(`保存表格 ${section.id}:`)
          console.log(`  当前表头数量: ${section.tableStructure?.headers?.length || 0}`)
          console.log(`  当前行数: ${rows ? rows.length : 0}`)
          console.log(`  表格结构: rowCount=${section.tableStructure?.rowCount}, columnCount=${section.tableStructure?.columnCount}`)
          
          if (rows && rows.length > 0) {
            // 将实际的表格数据保存到结构中
            section.tableData = rows.map(row => ({
              name: row.name || '',
              type: row.type || 'data',
              cells: row.cells || [],
              locked: row.locked || false
            }))
            
            console.log(`  保存的tableData: ${section.tableData.length}行`)
            section.tableData.forEach((row, index) => {
              console.log(`    行${index}: name="${row.name}", cells数量=${row.cells.length}`)
            })
          }
        }
      })
    }
    
    // 统一保存所有配置数据和完整结构
    await saveConfig(templateId.value, {
      tableNames: tableNames.value,
      tableDataSources: tableDataSources.value,
      tableRemarks: tableRemarks.value,
      statisticsEnabled: statisticsEnabled.value,
      cellConfigurations: cellConfigurations.value,
      statisticsConfigurations: statisticsConfigurations.value,
      mergedCells: mergedCells.value, // 添加合并单元格信息
      templateStructure: structureWithContent // 传递包含表格数据的完整结构
    })
    
    ElMessage.success('配置保存成功！')
  } catch (error) {
    console.error('保存失败:', error)
    ElMessage.error('保存失败: ' + error.message)
  } finally {
    saving.value = false
  }
}



const handleGridCellClick = (event) => {
  console.log('Grid cell clicked:', event)
  // 这里可以添加单元格点击的处理逻辑
}


// HTML表格编辑功能
const editingCell = ref('')
const editingValue = ref('')

const handleTableCellClick = (sectionId, rowIndex, colIndex) => {
  console.log('Table cell clicked:', sectionId, rowIndex, colIndex)
  // 单击不再打开弹窗，只用于调试日志
}

const editCell = (sectionId, rowIndex, colIndex) => {
  const rows = getContentRows(sectionId)
  editingCell.value = `${sectionId}_${rowIndex}_${colIndex}`
  
  editingValue.value = rows[rowIndex].cells?.[colIndex] || ''
  
  // 聚焦到输入框
  nextTick(() => {
    const input = document.querySelector('.cell-input')
    if (input) {
      input.focus()
      input.select()
    }
  })
}

const confirmEdit = (sectionId, rowIndex, colIndex) => {
  const rows = getContentRows(sectionId)
  
  if (!rows[rowIndex].cells) {
    rows[rowIndex].cells = []
  }
  
  // 保存新内容
  rows[rowIndex].cells[colIndex] = editingValue.value
  
  editingCell.value = ''
  editingValue.value = ''
}

const cancelEdit = () => {
  editingCell.value = ''
  editingValue.value = ''
}

// 切换合并模式
const toggleMergeMode = (sectionId) => {
  if (mergeMode.value[sectionId]) {
    mergeMode.value[sectionId] = false
    selectedCells.value[sectionId] = []
  } else {
    mergeMode.value[sectionId] = true
  }
}

// 鼠标按下事件 - 开始框选
const handleCellMouseDown = (sectionId, rowIndex, colIndex, event) => {
  if (mergeMode.value[sectionId]) {
    event.preventDefault()
    dragStart.value = { sectionId, row: rowIndex, col: colIndex }
    dragEnd.value = { sectionId, row: rowIndex, col: colIndex }
    updateSelectedCells(sectionId)
  }
}

// 鼠标进入事件 - 更新框选范围
const handleCellMouseEnter = (sectionId, rowIndex, colIndex, event) => {
  if (mergeMode.value[sectionId] && dragStart.value && event.buttons === 1) {
    dragEnd.value = { sectionId, row: rowIndex, col: colIndex }
    updateSelectedCells(sectionId)
  }
}

// 鼠标释放事件 - 完成框选
const handleCellMouseUp = (sectionId, rowIndex, colIndex, event) => {
  if (mergeMode.value[sectionId] && dragStart.value) {
    dragEnd.value = { sectionId, row: rowIndex, col: colIndex }
    updateSelectedCells(sectionId)
    // 自动执行合并
    if (selectedCells.value[sectionId] && selectedCells.value[sectionId].length > 1) {
      performMerge(sectionId)
    }
    // 重置拖拽状态
    dragStart.value = null
    dragEnd.value = null
  }
}

// 更新选中的单元格
const updateSelectedCells = (sectionId) => {
  if (!dragStart.value || !dragEnd.value) return
  
  const minRow = Math.min(dragStart.value.row, dragEnd.value.row)
  const maxRow = Math.max(dragStart.value.row, dragEnd.value.row)
  const minCol = Math.min(dragStart.value.col, dragEnd.value.col)
  const maxCol = Math.max(dragStart.value.col, dragEnd.value.col)
  
  selectedCells.value[sectionId] = []
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      selectedCells.value[sectionId].push({
        key: `${r}_${c}`,
        row: r,
        col: c
      })
    }
  }
}

// 自动添加新列 - 当操作到最后一列时自动扩展
const autoAddColumnIfNeeded = (sectionId, colIndex) => {
  if (!template.value?.structure?.sections) return
  
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (!section?.tableStructure) return
  
  // 重新获取当前实际的行数据（因为弹窗回填时可能已经动态修改过）
  const rows = getContentRows(sectionId)
  if (!rows || rows.length === 0) return
  
  // 重新计算实际的最大列数（基于实际数据）
  let actualMaxColumns = 0
  rows.forEach(row => {
    const cellCount = row.cells ? row.cells.length : 0
    if (cellCount > actualMaxColumns) {
      actualMaxColumns = cellCount
    }
  })
  
  // 重新获取表头的列数和配置的列数
  const headerCount = section.tableStructure.headers?.length || 0
  const configuredColumnCount = section.tableStructure.columnCount || 0
  const maxColumnCount = Math.max(actualMaxColumns, headerCount, configuredColumnCount)
  
  // 如果操作的是最后一列，自动添加新列
  console.log(`检查自动添加列: colIndex=${colIndex}, headerCount=${headerCount}, 是否为最后一列: ${colIndex === headerCount - 1}`)
  if (colIndex === headerCount - 1) {
    const newColumnCount = headerCount + 1
    
    // 更新表头
    if (!section.tableStructure.headers) {
      section.tableStructure.headers = []
    }
    
    // 确保表头数组有足够的长度
    while (section.tableStructure.headers.length < newColumnCount) {
      const newHeaderIndex = section.tableStructure.headers.length
      const newHeaderName = `列${newHeaderIndex + 1}`
      
      // 创建完整的header对象结构，保持与原始header对象一致
      section.tableStructure.headers.push({
        index: newHeaderIndex,
        name: newHeaderName,
        originalName: newHeaderName,
        parentHeader: null,
        dataType: "FIXED",
        value: "",
        sqlQuery: "",
        dataSourceId: null
      })
    }
    
    section.tableStructure.columnCount = newColumnCount
    
    // 为所有行添加新列的空单元格
    rows.forEach(row => {
      if (!row.cells) row.cells = []
      // 确保每行有足够的单元格
      while (row.cells.length < newColumnCount) {
        row.cells.push('')
      }
    })
    
    console.log(`自动添加新列：${section.name} 现在有 ${newColumnCount} 列`)
  }
}

// 自动添加新行 - 当操作到最后一行时自动扩展
const autoAddRowIfNeeded = (sectionId, rowIndex) => {
  if (!template.value?.structure?.sections) return
  
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (!section?.tableStructure) return
  
  // 重新获取当前全局表格行数（包含表头和数据行）
  const allRows = getAllTableRows(sectionId)
  if (!allRows || allRows.length === 0) return
  
  const totalRowCount = allRows.length
  
  console.log(`🔍 行扩展检查: 第${rowIndex + 1}行(索引${rowIndex}) 全局共${totalRowCount}行 是否最后一行:${rowIndex === totalRowCount - 1}`)
  
  if (rowIndex === totalRowCount - 1) {
    // 获取数据行来添加新行
    const dataRows = getContentRows(sectionId)
    
    // 获取当前最大列数
    let maxColumns = section.tableStructure.columnCount || 0
    const headerCount = section.tableStructure.headers?.length || 0
    maxColumns = Math.max(maxColumns, headerCount)
    
    // 计算实际需要的列数
    dataRows.forEach(row => {
      const cellCount = row.cells ? row.cells.length : 0
      if (cellCount > maxColumns) {
        maxColumns = cellCount
      }
    })
    
    // 创建新行，包含所有列的空单元格
    const newRow = {
      name: `行${dataRows.length + 1}`,
      type: 'data',
      cells: new Array(maxColumns).fill(''),
      locked: false
    }
    
    // 添加新行到数据中
    dataRows.push(newRow)
    
    // 更新表格行数
    section.tableStructure.rowCount = dataRows.length
    
    console.log(`自动添加新行：${section.name} 现在有 ${dataRows.length} 行数据`)
  }
}

// 单击事件 - 选择单个单元格或退出合并模式
const handleCellClickEvent = (sectionId, rowIndex, colIndex, event) => {
  if (!mergeMode.value[sectionId]) {
    // 非合并模式：选择单个单元格
    selectedCells.value[sectionId] = [{
      key: `${rowIndex}_${colIndex}`,
      row: rowIndex,
      col: colIndex
    }]
  }
}

// 处理双击事件
const handleCellDoubleClick = (sectionId, rowIndex, colIndex, event) => {
  event.preventDefault()
  
  if (!mergeMode.value[sectionId]) {
    // 双击时打开配置弹窗
    const cellKey = `${sectionId}_${rowIndex}_${colIndex}`
    const existingConfig = cellConfigurations.value[cellKey]
    console.log('打开单元格配置:', { cellKey, existingConfig })
    openCellConfigDialog(sectionId, rowIndex, colIndex, existingConfig)
  }
}

// 开始编辑单元格
const startCellEdit = (sectionId, rowIndex, colIndex) => {
  const cellKey = `${sectionId}_${rowIndex}_${colIndex}`
  editingCell.value = cellKey
  
  // 获取当前单元格的内容
  const rows = getContentRows(sectionId)
  const currentValue = rows[rowIndex]?.cells?.[colIndex] || ''
  editingValue.value = currentValue
  
  // 延迟聚焦到输入框
  nextTick(() => {
    const input = document.querySelector(`input[v-model="editingValue"]`)
    if (input) {
      input.focus()
      input.select()
    }
  })
}


// 检查单元格是否是合并单元格的主单元格
const isCellMerged = (sectionId, rowIndex, colIndex) => {
  if (!mergedCells.value[sectionId]) return false
  
  const cellKey = `${rowIndex}_${colIndex}`
  return !!mergedCells.value[sectionId][cellKey]
}

// 检查单元格是否被隐藏（被合并的单元格）
const isCellHidden = (sectionId, rowIndex, colIndex) => {
  if (!mergedCells.value[sectionId]) return false
  
  for (const mergeKey in mergedCells.value[sectionId]) {
    const merge = mergedCells.value[sectionId][mergeKey]
    if (rowIndex >= merge.startRow && rowIndex < merge.startRow + merge.rowspan &&
        colIndex >= merge.startCol && colIndex < merge.startCol + merge.colspan &&
        !(rowIndex === merge.startRow && colIndex === merge.startCol)) {
      return true
    }
  }
  return false
}

// 获取单元格的colspan
const getCellColspan = (sectionId, rowIndex, colIndex) => {
  if (!mergedCells.value[sectionId]) {
    return 1
  }
  
  const cellKey = `${rowIndex}_${colIndex}`
  const merge = mergedCells.value[sectionId][cellKey]
  
  return merge ? merge.colspan : 1
}

// 获取单元格的rowspan
const getCellRowspan = (sectionId, rowIndex, colIndex) => {
  if (!mergedCells.value[sectionId]) {
    return 1
  }
  
  const cellKey = `${rowIndex}_${colIndex}`
  const merge = mergedCells.value[sectionId][cellKey]
  
  return merge ? merge.rowspan : 1
}

// 执行合并操作
// 更新表头名称
const updateHeaderName = (sectionId, index, newName) => {
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (section && section.tableStructure && section.tableStructure.headers[index]) {
    section.tableStructure.headers[index].name = newName
  }
}

// 更新合并表头的名称
const updateMergedHeaderName = (sectionId, group, newName) => {
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (!section || !section.tableStructure) return
  
  // 如果是独立列，直接更新
  if (group.type === 'independent') {
    const header = section.tableStructure.headers.find(h => h.name === group.name)
    if (header) {
      header.name = newName
    }
  } else {
    // 如果是父表头，更新所有子表头的parentHeader
    section.tableStructure.headers.forEach(header => {
      if (header.parentHeader === group.name) {
        header.parentHeader = newName
      }
    })
  }
}

// === 无表头纯数据表格支持 ===

// 获取所有表格行数据（表头作为数据行 + 原内容行）
const getAllTableRows = (sectionId) => {
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (!section?.tableStructure?.headers) return []
  
  const allRows = []
  
  // 1. 添加表头作为普通数据行（简化处理）
  if (hasMergedHeaders(section.tableStructure.headers)) {
    // 有合并表头：生成两行
    const groups = getMergedHeaderGroups(section.tableStructure.headers)
    
    // 第一行：父表头
    const firstRow = []
    groups.forEach(group => {
      firstRow.push(group.name || '')
      // 合并列的其他位置填空
      for (let i = 1; i < group.colspan; i++) {
        firstRow.push('')
      }
    })
    allRows.push(firstRow)
    
    // 第二行：子表头
    const secondRow = section.tableStructure.headers.map(header => {
      if (!header.parentHeader || header.parentHeader.trim() === '') {
        return '' // 独立列在第二行为空
      }
      return header.name || ''
    })
    allRows.push(secondRow)
  } else {
    // 简单表头：生成一行
    const headerRow = section.tableStructure.headers.map(header => header.name || '')
    allRows.push(headerRow)
  }
  
  // 2. 添加原有的内容数据行（保持原逻辑不变）
  const contentRows = getContentRows(sectionId)
  contentRows.forEach(row => {
    const rowData = section.tableStructure.headers.map((header, index) => {
      if (header.name === '审计单位') {
        return row.name || ''
      }
      return row.cells?.[index] || ''
    })
    allRows.push(rowData)
  })
  
  return allRows
}


// 获取表头数据行的数量
const getHeaderRowCount = (sectionId) => {
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (!section?.tableStructure?.headers) return 0
  
  // 如果有合并表头，返回2行，否则返回1行
  return hasMergedHeaders(section.tableStructure.headers) ? 2 : 1
}

// 初始化表头行的合并单元格信息
const initializeHeaderMergedCells = (sectionId) => {
  const section = template.value.structure.sections.find(s => s.id === sectionId)
  if (!section?.tableStructure?.headers) {
    console.log(`No headers found for section ${sectionId}`)
    return
  }
  
  if (!mergedCells.value[sectionId]) {
    mergedCells.value[sectionId] = {}
  }
  
  console.log(`🔧 INITIALIZING HEADER MERGED CELLS FOR SECTION ${sectionId}`)
  console.log('📋 Headers:', section.tableStructure.headers)
  
  // 只处理有合并表头的情况
  if (hasMergedHeaders(section.tableStructure.headers)) {
    const groups = getMergedHeaderGroups(section.tableStructure.headers)
    console.log('Header groups:', groups)
    
    let currentCol = 0
    
    groups.forEach(group => {
      if (group.type === 'independent') {
        // 独立列：跨两行
        const mergeKey = `0_${currentCol}`
        mergedCells.value[sectionId][mergeKey] = {
          startRow: 0,
          startCol: currentCol,
          rowspan: 2,
          colspan: 1
        }
        console.log(`Created independent merge at ${mergeKey}:`, mergedCells.value[sectionId][mergeKey])
        currentCol++
      } else {
        // 合并列的父表头：跨多列
        const mergeKey = `0_${currentCol}`
        mergedCells.value[sectionId][mergeKey] = {
          startRow: 0,
          startCol: currentCol,
          rowspan: 1,
          colspan: group.colspan
        }
        console.log(`Created parent merge at ${mergeKey}:`, mergedCells.value[sectionId][mergeKey])
        currentCol += group.colspan
      }
    })
    
    console.log(`Final merged cells for ${sectionId}:`, mergedCells.value[sectionId])
  } else {
    console.log(`Section ${sectionId} has no merged headers`)
  }
}


const performMerge = (sectionId) => {
  const selected = selectedCells.value[sectionId]
  if (!selected || selected.length < 2) {
    return
  }
  
  // 计算合并范围
  const minRow = Math.min(...selected.map(cell => cell.row))
  const maxRow = Math.max(...selected.map(cell => cell.row))
  const minCol = Math.min(...selected.map(cell => cell.col))
  const maxCol = Math.max(...selected.map(cell => cell.col))
  
  const rowspan = maxRow - minRow + 1
  const colspan = maxCol - minCol + 1
  
  if (!mergedCells.value[sectionId]) {
    mergedCells.value[sectionId] = {}
  }
  
  const mergeKey = `${minRow}_${minCol}`
  mergedCells.value[sectionId][mergeKey] = {
    startRow: minRow,
    startCol: minCol,
    rowspan,
    colspan
  }
  
  // 清除选择并退出合并模式
  selectedCells.value[sectionId] = []
  mergeMode.value[sectionId] = false
  
  ElMessage.success('单元格合并成功')
}

// 拆分选中的单元格
const splitSelectedCell = (sectionId) => {
  const selected = selectedCells.value[sectionId]
  if (!selected || selected.length !== 1) {
    ElMessage.warning('请选择一个已合并的单元格进行拆分')
    return
  }
  
  const cell = selected[0]
  const mergeKey = `${cell.row}_${cell.col}`
  
  if (!mergedCells.value[sectionId] || !mergedCells.value[sectionId][mergeKey]) {
    ElMessage.warning('该单元格未被合并')
    return
  }
  
  delete mergedCells.value[sectionId][mergeKey]
  selectedCells.value[sectionId] = []
  
  ElMessage.success('单元格拆分成功')
}

// 检查是否有选中的合并单元格
const hasSelectedMergedCell = (sectionId) => {
  const selected = selectedCells.value[sectionId]
  if (!selected || selected.length !== 1 || !mergedCells.value[sectionId]) return false
  
  const cell = selected[0]
  const mergeKey = `${cell.row}_${cell.col}`
  return !!mergedCells.value[sectionId][mergeKey]
}

// 获取字段标签
const getFieldLabel = (field) => {
  const fieldLabels = {
    // 用户权限字段
    data_batch: '数据批次',
    org_code: '组织编号',
    org_name: '组织名称',
    parent_org_code: '父级机构编号',
    person_name: '姓名',
    account_type: '主账号类型',
    phone: '电话',
    entry_time: '录入时间',
    // 系统日志字段
    log_id: '日志ID',
    log_time: '日志时间',
    log_level: '日志级别',
    module: '模块名称',
    message: '日志消息',
    ip_address: 'IP地址',
    user_agent: '用户代理',
    // 业务数据字段
    business_id: '业务ID',
    customer_name: '客户名称',
    contract_amount: '合同金额',
    contract_date: '合同日期',
    status: '状态',
    manager: '负责人',
    department: '部门',
    // 财务数据字段
    account_code: '账户代码',
    account_name: '账户名称',
    balance: '余额',
    currency: '货币类型',
    last_update: '更新时间',
    account_type: '账户类型',
    // 人事数据字段
    employee_id: '员工ID',
    employee_name: '员工姓名',
    hire_date: '入职日期',
    position: '职位',
    salary: '薪资'
  }
  return fieldLabels[field] || field
}

// 获取单条字段值
const getSingleFieldValue = (field) => {
  const mockData = getPreviewData()
  if (mockData && mockData.length > 0) {
    return mockData[0][field] || '无数据'
  }
  return '无数据'
}

// 获取预览数据
const getPreviewData = () => {
  // 返回mock数据的前3条用于预览
  const mockData = [
    {
      data_batch: '2024-12',
      org_code: 'ORG001',
      org_name: '总公司',
      parent_org_code: '',
      person_name: '张三',
      account_type: '管理员账户',
      phone: '13800138000',
      entry_time: '2024-01-15',
      log_id: 'LOG001',
      log_time: '2024-12-10 08:30:00',
      log_level: 'INFO',
      module: '用户登录',
      message: '用户张三成功登录系统',
      ip_address: '192.168.1.100',
      user_agent: 'Mozilla/5.0...',
      originalLog: '{"createtime":"2025-09-09 16:52:37","dataid":"ark-gateway-prod.yml","elapseTime":123,"level":"INFO","message":"用户登录成功"}',
      business_id: 'BUS001',
      customer_name: '阿里巴巴集团',
      contract_amount: 1500000.00,
      contract_date: '2024-01-15',
      status: '执行中',
      manager: '张三',
      department: '销售部',
      account_code: 'ACC001',
      account_name: '银行存款',
      balance: 5680000.00,
      currency: 'CNY',
      last_update: '2024-12-09',
      employee_id: 'EMP001',
      employee_name: '张三',
      hire_date: '2020-01-15',
      position: '高级工程师',
      salary: 25000.00
    },
    {
      data_batch: '2024-12',
      org_code: 'ORG002',
      org_name: '分公司A',
      parent_org_code: 'ORG001',
      person_name: '李四',
      account_type: '普通用户',
      phone: '13900139000',
      entry_time: '2024-03-10',
      log_id: 'LOG002',
      log_time: '2024-12-10 09:15:00',
      log_level: 'WARN',
      module: '权限验证',
      message: '用户李四尝试访问无权限页面',
      ip_address: '192.168.1.101',
      user_agent: 'Chrome/118.0...',
      originalLog: '{"createtime":"2025-09-09 16:50:56","dataid":"ark-gateway-prod.yml","elapseTime":256,"level":"WARN","message":"权限验证失败"}',
      business_id: 'BUS002',
      customer_name: '腾讯科技',
      contract_amount: 2800000.00,
      contract_date: '2024-03-20',
      status: '已完成',
      manager: '李四',
      department: '商务部',
      account_code: 'ACC002',
      account_name: '应收账款',
      balance: 2340000.00,
      currency: 'CNY',
      last_update: '2024-12-09',
      employee_id: 'EMP002',
      employee_name: '李四',
      hire_date: '2019-06-01',
      position: '销售经理',
      salary: 22000.00
    },
    {
      data_batch: '2024-12',
      org_code: 'ORG003',
      org_name: '部门B',
      parent_org_code: 'ORG002',
      person_name: '王五',
      account_type: '部门主管',
      phone: '13700137000',
      entry_time: '2024-02-20',
      log_id: 'LOG003',
      log_time: '2024-12-10 10:45:00',
      log_level: 'ERROR',
      module: '数据库连接',
      message: '数据库连接超时',
      ip_address: '192.168.1.10',
      user_agent: 'System',
      originalLog: '{"createtime":"2025-09-09 16:50:56","dataid":"ark-gateway-prod.yml","elapseTime":5000,"level":"ERROR","message":"数据库连接超时异常"}',
      business_id: 'BUS003',
      customer_name: '字节跳动',
      contract_amount: 3200000.00,
      contract_date: '2024-05-10',
      status: '执行中',
      manager: '王五',
      department: '技术部',
      account_code: 'ACC003',
      account_name: '现金',
      balance: 890000.00,
      currency: 'CNY',
      last_update: '2024-12-09',
      employee_id: 'EMP003',
      employee_name: '王五',
      hire_date: '2021-03-15',
      position: '项目经理',
      salary: 28000.00
    }
  ]
  
  return mockData
}

// 设置全局函数，供HTML表格的onclick使用
window.handleCellClick = handleCellClickEvent
</script>

<style scoped lang="scss">
.template-config-page {
  .card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 16px;
    
    h3 {
      margin: 0;
      color: #303133;
      font-weight: 600;
    }
    
    .actions {
      display: flex;
      gap: 12px;
    }
  }
  
  .section {
    margin-bottom: 30px;
    border: 1px solid #e4e7ed;
    border-radius: 8px;
    overflow: hidden;
    background: white;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
    
    .section-header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 16px 20px;
      cursor: pointer;
      user-select: none;
      position: relative;
      
      &:hover {
        background: linear-gradient(135deg, #5a67d8 0%, #6b46c1 100%);
      }
      
      .section-title {
        font-size: 16px;
        font-weight: 600;
        margin: 0;
      }
      
      .section-info {
        margin-top: 6px;
        font-size: 13px;
        opacity: 0.9;
        display: flex;
        align-items: center;
        gap: 16px;
        
        .info-item {
          display: flex;
          align-items: center;
          
          .el-icon {
            margin-right: 4px;
          }
        }
      }
    }
    
    .section-content {
      padding: 20px;
      
      .no-table-content {
        text-align: center;
        color: #909399;
        padding: 40px;
        
        .el-icon {
          font-size: 48px;
          color: #d3d4d6;
          margin-bottom: 12px;
        }
      }
    }
  }
  
  .table-operations {
    display: flex;
    gap: 8px;
    margin-bottom: 16px;
    flex-wrap: wrap;
    
    .el-button {
      font-size: 13px;
      padding: 6px 12px;
      
      .el-icon {
        margin-right: 4px;
      }
    }
  }

  // 简化版表格操作工具栏
  .table-operations-simple {
    margin-bottom: 16px;
    
    .operation-buttons {
      display: flex;
      gap: 12px;
      align-items: center;
      margin-bottom: 8px;
      
      .el-button {
        font-size: 14px;
        padding: 8px 16px;
        border-radius: 6px;
        transition: all 0.2s ease;
        
        .el-icon {
          margin-right: 6px;
          
          svg {
            width: 16px;
            height: 16px;
            fill: currentColor;
          }
        }
        
        &.merge-mode-active {
          background: linear-gradient(135deg, #52c41a 0%, #73d13d 100%);
          border-color: #52c41a;
          color: white;
          
          &:hover {
            background: linear-gradient(135deg, #389e0d 0%, #52c41a 100%);
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(82, 196, 26, 0.3);
          }
        }
        
        &:not(.merge-mode-active):hover {
          transform: translateY(-1px);
          box-shadow: 0 4px 12px rgba(24, 144, 255, 0.3);
        }
        
        &.el-button--warning:hover {
          box-shadow: 0 4px 12px rgba(255, 193, 7, 0.3);
        }
      }
    }
    
    .operation-tips {
      font-size: 14px;
      color: #52c41a;
      background: linear-gradient(135deg, #f6ffed 0%, #d9f7be 100%);
      border: 2px solid #52c41a;
      border-radius: 6px;
      padding: 8px 16px;
      animation: fadeIn 0.3s ease, pulse-border 2s ease-in-out infinite;
      display: flex;
      align-items: center;
      box-shadow: 0 2px 8px rgba(82, 196, 26, 0.15);
      
      span {
        display: flex;
        align-items: center;
        
        &::before {
          content: '💡';
          margin-right: 6px;
        }
      }
    }
  }

  @keyframes fadeIn {
    from {
      opacity: 0;
      transform: translateY(-4px);
    }
    to {
      opacity: 1;
      transform: translateY(0);
    }
  }

  .table-toolbar {
    position: fixed;
    background: white;
    border: 1px solid #e4e7ed;
    border-radius: 6px;
    padding: 8px;
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
    z-index: 2000;
    display: flex;
    gap: 4px;
    flex-wrap: wrap;

    .el-button {
      font-size: 12px;
      padding: 4px 8px;
      margin: 0;
      
      .el-icon {
        margin-right: 2px;
      }
    }
  }
  
  .table-section {
    .table-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
      
      h4 {
        margin: 0;
        color: #303133;
        font-size: 15px;
        font-weight: 600;
        display: flex;
        align-items: center;
        
        .el-icon {
          margin-right: 6px;
          color: #1677ff;
        }
      }
      
      .table-info {
        color: #909399;
        font-size: 12px;
        
        .el-tag {
          margin-left: 8px;
        }
      }
    }
    
    .table-actions {
      display: flex;
      gap: 8px;
      margin-bottom: 12px;
      
      .el-button {
        font-size: 12px;
        padding: 5px 10px;
        
        .el-icon {
          margin-right: 4px;
        }
      }
    }
    
    .revogrid-container {
      border: 1px solid #e8eaec;
      border-radius: 8px;
      overflow: hidden;
      background: white;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
      margin: 16px 0;
      
      // RevoGrid 现代化样式 - 修复边框和表头合并显示
      :deep(revo-grid) {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif;
        font-size: 14px;
        border: 2px solid #e8eaec;
        position: relative;
        
        // 强制显示outline作为边框
        * {
          outline: 1px solid #ddd;
          outline-offset: -1px;
        }
        
        // 强制显示所有边框
        * {
          box-sizing: border-box;
        }
        
        // 表头容器
        .revo-header-viewport, 
        .revo-header {
          border-bottom: 2px solid #ddd !important;
        }
        
        // 表头单元格样式
        .revo-header-cell,
        .header-cell,
        [data-header="true"] {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
          color: white !important;
          font-weight: 600;
          text-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
          border: 1px solid rgba(255, 255, 255, 0.2) !important;
          border-right: 1px solid #ddd !important;
          border-bottom: 1px solid #ddd !important;
          text-align: center;
          padding: 8px 12px;
          
          // 确保合并表头的显示
          &.group-header {
            background: linear-gradient(135deg, #4a90e2 0%, #7b68ee 100%) !important;
            border-bottom: 2px solid #333 !important;
          }
        }
        
        // 表头输入框样式
        .header-input {
          width: 100%;
          border: none;
          background: transparent;
          color: inherit;
          font-weight: inherit;
          font-size: inherit;
          text-align: center;
          padding: 0;
          outline: none;
          
          &::placeholder {
            color: rgba(255, 255, 255, 0.7);
          }
          
          &:hover {
            background: rgba(255, 255, 255, 0.1);
          }
          
          &:focus {
            background: rgba(255, 255, 255, 0.2);
            border-radius: 2px;
          }
        }
        
        // 数据单元格样式
        .revo-data-cell,
        .data-cell,
        [data-cell="true"] {
          border: 1px solid #e8eaec !important;
          border-right: 1px solid #ddd !important;
          border-bottom: 1px solid #ddd !important;
          padding: 8px 12px;
          background: white;
          
          &:hover {
            background-color: #e3f2fd !important;
          }
          
          // 第一列样式
          &:first-child,
          &[data-col="0"] {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%) !important;
            font-weight: 600;
            text-align: center;
            color: #2c3e50;
          }
        }
        
        // 数据行样式
        .revo-data-row,
        .data-row,
        [data-row] {
          &:nth-child(even) {
            background-color: #f8f9fb;
          }
          
          &:hover {
            background-color: #f0f4ff !important;
          }
        }
        
        // 网格线强制显示
        .revo-grid-viewport {
          border-collapse: separate !important;
          border-spacing: 0 !important;
        }
        
        // 选中状态
        .revo-selected,
        .selected {
          background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%) !important;
          border: 2px solid #2196f3 !important;
          box-shadow: 0 0 0 1px rgba(33, 150, 243, 0.5);
        }
        
        // 表格主体容器边框
        .revo-viewport,
        .revo-data-viewport {
          border: 1px solid #ddd !important;
        }
        
        // 强制显示垂直分割线和所有边框
        revogr-data,
        revogr-header,
        revo-grid,
        .revo-grid {
          td, th, cell, div {
            border-right: 1px solid #ddd !important;
            border-bottom: 1px solid #ddd !important;
            border-left: 1px solid #ddd !important;
            border-top: 1px solid #ddd !important;
          }
        }
        
        // 强制显示表格内所有单元格边框
        revogr-data {
          .revo-cell,
          .cell,
          [data-cell],
          [data-col] {
            border: 1px solid #ddd !important;
            border-right: 1px solid #ddd !important;
            border-bottom: 1px solid #ddd !important;
          }
        }
        
        // 强制显示表头边框
        revogr-header {
          .revo-header-cell,
          .header-cell,
          [data-header] {
            border: 1px solid #ddd !important;
            border-right: 1px solid #ddd !important;
            border-bottom: 1px solid #ddd !important;
          }
        }
        
        // 通用边框样式强制应用
        * {
          &[data-cell],
          &[data-header],
          &[data-col],
          &[data-row] {
            border: 1px solid #ddd !important;
          }
        }
        
        // 直接使用table边框样式强制显示
        table,
        tbody,
        thead,
        tr,
        td,
        th {
          border-collapse: separate !important;
          border-spacing: 0 !important;
          border: 1px solid #ddd !important;
        }
        
        // Web Components 强制边框
        revo-grid,
        revogr-data,
        revogr-header,
        revogr-overlay-selection,
        revogr-focus,
        revogr-edit,
        revogr-viewport-scroll {
          * {
            border: 1px solid #ddd !important;
          }
        }
        
        // 确保网格线可见性
        .revo-grid-viewport,
        .revo-data-viewport,
        .revo-header-viewport {
          background-image: 
            linear-gradient(to right, #ddd 1px, transparent 1px),
            linear-gradient(to bottom, #ddd 1px, transparent 1px) !important;
          background-size: 120px 32px !important;
        }
        
        // 最强制的垂直边框显示
        revo-grid {
          // 强制表格样式
          table-layout: fixed !important;
          border-collapse: separate !important;
          border-spacing: 0 !important;
          
          // 强制所有元素显示边框
          *, 
          *::before, 
          *::after {
            border-right: 1px solid #ddd !important;
            border-bottom: 1px solid #ddd !important;
            box-sizing: border-box !important;
          }
          
          // 针对RevoGrid内部结构的强制样式
          revogr-viewport-scroll,
          revogr-data,
          revogr-header {
            border: 1px solid #ddd !important;
            
            // 内部每个cell都要有边框
            > * {
              border-right: 1px solid #ddd !important;
              border-bottom: 1px solid #ddd !important;
            }
          }
        }
        
        // 使用CSS Grid来模拟表格边框
        revo-grid::after {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          pointer-events: none;
          background-image: 
            repeating-linear-gradient(
              0deg,
              transparent,
              transparent 31px,
              #ddd 31px,
              #ddd 32px
            ),
            repeating-linear-gradient(
              90deg,
              transparent,
              transparent 119px,
              #ddd 119px,
              #ddd 120px
            );
        }
        
        // 最终解决方案：使用CSS变量和精确选择器
        &[style*="--grid-border"] {
          --revo-grid-border-color: #ddd !important;
        }
        
        // 强制显示网格线
        canvas {
          border: 1px solid #ddd !important;
        }
        
        // 最后的手段：使用box-shadow来模拟边框
        revo-grid {
          box-shadow: 
            inset 1px 0 0 #ddd,
            inset -1px 0 0 #ddd,
            inset 0 1px 0 #ddd,
            inset 0 -1px 0 #ddd !important;
        }
        
        // 终极解决方案：直接针对RevoGrid的内部元素
        revogr-viewport-scroll,
        revogr-data,
        revogr-header,
        revogr-overlay-selection {
          
          // 强制所有子元素显示边框
          > * {
            position: relative;
            
            &::before {
              content: '';
              position: absolute;
              top: 0;
              right: 0;
              bottom: 0;
              left: 0;
              border: 1px solid #ddd;
              pointer-events: none;
              z-index: 1;
            }
            
            &::after {
              content: '';
              position: absolute;
              top: 0;
              right: -1px;
              bottom: 0;
              width: 1px;
              background: #ddd;
              pointer-events: none;
              z-index: 2;
            }
          }
        }
      }
    }
    
    // 强制网格边框的全局样式
    .force-grid-borders {
      // 使用绝对定位的覆盖层来绘制网格线
      position: relative;
      
      &::after {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        pointer-events: none;
        z-index: 10;
        background-image: 
          repeating-linear-gradient(
            to right,
            transparent 0px,
            transparent 119px,
            #ddd 119px,
            #ddd 120px
          ),
          repeating-linear-gradient(
            to bottom,
            transparent 0px,
            transparent 31px,
            #ddd 31px,
            #ddd 32px
          );
      }
      
      // 强制所有内部元素显示边框
      * {
        border: 1px solid #ddd !important;
        border-collapse: separate !important;
      }
      
      // 针对RevoGrid的Web Components
      revo-grid,
      revogr-viewport-scroll,
      revogr-data,
      revogr-header {
        * {
          outline: 1px solid #ddd !important;
          outline-offset: -1px !important;
        }
      }
    }
    
    .el-alert {
      :deep(.el-alert__title) {
        font-weight: 500;
      }
    }
  }

  // 现代化表格样式
  .modern-table-container {
    border: 1px solid #e4e7ed;
    border-radius: 8px;
    overflow: hidden;
    background: white;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    margin: 16px 0;

    .modern-table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      background: white;
      font-size: 14px;

      .table-header {
        background: #f5f7fa;
        color: #303133;
        font-weight: 600;
        text-align: center;
        padding: 16px 12px;
        border-bottom: 1px solid #e4e7ed;
        border-right: 1px solid #e4e7ed;
        height: 50px;
        
        &.merged-header {
          background: #f0f2f5;
          font-weight: 700;
        }

        &.sub-header {
          background: #fafafa;
          font-size: 13px;
        }

        &:last-child {
          border-right: none;
        }
        
        // 表头输入框样式
        .header-input {
          width: 100%;
          border: none;
          background: transparent;
          color: inherit;
          font-weight: inherit;
          font-size: inherit;
          text-align: center;
          padding: 0;
          outline: none;
          
          &::placeholder {
            color: #909399;
            opacity: 0.7;
          }
          
          &:hover {
            background: rgba(64, 158, 255, 0.1);
            border-radius: 4px;
          }
          
          &:focus {
            background: rgba(64, 158, 255, 0.2);
            border-radius: 4px;
          }
        }
      }

      .table-cell {
        padding: 12px 16px;
        border-bottom: 1px solid #e4e7ed;
        border-right: 1px solid #e4e7ed;
        text-align: center;
        background: white;
        transition: background-color 0.2s ease;
        cursor: pointer;
        height: 44px;
        vertical-align: middle;

        &:hover {
          background: #f5f7fa;
        }
      }
      
      // 表头行样式（和普通数据行保持一致）
      .header-row .table-cell {
        // 移除特殊背景，使用和数据行相同的样式
        
        &:last-child {
          border-right: none;
        }
      }
      
      }
      
      .header-row .table-cell {
        .cell-input {
          width: 100%;
          border: none;
          outline: none;
          background: transparent;
          padding: 8px;
          text-align: center;
          font-size: 14px;
          height: 100%;
          line-height: 1.4;
        }

        // 单元格配置状态样式
        &.cell-configured-fixed {
          background: #e6f7ff !important;
          color: #1890ff;
          font-weight: 500;
          
          &:hover {
            background: #bae7ff !important;
          }
        }

        &.cell-configured-auto {
          background: #f0f9ff !important;
          color: #52c41a;
          font-weight: 500;
          
          &:hover {
            background: #d9f7be !important;
          }
        }

        &.cell-not-configured {
          &:hover {
            background: #f5f7fa;
          }
        }

        // 选中状态样式 - 添加动态边框效果
        &.cell-selected {
          background: linear-gradient(135deg, #e6f7ff 0%, #d4edda 100%) !important;
          border: 2px solid #52c41a !important;
          box-shadow: 0 0 0 2px rgba(82, 196, 26, 0.2);
          position: relative;
          animation: pulse-border 1.5s ease-in-out infinite;
          
          &:hover {
            background: linear-gradient(135deg, #bae7ff 0%, #c3e6cb 100%) !important;
          }
          
          // 添加选中指示器
          &::after {
            content: '';
            position: absolute;
            top: 2px;
            right: 2px;
            width: 8px;
            height: 8px;
            background: #52c41a;
            border-radius: 50%;
            box-shadow: 0 0 4px rgba(82, 196, 26, 0.5);
          }
        }

        // 合并单元格样式
        &.cell-merged {
          border: 2px solid #52c41a;
          background: #f6ffed;
          transition: all 0.3s ease;
          
          &:hover {
            background: #d9f7be !important;
            transform: scale(1.02);
            box-shadow: 0 4px 12px rgba(82, 196, 26, 0.3);
          }
        }
      }

      tbody tr {
        &:hover {
          background: #f9f9f9;
        }

        &:last-child .table-cell {
          border-bottom: none;
        }
      }
    }
  }

  // 统计表格样式
  .statistics-container {
    margin-top: 8px;
    border: 1px solid #1890ff;
    border-radius: 8px;
    overflow: hidden;
    background: white;
    box-shadow: 0 2px 8px rgba(24, 144, 255, 0.1);

    .statistics-table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      background: white;
      font-size: 14px;

      .stats-header {
        background: linear-gradient(135deg, #1890ff 0%, #40a9ff 100%);
        color: white;
        font-weight: 600;
        text-align: center;
        padding: 14px 12px;
        border: none;
        height: 46px;
      }

      .stats-cell {
        padding: 12px 16px;
        border-right: 1px solid #91d5ff;
        text-align: center;
        background: #f0f9ff;
        cursor: pointer;
        transition: background-color 0.2s ease;
        color: #1890ff;
        font-weight: 500;
        height: 44px;
        vertical-align: middle;

        &:hover {
          background: #e6f7ff;
        }

        &:last-child {
          border-right: none;
        }
      }
    }
  }

  // 合并模式激活状态
  .merge-mode-active {
    background: linear-gradient(135deg, #52c41a 0%, #73d13d 100%) !important;
    color: white !important;
    border-color: #52c41a !important;
    
    &:hover {
      background: linear-gradient(135deg, #389e0d 0%, #52c41a 100%) !important;
    }
  }

  // 字段值显示样式
  .field-value-display {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 8px 12px;
    background: #f5f7fa;
    border: 1px solid #e4e7ed;
    border-radius: 4px;
    color: #606266;

    .el-icon {
      color: #909399;
    }
  }
  
  // 字段预览表格样式
  .field-preview-table {
    .preview-note {
      margin-top: 8px;
      margin-bottom: 0;
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: #909399;
      
      .el-icon {
        color: #409eff;
      }
    }
    
    :deep(.el-table) {
      margin-top: 8px;
      
      .el-table__header {
        th {
          background-color: #f0f9ff;
          color: #1890ff;
          font-weight: 600;
        }
      }
      
      .el-table__row {
        &:nth-child(even) {
          background-color: #fafafa;
        }
      }
      
      .cell {
        font-size: 12px;
      }
    }
  }
  
  /* 单条数据预览样式 */
  .field-preview-single {
    .single-data-preview {
      background: #fafafa;
      border: 1px solid #e4e7ed;
      border-radius: 4px;
      padding: 12px;
      
      .field-item {
        display: flex;
        align-items: center;
        margin-bottom: 8px;
        
        &:last-child {
          margin-bottom: 0;
        }
        
        .field-label {
          font-weight: 500;
          color: #606266;
          min-width: 80px;
          flex-shrink: 0;
        }
        
        .field-value {
          color: #1890ff;
          background: #f0f9ff;
          padding: 2px 8px;
          border-radius: 3px;
          font-family: monospace;
          font-size: 12px;
        }
      }
    }
    
    .preview-note {
      margin-top: 8px;
      margin-bottom: 0;
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: #909399;
      
      .el-icon {
        color: #409eff;
      }
    }
  }
  
  .field-tip {
    margin-top: 8px;
    
    .el-text {
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: 12px;
      
      .el-icon {
        font-size: 14px;
      }
    }
  }
  
  /* 动画定义 */
  @keyframes pulse-border {
    0% {
      box-shadow: 0 0 0 2px rgba(82, 196, 26, 0.2);
    }
    50% {
      box-shadow: 0 0 0 6px rgba(82, 196, 26, 0.1);
    }
    100% {
      box-shadow: 0 0 0 2px rgba(82, 196, 26, 0.2);
    }
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
  
  /* 合并模式下的表格样式 */
  .merge-mode-active-table {
    td {
      cursor: crosshair !important;
      transition: all 0.2s ease;
      
      &:hover {
        background: rgba(82, 196, 26, 0.1) !important;
        border-color: #52c41a !important;
      }
    }
  }
  
  /* 拖拽选择时的预览效果 */
  .cell-selecting {
    background: linear-gradient(135deg, #e6f7ff 0%, #d4edda 100%) !important;
    border: 2px dashed #52c41a !important;
    animation: merge-highlight 2s ease-in-out infinite;
  }
</style>