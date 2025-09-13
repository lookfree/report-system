<template>
  <!-- 数据插入配置对话框 -->
  <el-dialog 
    :model-value="visible" 
    @update:model-value="$emit('update:visible', $event)"
    title="插入数据" 
    width="800px"
  >
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
      <el-button @click="$emit('update:visible', false)">取消</el-button>
      <el-button type="primary" @click="$emit('confirm', fieldForm)">确定</el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { reactive, watch } from 'vue'

const props = defineProps({
  visible: Boolean,
  dataSources: {
    type: Array,
    default: () => []
  },
  datasetFields: {
    type: Array, 
    default: () => []
  },
  previewData: {
    type: Array,
    default: () => []
  }
})

const emit = defineEmits([
  'update:visible',
  'confirm',
  'data-source-change'
])

const fieldForm = reactive({
  insertType: 'FIELD',
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

const onDataSourceChange = () => {
  emit('data-source-change', fieldForm.dataSourceId)
}

// 监听对话框可见性变化，重置表单
watch(() => props.visible, (newValue) => {
  if (newValue) {
    // 重置表单
    Object.assign(fieldForm, {
      insertType: 'FIELD',
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
  }
})
</script>

<style scoped>
/* 对话框内部样式如果需要 */
</style>