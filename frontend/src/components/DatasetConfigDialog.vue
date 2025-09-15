<template>
  <el-dialog
    v-model="visible"
    title="单元格数据集配置"
    width="700px"
    @close="handleClose"
  >
    <el-form :model="form" label-width="120px">
      <el-form-item label="展示内容">
        <el-radio-group v-model="form.displayType">
          <el-radio value="text">仅文字</el-radio>
          <el-radio value="dataset">数据集</el-radio>
        </el-radio-group>
      </el-form-item>

      <template v-if="form.displayType === 'dataset'">
        <el-form-item label="数据集名称">
          <el-input v-model="form.datasetName" placeholder="请输入数据集名称" />
        </el-form-item>

        <el-form-item label="数据结构">
          <el-radio-group v-model="form.datasetType">
            <el-radio value="single">单条</el-radio>
            <el-radio value="list">列表</el-radio>
          </el-radio-group>
        </el-form-item>

        <el-form-item label="数据源">
          <el-input v-model="dataSourceDisplay" disabled />
        </el-form-item>

        <el-form-item label="SQL查询">
          <el-input
            v-model="form.sqlQuery"
            type="textarea"
            :rows="4"
            placeholder="请输入SQL查询语句"
          />
        </el-form-item>

        <el-form-item label="展示字段">
          <el-tag
            v-for="field in form.fields"
            :key="field"
            closable
            @close="removeField(field)"
            style="margin-right: 5px;"
          >
            {{ field }}
          </el-tag>
          <el-input
            v-if="inputVisible"
            ref="saveTagInput"
            v-model="inputValue"
            size="small"
            style="width: 100px;"
            @keyup.enter="handleInputConfirm"
            @blur="handleInputConfirm"
          />
          <el-button v-else size="small" @click="showInput">
            + 添加字段
          </el-button>
        </el-form-item>

        <el-form-item label="预览数据">
          <el-button @click="previewData" :loading="previewLoading">
            刷新预览
          </el-button>
          <div v-if="previewResult" class="preview-container">
            <div v-if="previewResult.type === 'single'">
              <el-descriptions :column="1" border>
                <el-descriptions-item
                  v-for="field in previewResult.fields"
                  :key="field"
                  :label="field"
                >
                  {{ previewResult.data[field] || '-' }}
                </el-descriptions-item>
              </el-descriptions>
            </div>
            <el-table
              v-else-if="previewResult.type === 'list'"
              :data="previewResult.data"
              border
              max-height="300"
            >
              <el-table-column
                v-for="field in previewResult.fields"
                :key="field"
                :prop="field"
                :label="field"
              />
            </el-table>
            <el-alert
              v-else-if="previewResult.type === 'error'"
              :title="previewResult.message"
              type="error"
              :closable="false"
            />
          </div>
        </el-form-item>
      </template>

      <template v-else>
        <el-form-item label="静态文本">
          <el-input
            v-model="form.staticText"
            type="textarea"
            :rows="4"
            placeholder="请输入静态文本内容"
          />
        </el-form-item>
      </template>
    </el-form>

    <template #footer>
      <el-button @click="handleClose">取消</el-button>
      <el-button type="primary" @click="handleConfirm" :loading="saving">
        确定
      </el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { ref, reactive, nextTick, watch } from 'vue'
import { ElMessage } from 'element-plus'

const props = defineProps({
  modelValue: Boolean,
  templateId: String,
  cellId: String
})

const emit = defineEmits(['update:modelValue', 'confirm'])

const visible = ref(props.modelValue)
const saving = ref(false)
const previewLoading = ref(false)
const inputVisible = ref(false)
const inputValue = ref('')
const previewResult = ref(null)
const dataSourceDisplay = ref('默认PostgreSQL数据源')

const form = reactive({
  displayType: 'dataset',
  datasetName: '',
  datasetType: 'list',
  dataSourceId: 'default', // 使用默认数据源
  sqlQuery: '',
  fields: [],
  staticText: ''
})

// Watch for visibility changes
watch(() => props.modelValue, (newVal) => {
  visible.value = newVal
  if (newVal) {
    loadExistingConfig()
  }
})

// loadDataSources removed - using default PostgreSQL data source

// Load existing configuration
const loadExistingConfig = async () => {
  if (!props.templateId || !props.cellId) return

  try {
    const response = await fetch(`/api/templates/${props.templateId}/cells/${props.cellId}/dataset`)
    const data = await response.json()

    if (data.success && data.config) {
      Object.assign(form, data.config)
    }
  } catch (error) {
    console.error('Failed to load existing config:', error)
  }
}

// Add field
const showInput = () => {
  inputVisible.value = true
  nextTick(() => {
    document.querySelector('.el-input__inner')?.focus()
  })
}

const handleInputConfirm = () => {
  if (inputValue.value) {
    if (!form.fields.includes(inputValue.value)) {
      form.fields.push(inputValue.value)
    }
  }
  inputVisible.value = false
  inputValue.value = ''
}

const removeField = (field) => {
  const index = form.fields.indexOf(field)
  if (index > -1) {
    form.fields.splice(index, 1)
  }
}

// Preview data
const previewData = async () => {
  if (!form.sqlQuery) {
    ElMessage.warning('请先输入SQL查询语句')
    return
  }

  previewLoading.value = true
  try {
    // Save config first
    await saveConfig()

    // Then get preview data
    const response = await fetch(`/api/templates/${props.templateId}/cells/${props.cellId}/dataset-data`)
    const data = await response.json()

    if (data.success) {
      previewResult.value = data.data
    } else {
      throw new Error(data.error || '获取预览数据失败')
    }
  } catch (error) {
    ElMessage.error('预览失败: ' + error.message)
  } finally {
    previewLoading.value = false
  }
}

// Save configuration
const saveConfig = async () => {
  const response = await fetch(`/api/templates/${props.templateId}/cells/${props.cellId}/dataset`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(form)
  })

  const data = await response.json()
  if (!data.success) {
    throw new Error(data.error || '保存配置失败')
  }

  return data
}

// Handle confirm
const handleConfirm = async () => {
  if (form.displayType === 'dataset' && !form.sqlQuery) {
    ElMessage.warning('请输入SQL查询语句')
    return
  }

  saving.value = true
  try {
    await saveConfig()

    // Insert placeholder in editor
    const placeholder = `{{dataset:${props.cellId}}}`
    emit('confirm', placeholder)

    ElMessage.success('数据集配置已保存')
    handleClose()
  } catch (error) {
    ElMessage.error('保存失败: ' + error.message)
  } finally {
    saving.value = false
  }
}

// Handle close
const handleClose = () => {
  emit('update:modelValue', false)
  // Reset form
  Object.assign(form, {
    displayType: 'dataset',
    datasetName: '',
    datasetType: 'list',
    dataSourceId: 'default', // 使用默认数据源
    sqlQuery: '',
    fields: [],
    staticText: ''
  })
  previewResult.value = null
}
</script>

<style scoped>
.preview-container {
  margin-top: 10px;
  max-height: 400px;
  overflow-y: auto;
  border: 1px solid #e4e7ed;
  padding: 10px;
  border-radius: 4px;
}
</style>