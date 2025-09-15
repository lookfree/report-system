<template>
  <el-dialog
    v-model="visible"
    title="插入数据集"
    width="600px"
    @close="handleClose"
  >
    <div v-loading="loading">
      <el-form :model="form" label-width="100px">
        <el-form-item label="选择数据集">
          <el-select
            v-model="form.datasetId"
            placeholder="请选择数据集"
            @change="onDatasetSelect"
            style="width: 100%;"
          >
            <el-option
              v-for="dataset in datasets"
              :key="dataset.id"
              :label="dataset.name"
              :value="dataset.id"
            >
              <div>
                <div>{{ dataset.name }}</div>
                <div style="font-size: 12px; color: #999;">
                  {{ dataset.description }}
                </div>
              </div>
            </el-option>
          </el-select>
        </el-form-item>

        <template v-if="selectedDataset">
          <el-form-item label="数据类型">
            <el-tag :type="selectedDataset.type === 'single' ? 'warning' : 'primary'">
              {{ selectedDataset.type === 'single' ? '单条数据' : '列表数据' }}
            </el-tag>
          </el-form-item>

          <el-form-item label="展示字段">
            <el-tag
              v-for="field in selectedDataset.fields"
              :key="field"
              size="small"
              style="margin: 2px;"
            >
              {{ field }}
            </el-tag>
          </el-form-item>

          <el-form-item label="预览数据">
            <el-button @click="previewData" :loading="previewLoading" size="small">
              刷新预览
            </el-button>
            <div v-if="previewResult" class="preview-container">
              <div v-if="previewResult.type === 'single'">
                <el-descriptions :column="1" border size="small">
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
                size="small"
                max-height="200"
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
                size="small"
              />
            </div>
          </el-form-item>
        </template>
      </el-form>
    </div>

    <template #footer>
      <el-button @click="handleClose">取消</el-button>
      <el-button
        type="primary"
        @click="handleConfirm"
        :disabled="!selectedDataset"
      >
        插入
      </el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { ref, reactive, watch, computed } from 'vue'
import { ElMessage } from 'element-plus'

const props = defineProps({
  modelValue: Boolean
})

const emit = defineEmits(['update:modelValue', 'confirm'])

const visible = ref(props.modelValue)
const loading = ref(false)
const previewLoading = ref(false)
const datasets = ref([])
const previewResult = ref(null)

const form = reactive({
  datasetId: null
})

const selectedDataset = computed(() => {
  return datasets.value.find(d => d.id === form.datasetId)
})

// Watch for visibility changes
watch(() => props.modelValue, (newVal) => {
  visible.value = newVal
  if (newVal) {
    loadDatasets()
  }
})

// 加载数据集列表
const loadDatasets = async () => {
  loading.value = true
  try {
    const response = await fetch('/api/datasets')
    const data = await response.json()
    if (data.success) {
      datasets.value = data.datasets || []
      // 如果只有一个数据集，自动选中
      if (datasets.value.length === 1) {
        form.datasetId = datasets.value[0].id
        onDatasetSelect()
      }
    }
  } catch (error) {
    console.error('Failed to load datasets:', error)
    ElMessage.error('加载数据集失败')
  } finally {
    loading.value = false
  }
}

// 选择数据集
const onDatasetSelect = () => {
  previewResult.value = null
  if (form.datasetId) {
    previewData()
  }
}

// 预览数据
const previewData = async () => {
  if (!form.datasetId) return

  previewLoading.value = true
  try {
    const response = await fetch(`/api/datasets/${form.datasetId}/execute`)
    const data = await response.json()

    if (data.success) {
      previewResult.value = data.result
    } else {
      throw new Error(data.error || '获取数据失败')
    }
  } catch (error) {
    ElMessage.error('预览失败: ' + error.message)
    previewResult.value = {
      type: 'error',
      message: error.message
    }
  } finally {
    previewLoading.value = false
  }
}

// 确认插入
const handleConfirm = () => {
  if (!selectedDataset.value) {
    ElMessage.warning('请选择数据集')
    return
  }

  // 生成占位符
  const placeholder = `{{dataset:${selectedDataset.value.id}:${selectedDataset.value.name}}}`

  emit('confirm', {
    placeholder,
    dataset: selectedDataset.value
  })

  handleClose()
  ElMessage.success('数据集已插入')
}

// 关闭对话框
const handleClose = () => {
  emit('update:modelValue', false)
  form.datasetId = null
  previewResult.value = null
}
</script>

<style scoped>
.preview-container {
  margin-top: 10px;
  max-height: 300px;
  overflow-y: auto;
  border: 1px solid #e4e7ed;
  padding: 10px;
  border-radius: 4px;
  background-color: #f5f7fa;
}
</style>