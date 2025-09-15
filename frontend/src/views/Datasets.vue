<template>
  <div class="datasets-page">
    <div class="page-header">
      <h2>数据集管理</h2>
      <el-button type="primary" @click="showAddDialog">
        <el-icon><Plus /></el-icon>
        新建数据集
      </el-button>
    </div>

    <el-table :data="datasets" style="width: 100%" v-loading="loading">
      <el-table-column prop="name" label="数据集名称" width="200" />
      <el-table-column prop="description" label="描述" />
      <el-table-column prop="type" label="类型" width="100">
        <template #default="{ row }">
          <el-tag :type="row.type === 'single' ? 'warning' : 'primary'">
            {{ row.type === 'single' ? '单条' : '列表' }}
          </el-tag>
        </template>
      </el-table-column>
      <el-table-column prop="fields" label="字段" width="200">
        <template #default="{ row }">
          <el-tag
            v-for="field in row.fields"
            :key="field"
            size="small"
            style="margin: 2px;"
          >
            {{ field }}
          </el-tag>
        </template>
      </el-table-column>
      <el-table-column prop="createdAt" label="创建时间" width="180">
        <template #default="{ row }">
          {{ formatDate(row.createdAt) }}
        </template>
      </el-table-column>
      <el-table-column label="操作" width="200" fixed="right">
        <template #default="{ row }">
          <el-button size="small" @click="editDataset(row)">编辑</el-button>
          <el-button size="small" @click="testDataset(row)">测试</el-button>
          <el-button size="small" type="danger" @click="deleteDataset(row)">删除</el-button>
        </template>
      </el-table-column>
    </el-table>

    <!-- 新建/编辑数据集对话框 -->
    <el-dialog
      v-model="dialogVisible"
      :title="editingDataset ? '编辑数据集' : '新建数据集'"
      width="800px"
    >
      <el-form :model="form" label-width="100px">
        <el-form-item label="名称" required>
          <el-input v-model="form.name" placeholder="请输入数据集名称" />
        </el-form-item>

        <el-form-item label="描述">
          <el-input
            v-model="form.description"
            type="textarea"
            :rows="2"
            placeholder="请输入数据集描述"
          />
        </el-form-item>

        <el-form-item label="数据类型">
          <el-radio-group v-model="form.type">
            <el-radio value="single">单条数据</el-radio>
            <el-radio value="list">列表数据</el-radio>
          </el-radio-group>
        </el-form-item>

        <el-form-item label="SQL查询" required>
          <el-input
            v-model="form.sqlQuery"
            type="textarea"
            :rows="6"
            placeholder="请输入SQL查询语句"
            style="font-family: monospace;"
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
            执行查询
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
      </el-form>

      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="saveDataset" :loading="saving">
          保存
        </el-button>
      </template>
    </el-dialog>

    <!-- 测试数据集对话框 -->
    <el-dialog
      v-model="testDialogVisible"
      title="测试数据集"
      width="800px"
    >
      <div v-loading="testLoading">
        <el-descriptions :column="2" border style="margin-bottom: 20px;">
          <el-descriptions-item label="数据集名称">
            {{ testingDataset?.name }}
          </el-descriptions-item>
          <el-descriptions-item label="类型">
            {{ testingDataset?.type === 'single' ? '单条' : '列表' }}
          </el-descriptions-item>
        </el-descriptions>

        <div v-if="testResult">
          <h4>查询结果：</h4>
          <div v-if="testResult.type === 'single'">
            <el-descriptions :column="1" border>
              <el-descriptions-item
                v-for="field in testResult.fields"
                :key="field"
                :label="field"
              >
                {{ testResult.data[field] || '-' }}
              </el-descriptions-item>
            </el-descriptions>
          </div>
          <el-table
            v-else-if="testResult.type === 'list'"
            :data="testResult.data"
            border
            max-height="400"
          >
            <el-table-column
              v-for="field in testResult.fields"
              :key="field"
              :prop="field"
              :label="field"
            />
          </el-table>
          <el-alert
            v-else-if="testResult.type === 'error'"
            :title="testResult.message"
            type="error"
            :closable="false"
          />
        </div>
      </div>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, nextTick } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Plus } from '@element-plus/icons-vue'

// 数据
const datasets = ref([])
const loading = ref(false)
const dialogVisible = ref(false)
const testDialogVisible = ref(false)
const editingDataset = ref(null)
const testingDataset = ref(null)
const saving = ref(false)
const previewLoading = ref(false)
const testLoading = ref(false)
const previewResult = ref(null)
const testResult = ref(null)
const inputVisible = ref(false)
const inputValue = ref('')

const form = reactive({
  name: '',
  description: '',
  type: 'list',
  sqlQuery: '',
  fields: []
})

// 加载数据集列表
const loadDatasets = async () => {
  loading.value = true
  try {
    const response = await fetch('/api/datasets')
    const data = await response.json()
    if (data.success) {
      datasets.value = data.datasets || []
    }
  } catch (error) {
    console.error('Failed to load datasets:', error)
    ElMessage.error('加载数据集失败')
  } finally {
    loading.value = false
  }
}

// 显示新建对话框
const showAddDialog = () => {
  editingDataset.value = null
  Object.assign(form, {
    name: '',
    description: '',
    type: 'list',
    sqlQuery: '',
    fields: []
  })
  previewResult.value = null
  dialogVisible.value = true
}

// 编辑数据集
const editDataset = (dataset) => {
  editingDataset.value = dataset
  Object.assign(form, {
    name: dataset.name,
    description: dataset.description || '',
    type: dataset.type,
    sqlQuery: dataset.sqlQuery,
    fields: [...dataset.fields]
  })
  previewResult.value = null
  dialogVisible.value = true
}

// 测试数据集
const testDataset = async (dataset) => {
  testingDataset.value = dataset
  testResult.value = null
  testDialogVisible.value = true

  testLoading.value = true
  try {
    const response = await fetch(`/api/datasets/${dataset.id}/execute`)
    const data = await response.json()
    if (data.success) {
      testResult.value = data.result
    } else {
      throw new Error(data.error || '执行失败')
    }
  } catch (error) {
    ElMessage.error('测试失败: ' + error.message)
  } finally {
    testLoading.value = false
  }
}

// 删除数据集
const deleteDataset = async (dataset) => {
  try {
    await ElMessageBox.confirm(
      `确定要删除数据集 "${dataset.name}" 吗？`,
      '确认删除',
      {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }
    )

    const response = await fetch(`/api/datasets/${dataset.id}`, {
      method: 'DELETE'
    })
    const data = await response.json()

    if (data.success) {
      ElMessage.success('删除成功')
      loadDatasets()
    } else {
      throw new Error(data.error || '删除失败')
    }
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败: ' + error.message)
    }
  }
}

// 字段管理
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

// 预览数据
const previewData = async () => {
  if (!form.sqlQuery) {
    ElMessage.warning('请输入SQL查询语句')
    return
  }

  previewLoading.value = true
  try {
    const response = await fetch('/api/datasets/preview', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        sqlQuery: form.sqlQuery,
        type: form.type,
        fields: form.fields
      })
    })

    const data = await response.json()
    if (data.success) {
      previewResult.value = data.result

      // 自动提取字段名
      if (form.fields.length === 0 && data.result.data) {
        const sampleData = form.type === 'single' ? data.result.data : data.result.data[0]
        if (sampleData) {
          form.fields = Object.keys(sampleData)
        }
      }
    } else {
      throw new Error(data.error || '查询失败')
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

// 保存数据集
const saveDataset = async () => {
  if (!form.name) {
    ElMessage.warning('请输入数据集名称')
    return
  }
  if (!form.sqlQuery) {
    ElMessage.warning('请输入SQL查询语句')
    return
  }
  if (form.fields.length === 0) {
    ElMessage.warning('请添加展示字段')
    return
  }

  saving.value = true
  try {
    const url = editingDataset.value
      ? `/api/datasets/${editingDataset.value.id}`
      : '/api/datasets'

    const method = editingDataset.value ? 'PUT' : 'POST'

    const response = await fetch(url, {
      method,
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(form)
    })

    const data = await response.json()
    if (data.success) {
      ElMessage.success(editingDataset.value ? '更新成功' : '创建成功')
      dialogVisible.value = false
      loadDatasets()
    } else {
      throw new Error(data.error || '保存失败')
    }
  } catch (error) {
    ElMessage.error('保存失败: ' + error.message)
  } finally {
    saving.value = false
  }
}

// 格式化日期
const formatDate = (date) => {
  if (!date) return '-'
  return new Date(date).toLocaleString('zh-CN')
}

// 初始化
onMounted(() => {
  loadDatasets()
})
</script>

<style scoped>
.datasets-page {
  padding: 20px;
}

.page-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.page-header h2 {
  margin: 0;
}

.preview-container {
  margin-top: 10px;
  max-height: 400px;
  overflow-y: auto;
  border: 1px solid #e4e7ed;
  padding: 10px;
  border-radius: 4px;
}
</style>