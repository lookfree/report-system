<template>
  <div class="datasources-page">
    <el-card>
      <template #header>
        <div class="card-header">
          <span>数据源管理</span>
        </div>
      </template>
      
      <el-table :data="dataSources" style="width: 100%" v-loading="loading">
        <el-table-column prop="name" label="名称" />
        <el-table-column prop="type" label="类型" />
        <el-table-column prop="host" label="主机" />
        <el-table-column prop="port" label="端口" width="80" />
        <el-table-column prop="database" label="数据库" />
        <el-table-column prop="username" label="用户名" />
        <el-table-column label="状态" width="80">
          <template #default="{ row }">
            <el-tag :type="row.active ? 'success' : 'danger'">
              {{ row.active ? '启用' : '禁用' }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="200">
          <template #default="{ row }">
            <el-button size="small" @click="testConnection(row)">
              测试连接
            </el-button>
            <el-button size="small" type="primary" @click="configureDatasets(row)">
              数据集配置
            </el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>
    
    <!-- Add/Edit Dialog -->
    <el-dialog v-model="showDialog" title="新增数据源" width="500px">
      <el-form :model="form" :rules="rules" ref="formRef" label-width="100px">
        <el-form-item label="名称" prop="name">
          <el-input v-model="form.name" placeholder="数据源名称" />
        </el-form-item>
        <el-form-item label="类型" prop="type">
          <el-select v-model="form.type" style="width: 100%">
            <el-option label="PostgreSQL" value="postgresql" />
            <el-option label="MySQL" value="mysql" />
          </el-select>
        </el-form-item>
        <el-form-item label="主机" prop="host">
          <el-input v-model="form.host" placeholder="数据库主机地址" />
        </el-form-item>
        <el-form-item label="端口" prop="port">
          <el-input-number v-model="form.port" :min="1" :max="65535" />
        </el-form-item>
        <el-form-item label="数据库" prop="database">
          <el-input v-model="form.database" placeholder="数据库名称" />
        </el-form-item>
        <el-form-item label="用户名" prop="username">
          <el-input v-model="form.username" placeholder="数据库用户名" />
        </el-form-item>
        <el-form-item label="密码" prop="password">
          <el-input v-model="form.password" type="password" placeholder="数据库密码" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showDialog = false">取消</el-button>
        <el-button @click="handleTest" :loading="testing">测试连接</el-button>
        <el-button type="primary" @click="handleSubmit" :loading="submitting">
          确定
        </el-button>
      </template>
    </el-dialog>

    <!-- 数据集配置对话框 -->
    <el-dialog v-model="showDatasetDialog" title="数据集配置" width="900px">
      <div v-loading="datasetLoading">
        <div style="margin-bottom: 20px;">
          <el-descriptions :column="2" border>
            <el-descriptions-item label="数据源名称">
              {{ currentDataSource?.name }}
            </el-descriptions-item>
            <el-descriptions-item label="数据源类型">
              {{ currentDataSource?.type }}
            </el-descriptions-item>
          </el-descriptions>
        </div>

        <div style="margin-bottom: 20px;">
          <el-button type="primary" size="small" @click="showAddDatasetDialog">
            新增数据集
          </el-button>
        </div>

        <el-table :data="currentDatasets" border>
          <el-table-column prop="name" label="数据集名称" />
          <el-table-column prop="description" label="描述" />
          <el-table-column prop="type" label="类型" width="100">
            <template #default="{ row }">
              <el-tag :type="row.type === 'single' ? 'warning' : 'primary'" size="small">
                {{ row.type === 'single' ? '单条' : '列表' }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column label="操作" width="180">
            <template #default="{ row }">
              <el-button size="small" @click="editDataset(row)">编辑</el-button>
              <el-button size="small" @click="testDataset(row)">测试</el-button>
              <el-button size="small" type="danger" @click="deleteDataset(row)">删除</el-button>
            </template>
          </el-table-column>
        </el-table>
      </div>
    </el-dialog>

    <!-- 新增/编辑数据集对话框 -->
    <el-dialog
      v-model="showDatasetFormDialog"
      :title="editingDataset ? '编辑数据集' : '新增数据集'"
      width="700px"
    >
      <el-form :model="datasetForm" label-width="100px">
        <el-form-item label="名称" required>
          <el-input v-model="datasetForm.name" placeholder="请输入数据集名称" />
        </el-form-item>

        <el-form-item label="描述">
          <el-input
            v-model="datasetForm.description"
            type="textarea"
            :rows="2"
            placeholder="请输入数据集描述"
          />
        </el-form-item>

        <el-form-item label="数据类型">
          <el-radio-group v-model="datasetForm.type">
            <el-radio value="single">单条数据</el-radio>
            <el-radio value="list">列表数据</el-radio>
          </el-radio-group>
        </el-form-item>

        <el-form-item label="SQL查询" required>
          <el-input
            v-model="datasetForm.sqlQuery"
            type="textarea"
            :rows="6"
            placeholder="请输入SQL查询语句"
            style="font-family: monospace;"
          />
        </el-form-item>

        <el-form-item label="展示字段">
          <el-tag
            v-for="field in datasetForm.fields"
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
          <el-button @click="previewDatasetData" :loading="previewLoading" size="small">
            执行查询
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
      </el-form>

      <template #footer>
        <el-button @click="showDatasetFormDialog = false">取消</el-button>
        <el-button type="primary" @click="saveDataset" :loading="savingDataset">
          保存
        </el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, nextTick } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import api from '@/api'

const dataSources = ref([])
const loading = ref(false)
const showDialog = ref(false)
const testing = ref(false)
const submitting = ref(false)
const formRef = ref()

// 数据集相关
const showDatasetDialog = ref(false)
const showDatasetFormDialog = ref(false)
const currentDataSource = ref(null)
const currentDatasets = ref([])
const datasetLoading = ref(false)
const editingDataset = ref(null)
const savingDataset = ref(false)
const previewLoading = ref(false)
const previewResult = ref(null)
const inputVisible = ref(false)
const inputValue = ref('')

const datasetForm = reactive({
  name: '',
  description: '',
  type: 'list',
  sqlQuery: '',
  fields: [],
  dataSourceId: null
})

const form = reactive({
  name: '',
  type: 'postgresql',
  host: '',
  port: 5432,
  database: '',
  username: '',
  password: ''
})

const rules = {
  name: [{ required: true, message: '请输入名称', trigger: 'blur' }],
  type: [{ required: true, message: '请选择类型', trigger: 'change' }],
  host: [{ required: true, message: '请输入主机地址', trigger: 'blur' }],
  port: [{ required: true, message: '请输入端口', trigger: 'blur' }],
  database: [{ required: true, message: '请输入数据库名称', trigger: 'blur' }],
  username: [{ required: true, message: '请输入用户名', trigger: 'blur' }],
  password: [{ required: true, message: '请输入密码', trigger: 'blur' }]
}

const loadDataSources = async () => {
  loading.value = true
  try {
    dataSources.value = await api.getDataSources()
  } catch (error) {
    ElMessage.error('加载数据源失败')
  }
  loading.value = false
}

const testConnection = async (dataSource) => {
  try {
    await api.testDataSource(dataSource)
    ElMessage.success('连接成功')
  } catch (error) {
    ElMessage.error('连接失败: ' + (error.response?.data?.error || '未知错误'))
  }
}

const handleTest = async () => {
  await formRef.value.validate()
  testing.value = true
  try {
    await api.testDataSource(form)
    ElMessage.success('连接成功')
  } catch (error) {
    ElMessage.error('连接失败: ' + (error.response?.data?.error || '未知错误'))
  }
  testing.value = false
}

const handleSubmit = async () => {
  await formRef.value.validate()
  submitting.value = true
  try {
    await api.createDataSource(form)
    ElMessage.success('数据源创建成功')
    showDialog.value = false
    resetForm()
    loadDataSources()
  } catch (error) {
    ElMessage.error('创建失败')
  }
  submitting.value = false
}

const resetForm = () => {
  Object.assign(form, {
    name: '',
    type: 'postgresql',
    host: '',
    port: 5432,
    database: '',
    username: '',
    password: ''
  })
}

// 数据集配置功能
const configureDatasets = (dataSource) => {
  currentDataSource.value = dataSource
  loadDatasets(dataSource.id)
  showDatasetDialog.value = true
}

const loadDatasets = async (dataSourceId) => {
  datasetLoading.value = true
  try {
    const response = await fetch(`/api/datasources/${dataSourceId}/datasets`)
    const data = await response.json()
    if (data.success) {
      currentDatasets.value = data.datasets || []
    }
  } catch (error) {
    console.error('Failed to load datasets:', error)
    currentDatasets.value = []
  } finally {
    datasetLoading.value = false
  }
}

const showAddDatasetDialog = () => {
  editingDataset.value = null
  Object.assign(datasetForm, {
    name: '',
    description: '',
    type: 'list',
    sqlQuery: '',
    fields: [],
    dataSourceId: currentDataSource.value.id
  })
  previewResult.value = null
  showDatasetFormDialog.value = true
}

const editDataset = (dataset) => {
  editingDataset.value = dataset
  Object.assign(datasetForm, {
    name: dataset.name,
    description: dataset.description || '',
    type: dataset.type,
    sqlQuery: dataset.sqlQuery,
    fields: [...dataset.fields],
    dataSourceId: currentDataSource.value.id
  })
  previewResult.value = null
  showDatasetFormDialog.value = true
}

const testDataset = async (dataset) => {
  try {
    const response = await fetch(`/api/datasets/${dataset.id}/execute`)
    const data = await response.json()
    if (data.success && data.result) {
      ElMessage.success('测试成功，数据查询正常')
    } else {
      throw new Error(data.error || '查询失败')
    }
  } catch (error) {
    ElMessage.error('测试失败: ' + error.message)
  }
}

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
      loadDatasets(currentDataSource.value.id)
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
    if (!datasetForm.fields.includes(inputValue.value)) {
      datasetForm.fields.push(inputValue.value)
    }
  }
  inputVisible.value = false
  inputValue.value = ''
}

const removeField = (field) => {
  const index = datasetForm.fields.indexOf(field)
  if (index > -1) {
    datasetForm.fields.splice(index, 1)
  }
}

// 预览数据
const previewDatasetData = async () => {
  if (!datasetForm.sqlQuery) {
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
        sqlQuery: datasetForm.sqlQuery,
        type: datasetForm.type,
        fields: datasetForm.fields
      })
    })

    const data = await response.json()
    if (data.success) {
      previewResult.value = data.result

      // 自动提取字段名
      if (datasetForm.fields.length === 0 && data.result.data) {
        const sampleData = datasetForm.type === 'single' ? data.result.data : data.result.data[0]
        if (sampleData) {
          datasetForm.fields = Object.keys(sampleData)
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
  if (!datasetForm.name) {
    ElMessage.warning('请输入数据集名称')
    return
  }
  if (!datasetForm.sqlQuery) {
    ElMessage.warning('请输入SQL查询语句')
    return
  }
  if (datasetForm.fields.length === 0) {
    ElMessage.warning('请添加展示字段')
    return
  }

  savingDataset.value = true
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
      body: JSON.stringify({
        ...datasetForm,
        dataSourceId: currentDataSource.value.id
      })
    })

    const data = await response.json()
    if (data.success) {
      ElMessage.success(editingDataset.value ? '更新成功' : '创建成功')
      showDatasetFormDialog.value = false
      loadDatasets(currentDataSource.value.id)
    } else {
      throw new Error(data.error || '保存失败')
    }
  } catch (error) {
    ElMessage.error('保存失败: ' + error.message)
  } finally {
    savingDataset.value = false
  }
}

onMounted(() => {
  loadDataSources()
})
</script>

<style scoped lang="scss">
.datasources-page {
  .card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
  }

  .preview-container {
    margin-top: 10px;
    max-height: 300px;
    overflow-y: auto;
    border: 1px solid #e4e7ed;
    padding: 10px;
    border-radius: 4px;
    background-color: #f5f7fa;
  }
}
</style>