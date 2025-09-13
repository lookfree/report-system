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
        <el-table-column label="操作" width="150">
          <template #default="{ row }">
            <el-button size="small" @click="testConnection(row)">
              测试连接
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
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import api from '@/api'

const dataSources = ref([])
const loading = ref(false)
const showDialog = ref(false)
const testing = ref(false)
const submitting = ref(false)
const formRef = ref()

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
}
</style>