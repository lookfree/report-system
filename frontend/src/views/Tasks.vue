<template>
  <div class="report-tasks-page">
    <!-- 新建报告任务卡片 -->
    <el-card class="create-task-card">
      <div class="task-form">
        <div class="form-row">
          <div class="form-group">
            <label class="required-label">报告名称</label>
            <el-input 
              v-model="form.name" 
              placeholder="请输入报告名称"
              class="form-input"
            />
          </div>
          <div class="form-group">
            <label class="required-label">所属组织</label>
            <el-select 
              v-model="form.organizationId" 
              placeholder="请选择"
              class="form-input"
            >
              <el-option label="组织1" value="org1" />
              <el-option label="组织2" value="org2" />
            </el-select>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group task-type-group">
            <label class="required-label">任务类型</label>
            <div class="task-type-buttons">
              <el-button 
                :type="form.taskType === 'daily' ? 'primary' : ''"
                @click="form.taskType = 'daily'"
                class="task-type-btn"
              >
                每日
              </el-button>
              <el-button 
                :type="form.taskType === 'weekly' ? 'primary' : ''"
                @click="form.taskType = 'weekly'"
                class="task-type-btn"
              >
                每周
              </el-button>
              <el-button 
                :type="form.taskType === 'monthly' ? 'primary' : ''"
                @click="form.taskType = 'monthly'"
                class="task-type-btn"
              >
                每月
              </el-button>
            </div>
          </div>
          <div class="form-group">
            <label class="required-label">审计周期</label>
            <el-time-picker
              v-model="form.auditPeriod"
              placeholder="请选择时间"
              format="HH:mm"
              value-format="HH:mm"
              class="form-input"
            />
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group template-group">
            <label>模板名称</label>
            <el-select 
              v-model="form.templateId" 
              placeholder="请选择模板"
              class="form-input"
            >
              <el-option
                v-for="template in templates"
                :key="template.id"
                :label="template.name"
                :value="template.id"
              />
            </el-select>
          </div>
          <div class="form-group status-group">
            <label>状态</label>
            <el-switch 
              v-model="form.active" 
              size="large"
              class="status-switch"
            />
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group description-group">
            <label>报告描述</label>
            <el-input 
              v-model="form.description" 
              type="textarea"
              :rows="3"
              placeholder="请输入报告描述"
              class="form-textarea"
            />
          </div>
        </div>
        
        <div class="form-row attachment-row">
          <div class="form-group attachment-group">
            <label>事件报告附件</label>
            <div class="attachment-container">
              <el-upload
                ref="uploadRef"
                class="upload-demo"
                :action="uploadAction"
                :on-preview="handlePreview"
                :on-remove="handleRemove"
                :before-remove="beforeRemove"
                :on-success="handleSuccess"
                :on-error="handleError"
                :file-list="fileList"
                :limit="5"
                :on-exceed="handleExceed"
                accept=".pdf,.doc,.docx,.xls,.xlsx,.txt,.jpg,.jpeg,.png"
                multiple
              >
                <el-button type="primary" class="upload-btn">
                  <el-icon><UploadFilled /></el-icon>
                  上传附件
                </el-button>
                <template #tip>
                  <div class="el-upload__tip">
                    支持 PDF, Word, Excel, 图片等格式，最多上传5个文件，单个文件不超过10MB
                  </div>
                </template>
              </el-upload>
            </div>
          </div>
        </div>
        
        <div class="form-row action-row">
          <div class="form-group">
            <el-button type="primary" @click="createTask" :loading="submitting" class="create-task-btn">
              创建报告任务
            </el-button>
            <el-button @click="resetForm" class="reset-btn">
              重置
            </el-button>
          </div>
        </div>
      </div>
    </el-card>
    
    <!-- 现有任务列表 -->
    <el-card class="tasks-list-card" v-if="tasks.length > 0">
      <template #header>
        <div class="card-header">
          <span>报告任务列表</span>
        </div>
      </template>
      
      <el-table :data="tasks" style="width: 100%" v-loading="loading">
        <el-table-column prop="name" label="任务名称" />
        <el-table-column label="模板">
          <template #default="{ row }">
            {{ row.template?.name }}
          </template>
        </el-table-column>
        <el-table-column prop="taskType" label="任务类型">
          <template #default="{ row }">
            <el-tag>{{ getTaskTypeLabel(row.taskType) }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="状态" width="80">
          <template #default="{ row }">
            <el-switch
              v-model="row.active"
              @change="toggleTask(row)"
              :loading="row.loading"
            />
          </template>
        </el-table-column>
        <el-table-column label="上次执行">
          <template #default="{ row }">
            {{ row.lastRunTime ? formatDate(row.lastRunTime) : '未执行' }}
          </template>
        </el-table-column>
        <el-table-column label="下次执行">
          <template #default="{ row }">
            {{ row.nextRunTime ? formatDate(row.nextRunTime) : '-' }}
          </template>
        </el-table-column>
        <el-table-column label="操作" width="120">
          <template #default="{ row }">
            <el-button type="primary" size="small" @click="editTask(row)">编辑</el-button>
            <el-button type="danger" size="small" @click="deleteTask(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>
    
    <!-- 编辑任务对话框 -->
    <el-dialog v-model="showEditDialog" title="编辑报告任务" width="500px">
      <el-form :model="editForm" :rules="rules" ref="editFormRef" label-width="100px">
        <el-form-item label="任务名称" prop="name">
          <el-input v-model="editForm.name" placeholder="任务名称" />
        </el-form-item>
        <el-form-item label="选择模板" prop="templateId">
          <el-select v-model="editForm.templateId" placeholder="请选择模板" style="width: 100%">
            <el-option
              v-for="template in templates"
              :key="template.id"
              :label="template.name"
              :value="template.id"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="任务类型" prop="taskType">
          <el-select v-model="editForm.taskType" placeholder="请选择任务类型" style="width: 100%">
            <el-option label="每日" value="daily" />
            <el-option label="每周" value="weekly" />
            <el-option label="每月" value="monthly" />
          </el-select>
        </el-form-item>
        <el-form-item label="状态">
          <el-switch v-model="editForm.active" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showEditDialog = false">取消</el-button>
        <el-button type="primary" @click="handleUpdate" :loading="submitting">
          确定
        </el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { UploadFilled } from '@element-plus/icons-vue'
import api from '@/api'
import dayjs from 'dayjs'

const tasks = ref([])
const templates = ref([])
const loading = ref(false)
const showEditDialog = ref(false)
const submitting = ref(false)
const editFormRef = ref()
const editingTask = ref(null)
const uploadRef = ref()
const fileList = ref([])
const uploadAction = ref('/api/upload/attachment')

const form = reactive({
  name: '',
  organizationId: '',
  taskType: 'daily',
  auditPeriod: '',
  templateId: '',
  description: '',
  active: true
})

const editForm = reactive({
  name: '',
  templateId: '',
  taskType: 'daily',
  active: true
})

const rules = {
  name: [{ required: true, message: '请输入任务名称', trigger: 'blur' }],
  templateId: [{ required: true, message: '请选择模板', trigger: 'change' }],
  taskType: [{ required: true, message: '请选择任务类型', trigger: 'change' }]
}

const formatDate = (date) => {
  return dayjs(date).format('YYYY-MM-DD HH:mm')
}

const getTaskTypeLabel = (type) => {
  const labels = {
    daily: '每日',
    weekly: '每周',
    monthly: '每月'
  }
  return labels[type] || type
}

const loadTasks = async () => {
  loading.value = true
  try {
    const data = await api.getTasks()
    tasks.value = data.map(t => ({ ...t, loading: false }))
  } catch (error) {
    ElMessage.error('加载任务失败')
  }
  loading.value = false
}

const loadTemplates = async () => {
  try {
    templates.value = await api.getTemplates()
  } catch (error) {
    ElMessage.error('加载模板失败')
  }
}

const toggleTask = async (task) => {
  task.loading = true
  try {
    await api.toggleTask(task.id)
    ElMessage.success(task.active ? '任务已启用' : '任务已停用')
    loadTasks()
  } catch (error) {
    ElMessage.error('操作失败')
    task.active = !task.active
  }
  task.loading = false
}

const editTask = (task) => {
  editingTask.value = task
  Object.assign(editForm, {
    name: task.name,
    templateId: task.templateId,
    taskType: task.taskType || 'daily',
    active: task.active
  })
  showEditDialog.value = true
}

const deleteTask = async (task) => {
  try {
    await ElMessageBox.confirm('确定要删除这个任务吗？', '提示', {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'warning'
    })
    await api.deleteTask(task.id)
    ElMessage.success('删除成功')
    loadTasks()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败')
    }
  }
}

const handleUpdate = async () => {
  await editFormRef.value.validate()
  submitting.value = true
  try {
    await api.updateTask(editingTask.value.id, editForm)
    ElMessage.success('更新成功')
    showEditDialog.value = false
    loadTasks()
  } catch (error) {
    ElMessage.error('更新失败')
  }
  submitting.value = false
}

const createTask = async () => {
  if (!form.name.trim()) {
    ElMessage.warning('请输入报告名称')
    return
  }
  
  submitting.value = true
  try {
    const taskData = {
      ...form,
      cronExpression: getCronExpression(form.taskType)
    }
    await api.createTask(taskData)
    ElMessage.success('任务创建成功')
    resetForm()
    loadTasks()
  } catch (error) {
    ElMessage.error('创建失败')
  }
  submitting.value = false
}

const getCronExpression = (taskType) => {
  const expressions = {
    daily: '0 0 * * *',
    weekly: '0 0 * * 1',
    monthly: '0 0 1 * *'
  }
  return expressions[taskType] || '0 0 * * *'
}

const resetForm = () => {
  Object.assign(form, {
    name: '',
    organizationId: '',
    taskType: 'daily',
    auditPeriod: '',
    templateId: '',
    description: '',
    active: true
  })
  // 重置文件列表
  fileList.value = []
}

// 文件上传相关方法
const handleSuccess = (response, file, fileList) => {
  ElMessage.success(`${file.name} 上传成功`)
}

const handleError = (error, file, fileList) => {
  ElMessage.error(`${file.name} 上传失败`)
}

const handlePreview = (file) => {
  if (file.url) {
    window.open(file.url, '_blank')
  } else if (file.response && file.response.url) {
    window.open(file.response.url, '_blank')
  }
}

const handleRemove = (file, fileList) => {
  ElMessage.info(`${file.name} 已移除`)
}

const beforeRemove = (file) => {
  return ElMessageBox.confirm(
    `确定移除 ${file.name}？`
  ).then(
    () => true,
    () => false
  )
}

const handleExceed = (files, fileList) => {
  ElMessage.warning(`当前限制选择 5 个文件，本次选择了 ${files.length} 个文件，共选择了 ${files.length + fileList.length} 个文件`)
}

onMounted(() => {
  loadTasks()
  loadTemplates()
})
</script>

<style scoped lang="scss">
.report-tasks-page {
  .create-task-card {
    margin-bottom: 24px;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
    
    .task-form {
      padding: 24px;
      
      .form-row {
        display: flex;
        gap: 24px;
        margin-bottom: 24px;
        
        .form-group {
          flex: 1;
          
          label {
            display: block;
            margin-bottom: 8px;
            font-size: 14px;
            color: #303133;
            font-weight: 500;
            
            &.required-label::before {
              content: '*';
              color: #f56c6c;
              margin-right: 4px;
            }
          }
          
          .form-input {
            width: 100%;
          }
          
          .form-textarea {
            width: 100%;
          }
          
          &.task-type-group {
            .task-type-buttons {
              display: flex;
              gap: 8px;
              
              .task-type-btn {
                flex: 1;
                border-radius: 6px;
                font-weight: 500;
                
                &.el-button--primary {
                  background: #1890ff;
                  border-color: #1890ff;
                }
              }
            }
          }
          
          &.template-group {
            flex: 2;
          }
          
          &.status-group {
            display: flex;
            flex-direction: column;
            align-items: flex-start;
            
            .status-switch {
              margin-top: auto;
            }
          }
          
          &.description-group {
            width: 100%;
          }
        }
        
        &.attachment-row {
          .attachment-group {
            width: 100%;
            
            .attachment-container {
              .upload-demo {
                width: 100%;
                
                .upload-btn {
                  background: #1890ff;
                  border-color: #1890ff;
                  border-radius: 6px;
                  padding: 8px 24px;
                  font-weight: 500;
                  
                  .el-icon {
                    margin-right: 6px;
                  }
                }
              }
            }
          }
        }
        
        &.action-row {
          margin-top: 32px;
          padding-top: 24px;
          border-top: 1px solid #f0f0f0;
          
          .form-group {
            display: flex;
            gap: 16px;
            justify-content: flex-end;
            
            .create-task-btn {
              background: #1890ff;
              border-color: #1890ff;
              border-radius: 6px;
              padding: 10px 32px;
              font-weight: 600;
              font-size: 14px;
            }
            
            .reset-btn {
              border-radius: 6px;
              padding: 10px 24px;
              font-weight: 500;
            }
          }
        }
      }
    }
  }
  
  .tasks-list-card {
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
    
    .card-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-weight: 600;
      color: #303133;
    }
  }
  
  // 响应式设计
  @media (max-width: 768px) {
    .task-form {
      .form-row {
        flex-direction: column;
        gap: 16px;
        
        .form-group {
          &.task-type-group {
            .task-type-buttons {
              .task-type-btn {
                flex: none;
                min-width: 80px;
              }
            }
          }
        }
      }
    }
  }
}

// Element Plus 样式覆盖
:deep(.el-card__header) {
  padding: 16px 24px;
  border-bottom: 1px solid #f0f0f0;
}

:deep(.el-card__body) {
  padding: 0;
}

:deep(.el-input__wrapper) {
  border-radius: 6px;
  box-shadow: 0 0 0 1px #d9d9d9;
  
  &:hover {
    box-shadow: 0 0 0 1px #40a9ff;
  }
  
  &.is-focus {
    box-shadow: 0 0 0 2px rgba(24, 144, 255, 0.2);
  }
}

:deep(.el-select .el-input__wrapper) {
  border-radius: 6px;
}

:deep(.el-button) {
  border-radius: 6px;
  font-weight: 500;
  
  &.el-button--primary {
    background: #1890ff;
    border-color: #1890ff;
    
    &:hover {
      background: #40a9ff;
      border-color: #40a9ff;
    }
  }
  
  &.el-button--danger {
    background: #ff4d4f;
    border-color: #ff4d4f;
    
    &:hover {
      background: #ff7875;
      border-color: #ff7875;
    }
  }
}

:deep(.el-switch) {
  &.is-checked .el-switch__core {
    background-color: #1890ff;
    border-color: #1890ff;
  }
}

:deep(.el-tag) {
  border-radius: 4px;
  font-weight: 500;
}

// Upload 组件样式
:deep(.el-upload) {
  .el-upload__tip {
    font-size: 12px;
    color: #909399;
    margin-top: 8px;
    line-height: 1.4;
  }
}

:deep(.el-upload-list) {
  margin-top: 12px;
  
  .el-upload-list__item {
    border-radius: 6px;
    border: 1px solid #e4e7ed;
    
    &.is-success {
      border-color: #67c23a;
      background-color: #f0f9ff;
    }
    
    .el-upload-list__item-name {
      color: #303133;
      font-weight: 500;
    }
    
    .el-upload-list__item-status-label {
      .el-icon {
        color: #67c23a;
      }
    }
  }
}
</style>