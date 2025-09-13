<template>
  <div class="templates-page">
    <!-- Header -->
    <div class="page-header">
      <div class="header-content">
        <h1>报告模板管理</h1>
        <div class="header-actions">
          <el-button type="primary" @click="showUploadDialog = true">
            <el-icon><Upload /></el-icon>
            上传新模板
          </el-button>
        </div>
      </div>
    </div>

    <!-- Table -->
    <div class="table-container">
      <el-table 
        :data="paginatedTemplates" 
        style="width: 100%" 
        v-loading="loading"
        :header-cell-style="{ backgroundColor: '#fafafa', color: '#606266', fontWeight: '500' }"
      >
        <el-table-column prop="name" label="名称" min-width="200" show-overflow-tooltip />
        <el-table-column label="描述" min-width="250" show-overflow-tooltip>
          <template #default="{ row }">
            <span>{{ getTemplateDescription(row) }}</span>
          </template>
        </el-table-column>
        
        <el-table-column label="配置状态" width="120" align="center">
          <template #default="{ row }">
            <el-tag 
              :type="getConfigStatus(row).type" 
              size="small"
              style="border-radius: 12px; font-weight: 500;"
            >
              {{ getConfigStatus(row).text }}
            </el-tag>
          </template>
        </el-table-column>
        
        <el-table-column label="状态" width="100" align="center">
          <template #default="{ row }">
            <el-icon 
              :color="row.id ? '#67C23A' : '#F56C6C'"
              size="14"
              style="margin-right: 4px;"
            >
              <CircleFilled />
            </el-icon>
            <span :style="{ color: row.id ? '#67C23A' : '#F56C6C', fontSize: '14px', fontWeight: '500' }">
              {{ row.id ? '启用' : '禁用' }}
            </span>
          </template>
        </el-table-column>

        <el-table-column label="操作" width="280">
          <template #default="{ row }">
            <div class="action-buttons">
              <el-button link type="primary" @click="editTemplate(row)">
                编辑
              </el-button>
              <el-button link type="success" @click="openTinyMCEEditor(row)">
                Word编辑器
              </el-button>
              <el-dropdown @command="handleCommand" trigger="click">
                <el-button link type="primary">
                  更多
                  <el-icon class="el-icon--right"><ArrowDown /></el-icon>
                </el-button>
                <template #dropdown>
                  <el-dropdown-menu>
                    <el-dropdown-item :command="{ action: 'generate', row }">
                      <el-icon><Document /></el-icon>
                      生成报告
                    </el-dropdown-item>
                    <el-dropdown-item :command="{ action: 'duplicate', row }">
                      <el-icon><CopyDocument /></el-icon>
                      复制模板
                    </el-dropdown-item>
                    <el-dropdown-item :command="{ action: 'toggle', row }" divided>
                      <el-icon>
                        <CircleCheck v-if="!row.id" />
                        <CircleClose v-else />
                      </el-icon>
                      {{ row.id ? '禁用模板' : '启用模板' }}
                    </el-dropdown-item>
                    <el-dropdown-item :command="{ action: 'delete', row }">
                      <el-icon><Delete /></el-icon>
                      删除模板
                    </el-dropdown-item>
                  </el-dropdown-menu>
                </template>
              </el-dropdown>
            </div>
          </template>
        </el-table-column>
      </el-table>
      
      <!-- Pagination -->
      <div class="pagination-container">
        <div class="pagination-info">
          共 {{ templates.length }} 条
        </div>
        <el-pagination
          v-model:current-page="currentPage"
          v-model:page-size="pageSize"
          :page-sizes="[10, 20, 50, 100]"
          :total="templates.length"
          layout="sizes, prev, pager, next, jumper"
          background
          small
        />
      </div>
    </div>
    
    <!-- Upload Dialog -->
    <el-dialog v-model="showUploadDialog" title="上传模板" width="500px">
      <el-form :model="uploadForm" label-width="100px">
        <el-form-item label="模板名称">
          <el-input v-model="uploadForm.name" placeholder="请输入模板名称" />
        </el-form-item>
        <el-form-item label="选择文件">
          <el-upload
            ref="uploadRef"
            :auto-upload="false"
            :limit="1"
            accept=".doc,.docx"
            :on-change="handleFileChange"
          >
            <el-button type="primary">
              <el-icon><Upload /></el-icon>
              选择Word文档
            </el-button>
            <template #tip>
              <div class="el-upload__tip">只支持 .doc 和 .docx 格式</div>
            </template>
          </el-upload>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showUploadDialog = false">取消</el-button>
        <el-button type="primary" @click="handleUpload" :loading="uploading">
          上传
        </el-button>
      </template>
    </el-dialog>
    
    <!-- Edit Template Dialog -->
    <el-dialog v-model="showEditDialog" title="编辑模板" width="600px">
      <el-form :model="editForm" label-width="100px" label-position="right">
        <el-form-item label="模板名称" required>
          <el-input 
            v-model="editForm.name" 
            placeholder="请输入模板名称"
            style="width: 100%;"
          />
        </el-form-item>
        <el-form-item label="模板描述">
          <el-input 
            v-model="editForm.description" 
            type="textarea"
            :rows="3"
            placeholder="请输入模板描述"
            style="width: 100%;"
          />
        </el-form-item>
        <el-form-item label="流程类型" required>
          <el-select 
            v-model="editForm.workflow" 
            placeholder="请选择流程类型"
            style="width: 100%;"
          >
            <el-option label="流程" value="流程" />
            <el-option label="邮件" value="邮件" />
          </el-select>
        </el-form-item>
        
        <!-- 选择流程时显示流程选择框 -->
        <el-form-item 
          v-if="editForm.workflow === '流程'"
          label="流程选择" 
          required
        >
          <el-select 
            v-model="editForm.selectedProcess" 
            placeholder="请选择具体流程"
            style="width: 100%;"
          >
            <el-option 
              v-for="option in processOptions"
              :key="option.value"
              :label="option.label"
              :value="option.value"
            />
          </el-select>
        </el-form-item>
        
        <!-- 选择邮件时显示邮箱输入框 -->
        <el-form-item 
          v-if="editForm.workflow === '邮件'"
          label="邮箱地址" 
          required
        >
          <el-input 
            v-model="editForm.emailAddress" 
            type="email"
            placeholder="请输入邮箱地址"
            style="width: 100%;"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <div style="text-align: right;">
          <el-button @click="showEditDialog = false">取消</el-button>
          <el-button type="primary" @click="handleEditSave" :loading="saving">
            保存
          </el-button>
        </div>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted, computed } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import api from '@/api'
import dayjs from 'dayjs'

const router = useRouter()
const templates = ref([])
const loading = ref(false)
const showUploadDialog = ref(false)
const showEditDialog = ref(false)
const uploading = ref(false)
const saving = ref(false)
const uploadRef = ref()
const currentPage = ref(1)
const pageSize = ref(10)

const uploadForm = ref({
  name: '',
  file: null
})

const editForm = ref({
  id: null,
  name: '',
  description: '',
  workflow: '流程',
  selectedProcess: '',
  emailAddress: ''
})

// 流程选项列表
const processOptions = ref([
  { value: 'approval', label: '审批流程' },
  { value: 'review', label: '审核流程' },
  { value: 'notification', label: '通知流程' },
  { value: 'custom', label: '自定义流程' }
])

const getConfigStatus = (row) => {
  const hasConfigs = row.hasConfigs
  return {
    type: hasConfigs ? 'success' : 'warning',
    text: hasConfigs ? '已配置' : '待配置'
  }
}

const getTemplateDescription = (row) => {
  if (!row.originalFileName) {
    return '无描述'
  }
  
  try {
    // 移除文件扩展名并处理可能的编码问题
    let description = row.originalFileName.replace(/\.(doc|docx)$/i, '')
    
    // 检查是否包含乱码字符，如果有则使用模板名称
    const hasGarbledText = /[��]/.test(description) || 
                          description.includes('ÿ') || 
                          /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/.test(description)
    
    if (hasGarbledText || description.length === 0) {
      return row.name || '无描述'
    }
    
    return description
  } catch (error) {
    console.warn('Error processing template description:', error)
    return row.name || '无描述'
  }
}

// 分页计算
const paginatedTemplates = computed(() => {
  const start = (currentPage.value - 1) * pageSize.value
  const end = start + pageSize.value
  return templates.value.slice(start, end)
})

const formatDate = (date) => {
  return dayjs(date).format('YYYY-MM-DD HH:mm')
}

const loadTemplates = async () => {
  loading.value = true
  try {
    const templateList = await api.getTemplates()
    // 为每个模板加载配置状态
    templates.value = await Promise.all(
      templateList.map(async (template) => {
        try {
          const templateDetail = await api.getTemplate(template.id)
          template.configs = templateDetail.configs || []
          template.hasConfigs = templateDetail.configs && templateDetail.configs.length > 0
        } catch (error) {
          template.configs = []
          template.hasConfigs = false
        }
        return template
      })
    )
  } catch (error) {
    ElMessage.error('加载模板失败')
  }
  loading.value = false
}

const handleFileChange = (file) => {
  uploadForm.value.file = file.raw
  if (!uploadForm.value.name) {
    uploadForm.value.name = file.name.replace(/\.(doc|docx)$/i, '')
  }
}

const handleUpload = async () => {
  if (!uploadForm.value.file) {
    ElMessage.warning('请选择文件')
    return
  }
  
  uploading.value = true
  const formData = new FormData()
  formData.append('name', uploadForm.value.name)
  formData.append('template', uploadForm.value.file)
  
  try {
    const result = await api.uploadTemplate(formData)
    ElMessage.success('模板上传成功，正在解析模板结构...')
    showUploadDialog.value = false
    uploadForm.value = { name: '', file: null }
    uploadRef.value?.clearFiles()
    
    // 重新加载模板列表
    await loadTemplates()
    
    // 自动跳转到TinyMCE编辑器页面
    if (result.template?.id) {
      ElMessage.success('模板解析完成，即将跳转到Word编辑器')
      setTimeout(() => {
        router.push(`/templates/${result.template.id}/editor`)
      }, 1500)
    }
  } catch (error) {
    ElMessage.error('模板上传失败: ' + (error.response?.data?.error || '未知错误'))
  }
  uploading.value = false
}



const openTinyMCEEditor = (template) => {
  router.push(`/templates/${template.id}/editor`)
}

const editTemplate = (template) => {
  editForm.value = {
    id: template.id,
    name: template.name,
    description: getTemplateDescription(template),
    workflow: template.workflow || '流程',
    selectedProcess: template.selectedProcess || '',
    emailAddress: template.emailAddress || ''
  }
  showEditDialog.value = true
}

const handleCommand = ({ action, row }) => {
  switch (action) {
    case 'generate':
      generateReport(row)
      break
    case 'duplicate':
      duplicateTemplate(row)
      break
    case 'toggle':
      toggleTemplateStatus(row)
      break
    case 'delete':
      deleteTemplate(row)
      break
  }
}

const generateReport = async (template) => {
  try {
    await api.generateReport(template.id)
    ElMessage.success('报告生成成功')
    router.push('/reports')
  } catch (error) {
    ElMessage.error('报告生成失败')
  }
}

const duplicateTemplate = async (template) => {
  try {
    const { value: newName } = await ElMessageBox.prompt(
      '请输入新模板的名称',
      '复制模板',
      {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        inputValue: `${template.name}_副本`,
        inputValidator: (value) => {
          if (!value || value.trim() === '') {
            return '模板名称不能为空'
          }
          return true
        }
      }
    )
    
    // 调用复制API
    await api.duplicateTemplate(template.id, { name: newName.trim() })
    ElMessage.success('模板复制成功')
    
    // 重新加载模板列表
    loadTemplates()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('复制失败: ' + (error.response?.data?.error || error.message || '未知错误'))
    }
  }
}

const toggleTemplateStatus = async (template) => {
  const action = template.id ? '禁用' : '启用'
  const actionStatus = template.id ? 'disabled' : 'enabled'
  
  try {
    await ElMessageBox.confirm(
      `确定要${action}模板"${template.name}"吗？`,
      `${action}确认`,
      {
        confirmButtonText: `确定${action}`,
        cancelButtonText: '取消',
        type: 'info'
      }
    )
    
    // 调用启用/禁用API
    await api.updateTemplateStatus(template.id || template.tempId, { status: actionStatus })
    ElMessage.success(`模板${action}成功`)
    
    // 重新加载模板列表
    loadTemplates()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error(`${action}失败: ` + (error.response?.data?.error || error.message || '未知错误'))
    }
  }
}

const deleteTemplate = async (template) => {
  try {
    await ElMessageBox.confirm(
      `确定要删除模板"${template.name}"吗？此操作不可恢复。`,
      '删除确认',
      {
        confirmButtonText: '确定删除',
        cancelButtonText: '取消',
        type: 'warning',
        confirmButtonClass: 'el-button--danger'
      }
    )
    
    // 调用删除API
    await api.deleteTemplate(template.id)
    ElMessage.success('模板删除成功')
    
    // 重新加载模板列表
    loadTemplates()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败: ' + (error.response?.data?.error || error.message || '未知错误'))
    }
  }
}

const handleEditSave = async () => {
  if (!editForm.value.name.trim()) {
    ElMessage.warning('模板名称不能为空')
    return
  }
  
  // 根据workflow类型验证必填字段
  if (editForm.value.workflow === '流程' && !editForm.value.selectedProcess) {
    ElMessage.warning('请选择流程类型')
    return
  }
  
  if (editForm.value.workflow === '邮件' && !editForm.value.emailAddress.trim()) {
    ElMessage.warning('请填写邮箱地址')
    return
  }
  
  saving.value = true
  try {
    await api.updateTemplate(editForm.value.id, {
      name: editForm.value.name.trim(),
      description: editForm.value.description.trim(),
      workflow: editForm.value.workflow,
      selectedProcess: editForm.value.selectedProcess,
      emailAddress: editForm.value.emailAddress.trim()
    })
    ElMessage.success('模板更新成功')
    showEditDialog.value = false
    
    // 重新加载模板列表
    await loadTemplates()
  } catch (error) {
    ElMessage.error('更新失败: ' + (error.response?.data?.error || error.message || '未知错误'))
  }
  saving.value = false
}


onMounted(() => {
  loadTemplates()
})
</script>

<style scoped lang="scss">
.templates-page {
  padding: 0;
  
  .page-header {
    background: #fff;
    padding: 16px 24px;
    border-bottom: 1px solid #e8eaec;
    margin-bottom: 16px;
    
    .header-content {
      display: flex;
      justify-content: space-between;
      align-items: center;
      
      h1 {
        margin: 0;
        font-size: 20px;
        font-weight: 500;
        color: #1f2329;
      }
      
      .header-actions {
        display: flex;
        gap: 8px;
        align-items: center;
      }
    }
  }
  
  .table-container {
    background: #fff;
    border-radius: 6px;
    border: 1px solid #e8eaec;
    
    .action-buttons {
      display: flex;
      align-items: center;
      gap: 8px;
      
      .el-button {
        padding: 0;
        min-height: auto;
        font-size: 14px;
        
        &.is-link {
          color: #1677ff;
          
          &:hover {
            color: #4096ff;
          }
        }
      }
      
      .el-dropdown {
        .el-button {
          display: flex;
          align-items: center;
          gap: 4px;
        }
      }
    }
    
    :deep(.el-table) {
      border: none;
      
      .el-table__header {
        th {
          border-bottom: 1px solid #e8eaec;
          background-color: #fafafa !important;
          
          .cell {
            color: #606266;
            font-weight: 500;
          }
        }
      }
      
      .el-table__body {
        tr {
          &:hover {
            background-color: #f5f7fa;
          }
          
          td {
            border-bottom: 1px solid #e8eaec;
            
            .cell {
              color: #1f2329;
              font-size: 14px;
            }
          }
        }
      }
    }
    
    .pagination-container {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 16px 24px;
      border-top: 1px solid #e8eaec;
      
      .pagination-info {
        color: #86909c;
        font-size: 14px;
      }
      
      :deep(.el-pagination) {
        .el-pagination__sizes {
          .el-select {
            .el-select__wrapper {
              height: 28px;
              min-height: 28px;
            }
          }
        }
        
        .btn-prev,
        .btn-next {
          height: 28px;
          min-width: 28px;
        }
        
        .el-pager li {
          height: 28px;
          min-width: 28px;
          line-height: 28px;
          font-size: 14px;
        }
        
        .el-pagination__jump {
          .el-input {
            .el-input__wrapper {
              height: 28px;
              
              .el-input__inner {
                height: 28px;
              }
            }
          }
        }
      }
    }
  }
}

// Dialog样式调整
:deep(.el-dialog) {
  .el-dialog__header {
    border-bottom: 1px solid #e8eaec;
    padding: 16px 24px;
    
    .el-dialog__title {
      font-size: 16px;
      font-weight: 500;
    }
  }
  
  .el-dialog__body {
    padding: 24px;
  }
  
  .el-dialog__footer {
    border-top: 1px solid #e8eaec;
    padding: 12px 24px;
  }
}

// 下拉菜单样式
:deep(.el-dropdown-menu) {
  .el-dropdown-menu__item {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 14px;
    
    .el-icon {
      font-size: 16px;
    }
  }
}
</style>