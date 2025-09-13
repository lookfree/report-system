<template>
  <div class="reports-page">
    <el-card>
      <template #header>
        <div class="card-header">
          <span>报告历史</span>
          <el-button @click="loadReports" :loading="loading">
            <el-icon><Refresh /></el-icon>
            刷新
          </el-button>
        </div>
      </template>
      
      <el-table :data="reports" style="width: 100%" v-loading="loading">
        <el-table-column prop="reportName" label="报告名称" />
        <el-table-column label="模板">
          <template #default="{ row }">
            {{ row.template?.name }}
          </template>
        </el-table-column>
        <el-table-column label="来源">
          <template #default="{ row }">
            {{ row.task ? '定时任务' : '手动生成' }}
          </template>
        </el-table-column>
        <el-table-column label="生成时间">
          <template #default="{ row }">
            {{ formatDate(row.generatedAt) }}
          </template>
        </el-table-column>
        <el-table-column label="状态" width="100">
          <template #default="{ row }">
            <el-tag
              :type="row.status === 'SUCCESS' ? 'success' : row.status === 'FAILED' ? 'danger' : 'info'"
            >
              {{ statusText[row.status] }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="100">
          <template #default="{ row }">
            <el-button
              type="primary"
              size="small"
              @click="downloadReport(row)"
              :disabled="row.status !== 'SUCCESS'"
            >
              下载
            </el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import api from '@/api'
import dayjs from 'dayjs'

const reports = ref([])
const loading = ref(false)

const statusText = {
  SUCCESS: '成功',
  FAILED: '失败',
  PENDING: '处理中'
}

const formatDate = (date) => {
  return dayjs(date).format('YYYY-MM-DD HH:mm:ss')
}

const loadReports = async () => {
  loading.value = true
  try {
    reports.value = await api.getReports()
  } catch (error) {
    ElMessage.error('加载报告失败')
  }
  loading.value = false
}

const downloadReport = (report) => {
  api.downloadReport(report.id)
}

onMounted(() => {
  loadReports()
})
</script>

<style scoped lang="scss">
.reports-page {
  .card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
}
</style>