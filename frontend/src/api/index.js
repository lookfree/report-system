import axios from 'axios'

const api = axios.create({
  baseURL: '/api',
  timeout: 30000
})

// Request interceptor
api.interceptors.request.use(
  config => config,
  error => Promise.reject(error)
)

// Response interceptor
api.interceptors.response.use(
  response => response.data,
  error => {
    console.error('API Error:', error)
    return Promise.reject(error)
  }
)

export default {
  // Templates
  uploadTemplate(formData) {
    return api.post('/templates/upload', formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    })
  },
  getTemplates() {
    return api.get('/templates')
  },
  getTemplate(id) {
    return api.get(`/templates/${id}`)
  },
  getTemplateFullStructure(id) {
    return api.get(`/templates/${id}/full-structure`)
  },
  deleteTemplate(id) {
    return api.delete(`/templates/${id}`)
  },
  updateTemplateStatus(id, data) {
    return api.patch(`/templates/${id}/status`, data)
  },
  duplicateTemplate(id, data) {
    return api.post(`/templates/${id}/duplicate`, data)
  },
  updateTemplate(id, data) {
    return api.put(`/templates/${id}`, data)
  },
  saveTemplateConfig(id, configs, templateStructure) {
    return api.post(`/templates/${id}/config`, { configs, templateStructure })
  },
  fixTemplateDescriptions() {
    return api.post('/templates/fix-descriptions')
  },
  
  // TinyMCE Word Editor
  importWordToHtml(formData) {
    return api.post('/templates/import-word', formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    })
  },
  saveTemplateHtml(id, htmlContent) {
    return api.post(`/templates/${id}/save-html`, { html: htmlContent })
  },
  exportHtmlToWord(id, htmlContent) {
    return api.post(`/templates/${id}/export-word`, { htmlContent: htmlContent }, {
      responseType: 'blob'
    })
  },
  processTemplateVariables(id, htmlContent) {
    return api.post(`/templates/${id}/process-variables`, { html: htmlContent })
  },
  
  // Data sources
  getDataSources() {
    return api.get('/datasources')
  },
  createDataSource(data) {
    return api.post('/datasources', data)
  },
  testDataSource(data) {
    return api.post('/datasources/test', data)
  },
  getDataSourceFields(dataSourceId) {
    return api.get(`/datasources/${dataSourceId}/fields`)
  },
  
  // Tasks
  getTasks() {
    return api.get('/tasks')
  },
  createTask(data) {
    return api.post('/tasks', data)
  },
  toggleTask(id) {
    return api.put(`/tasks/${id}/toggle`)
  },
  
  // Reports
  generateReport(templateId) {
    return api.post(`/reports/generate/${templateId}`)
  },
  getReports() {
    return api.get('/reports')
  },
  downloadReport(id) {
    window.open(`/api/reports/download/${id}`)
  }
}