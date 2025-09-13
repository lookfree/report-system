import { ref } from 'vue'
import { ElMessage } from 'element-plus'
import api from '@/api'

export function useTemplateConfig() {
  const template = ref(null)
  const loading = ref(true)
  const saving = ref(false)
  const configs = ref([])
  const dataSources = ref([])
  const showTestDialog = ref(false)
  const testing = ref(false)
  const testResult = ref({})

  const getConfigForSection = (sectionId) => {
    let config = configs.value.find(c => c.sectionId === sectionId)
    if (!config) {
      config = {
        sectionId,
        sectionName: '',
        dataType: 'FIXED',
        value: '',
        sqlQuery: '',
        dataSourceId: null
      }
      configs.value.push(config)
    }
    return config
  }

  const onColumnConfigChange = (sectionId, columnIndex, columnConfig) => {
    // 为每个列创建独立的配置记录
    const configId = `${sectionId}_col_${columnIndex}`
    let config = configs.value.find(c => c.sectionId === configId)
    
    if (!config) {
      config = {
        sectionId: configId,
        sectionName: `${columnConfig.name}列`,
        dataType: columnConfig.dataType,
        value: columnConfig.value,
        sqlQuery: columnConfig.sqlQuery,
        dataSourceId: columnConfig.dataSourceId,
        columnIndex: columnIndex,
        parentSectionId: sectionId
      }
      configs.value.push(config)
    } else {
      config.dataType = columnConfig.dataType
      config.value = columnConfig.value
      config.sqlQuery = columnConfig.sqlQuery
      config.dataSourceId = columnConfig.dataSourceId
    }
  }

  const resetTableData = (sectionId) => {
    const section = template.value.structure.sections.find(s => s.id === sectionId)
    if (!section || !section.tableStructure) return
    
    // 重置所有表头的合并状态
    section.tableStructure.headers.forEach(header => {
      if (header.parentHeader) {
        header.parentHeader = ''
        header.name = header.originalName || header.name
      }
    })
    
    ElMessage.success('已拆分合并单元格')
  }

  const initializeTableConfigs = () => {
    if (!template.value.structure || !template.value.structure.sections) return
    
    template.value.structure.sections.forEach(section => {
      if (section.hasTable && section.tableStructure?.headers) {
        section.tableStructure.headers.forEach((header, index) => {
          const configId = `${section.id}_col_${index}`
          
          // 检查是否已有配置
          const existingConfig = configs.value.find(c => {
            return c.sectionId === configId || 
              (c.parentSectionId === section.id && c.columnIndex === index)
          })
          
          // 更新表头对象的配置
          if (existingConfig) {
            header.dataType = existingConfig.dataType
            header.value = existingConfig.value
            header.sqlQuery = existingConfig.sqlQuery
            header.dataSourceId = existingConfig.dataSourceId
          }
        })
      }
    })
  }

  const loadDataSources = async () => {
    try {
      dataSources.value = await api.getDataSources()
    } catch (error) {
      ElMessage.error('加载数据源失败')
    }
  }

  // 添加默认表格（报告开头内容表格）
  const addDefaultHeaderTableIfNeeded = (template) => {
    if (!template?.structure?.sections) {
      template.structure = { sections: [] }
    }
    
    // 检查是否已经有默认表格
    const hasDefaultTable = template.structure.sections.some(section => 
      section.id === 'default_header_table' || section.name === '报告开头内容'
    )
    
    if (!hasDefaultTable) {
      // 创建默认空白表格结构（10行10列）
      const headers = []
      for (let i = 0; i < 10; i++) {
        headers.push({
          index: i,
          name: '',
          originalName: '',
          parentHeader: null,
          dataType: 'FIXED',
          value: '',
          sqlQuery: '',
          dataSourceId: null
        })
      }
      
      // 创建10行空白数据
      const tableData = []
      for (let i = 0; i < 10; i++) {
        const cells = new Array(10).fill('')
        tableData.push({ 
          name: '', 
          type: 'data',
          cells: cells,
          locked: false
        })
      }
      
      const defaultHeaderTable = {
        id: 'default_header_table',
        name: '默认表格',
        hasTable: true,
        tableStructure: {
          headers: headers,
          columnCount: 10,
          rowCount: 10
        },
        tableData: tableData
      }
      
      // 将默认表格添加到第一个位置
      template.structure.sections.unshift(defaultHeaderTable)
      
      console.log('✅ 已自动添加默认报告开头内容表格')
    } else {
      console.log('ℹ️ 默认表格已存在，跳过创建')
    }
  }

  const saveConfig = async (templateId, additionalData = {}) => {
    saving.value = true
    try {
      // 构建要保存的配置数据
      const configsToSave = []
      
      // 保存表格名称和数据源配置（从TemplateConfig.vue传递过来）
      if (additionalData.tableNames) {
        Object.keys(additionalData.tableNames).forEach(sectionId => {
          if (additionalData.tableNames[sectionId]) {
            configsToSave.push({
              sectionId: `${sectionId}_name`,
              sectionName: '表格名称',
              dataType: 'FIXED',
              value: additionalData.tableNames[sectionId],
              sqlQuery: '',
              dataSourceId: null,
              columnIndex: null,
              parentSectionId: sectionId
            })
          }
        })
      }
      
      if (additionalData.tableDataSources) {
        Object.keys(additionalData.tableDataSources).forEach(sectionId => {
          if (additionalData.tableDataSources[sectionId]) {
            configsToSave.push({
              sectionId: `${sectionId}_datasource`,
              sectionName: '数据源',
              dataType: 'FIXED',
              value: additionalData.tableDataSources[sectionId],
              sqlQuery: '',
              dataSourceId: additionalData.tableDataSources[sectionId],
              columnIndex: null,
              parentSectionId: sectionId
            })
          }
        })
      }
      
      // 保存备注配置
      if (additionalData.tableRemarks) {
        Object.keys(additionalData.tableRemarks).forEach(sectionId => {
          if (additionalData.tableRemarks[sectionId]) {
            configsToSave.push({
              sectionId: `${sectionId}_remark`,
              sectionName: '备注',
              dataType: 'FIXED',
              value: additionalData.tableRemarks[sectionId],
              sqlQuery: '',
              dataSourceId: null,
              columnIndex: null,
              parentSectionId: sectionId
            })
          }
        })
      }
      
      // 保存统计开关配置
      if (additionalData.statisticsEnabled) {
        Object.keys(additionalData.statisticsEnabled).forEach(sectionId => {
          configsToSave.push({
            sectionId: `${sectionId}_stats`,
            sectionName: '统计开关',
            dataType: 'FIXED',
            value: additionalData.statisticsEnabled[sectionId] ? 'true' : 'false',
            sqlQuery: '',
            dataSourceId: null,
            columnIndex: null,
            parentSectionId: sectionId
          })
        })
      }
      
      
      // 保存单元格配置
      if (additionalData.cellConfigurations) {
        Object.keys(additionalData.cellConfigurations).forEach(cellKey => {
          const config = additionalData.cellConfigurations[cellKey]
          const [sectionId, rowIndex, colIndex] = cellKey.split('_')
          
          // 支持新的配置格式
          let dataType, value, sqlQuery
          
          let datasetId = null
          
          // 统一格式处理
          if (config.displayType === 'text') {
            // 文字模式
            dataType = 'FIXED'
            value = config.textContent || ''
            sqlQuery = ''
          } else {
            // 数据集模式
            dataType = 'DYNAMIC' 
            const fieldsStr = config.displayFields ? config.displayFields.join(',') : ''
            const structureInfo = config.dataStructure || 'single'
            value = `${structureInfo}|${fieldsStr}` // 格式: "数据结构|字段列表"
            sqlQuery = config.displayFields ? `SELECT ${config.displayFields.join(', ')} FROM table` : ''
            datasetId = config.datasetId || null
          }
          
          configsToSave.push({
            sectionId: `cell_${cellKey}`,
            sectionName: `单元格(${rowIndex},${colIndex})`,
            dataType: dataType,
            value: value,
            sqlQuery: sqlQuery,
            dataSourceId: datasetId, // 保存数据集ID
            columnIndex: parseInt(colIndex),
            parentSectionId: sectionId
          })
        })
      }
      
      // 保存统计行配置
      if (additionalData.statisticsConfigurations) {
        Object.keys(additionalData.statisticsConfigurations).forEach(cellKey => {
          const config = additionalData.statisticsConfigurations[cellKey]
          // cellKey 格式为 "stats_sectionId_colIndex"
          const [, sectionId, colIndex] = cellKey.split('_')
          
          configsToSave.push({
            sectionId: cellKey,
            sectionName: `统计单元格(${colIndex})`,
            dataType: config.fillType === 'fixed' ? 'FIXED' : 'DYNAMIC',
            value: config.fillType === 'fixed' ? config.fixedValue : '',
            sqlQuery: config.fillType === 'auto' ? `SELECT ${config.linkedField} FROM table` : '',
            dataSourceId: null, // 暂时设为null，避免外键约束问题
            columnIndex: parseInt(colIndex),
            parentSectionId: sectionId
          })
        })
      }
      
      // 保存合并单元格信息
      if (additionalData.mergedCells) {
        Object.keys(additionalData.mergedCells).forEach(sectionId => {
          const sectionMerges = additionalData.mergedCells[sectionId]
          if (sectionMerges && typeof sectionMerges === 'object') {
            Object.keys(sectionMerges).forEach(mergeKey => {
              const mergeInfo = sectionMerges[mergeKey]
              configsToSave.push({
                sectionId: `merge_${sectionId}_${mergeKey}`,
                sectionName: `合并单元格(${mergeKey})`,
                dataType: 'FIXED',
                value: JSON.stringify({
                  startRow: mergeInfo.startRow,
                  startCol: mergeInfo.startCol,
                  rowspan: mergeInfo.rowspan,
                  colspan: mergeInfo.colspan
                }),
                sqlQuery: '',
                dataSourceId: null,
                columnIndex: null,
                parentSectionId: sectionId
              })
            })
          }
        })
      }
      
      // 合并原有的配置数据
      configsToSave.push(...configs.value)
      
      
      await api.saveTemplateConfig(templateId, configsToSave, additionalData.templateStructure)
      ElMessage.success('配置保存成功')
    } catch (error) {
      console.error('保存配置失败:', error)
      ElMessage.error('配置保存失败: ' + (error.response?.data?.error || error.message))
    }
    saving.value = false
  }

  const testQuery = async (config) => {
    const dataSource = dataSources.value.find(ds => ds.id === config.dataSourceId)
    if (!dataSource) {
      ElMessage.error('请选择数据源')
      return
    }
    
    showTestDialog.value = true
    testing.value = true
    testResult.value = {}
    
    try {
      const result = await api.testDataSource({
        ...dataSource,
        sqlQuery: config.sqlQuery
      })
      testResult.value = {
        data: true,
        rowCount: result.rowCount,
        sample: result.sample || []
      }
    } catch (error) {
      testResult.value = {
        error: error.response?.data?.error || '查询失败'
      }
    }
    testing.value = false
  }

  const testColumnQuery = async (columnConfig) => {
    const dataSource = dataSources.value.find(ds => ds.id === columnConfig.dataSourceId)
    if (!dataSource) {
      ElMessage.error('请选择数据源')
      return
    }
    
    showTestDialog.value = true
    testing.value = true
    testResult.value = {}
    
    try {
      const result = await api.testDataSource({
        ...dataSource,
        sqlQuery: columnConfig.sqlQuery
      })
      testResult.value = {
        data: true,
        rowCount: result.rowCount,
        sample: result.sample || []
      }
    } catch (error) {
      testResult.value = {
        error: error.response?.data?.error || '查询失败'
      }
    }
    testing.value = false
  }

  const loadTemplate = async (templateId, getContentRows, initializeSpreadsheets) => {
    try {
      // 获取模板基本信息和配置
      const templateInfo = await api.getTemplate(templateId)
      
      // 获取完整的结构数据（从缓存文件）
      const fullStructure = await api.getTemplateFullStructure(templateId)
      
      // 合并模板信息和完整结构
      template.value = {
        ...templateInfo,
        structure: fullStructure
      }
      
      // 检查是否需要添加默认表格（报告开头内容表格）
      addDefaultHeaderTableIfNeeded(template.value)
      
      // 加载已有的配置
      if (templateInfo.rs_template_configs && templateInfo.rs_template_configs.length > 0) {
        configs.value = templateInfo.rs_template_configs.map(c => ({
          ...c,
          columnIndex: c.columnIndex,
          parentSectionId: c.parentSectionId
        }))
        
      }
      
      // 初始化表格列配置
      initializeTableConfigs()
      
      // 初始化Jspreadsheet表格
      await new Promise(resolve => setTimeout(resolve, 100)) // nextTick equivalent
      initializeSpreadsheets(template.value, getContentRows)
    } catch (error) {
      console.error('Loading template error:', error)
      ElMessage.error('加载模板失败: ' + (error.response?.data?.error || error.message))
    }
    loading.value = false
  }

  return {
    // State
    template,
    loading,
    saving,
    configs,
    dataSources,
    showTestDialog,
    testing,
    testResult,
    
    // Methods
    getConfigForSection,
    onColumnConfigChange,
    resetTableData,
    initializeTableConfigs,
    loadDataSources,
    saveConfig,
    testQuery,
    testColumnQuery,
    loadTemplate
  }
}