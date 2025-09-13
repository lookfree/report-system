import { ref } from 'vue'
import { ElMessage } from 'element-plus'

export function useTableConfig() {
  const selectedRows = ref(new Set())
  const selectedColumns = ref(new Set())
  const contentRows = ref(new Map())
  const showTableToolbar = ref({})
  const toolbarPosition = ref({})

  // 内容行管理
  const getContentRows = (sectionId) => {
    if (!contentRows.value.has(sectionId)) {
      const columnCount = 5 // 默认5列，实际会根据模板调整
      
      contentRows.value.set(sectionId, [
        // 默认添加10行可编辑的数据行
        ...Array.from({ length: 10 }, (_, i) => ({
          name: `行${i + 1}`,
          type: 'data',
          cells: new Array(columnCount).fill(''),
          locked: false
        }))
      ])
    }
    return contentRows.value.get(sectionId)
  }

  const onContentRowChange = (sectionId, rowIndex, colIndex, value) => {
    const rows = getContentRows(sectionId)
    if (!rows[rowIndex].cells) {
      rows[rowIndex].cells = []
    }
    rows[rowIndex].cells[colIndex] = value
  }

  // 行列选择
  const selectRow = (sectionId, rowIndex) => {
    const key = `${sectionId}_${rowIndex}`
    if (selectedRows.value.has(key)) {
      selectedRows.value.delete(key)
    } else {
      selectedRows.value.add(key)
    }
  }

  const selectColumn = (sectionId, colIndex) => {
    const key = `${sectionId}_${colIndex}`
    if (selectedColumns.value.has(key)) {
      selectedColumns.value.delete(key)
    } else {
      selectedColumns.value.add(key)
    }
  }

  const isRowSelected = (sectionId, rowIndex) => {
    return selectedRows.value.has(`${sectionId}_${rowIndex}`)
  }

  const isColumnSelected = (sectionId, colIndex) => {
    return selectedColumns.value.has(`${sectionId}_${colIndex}`)
  }

  // 表格工具栏
  const showTableMenu = (event, sectionId) => {
    event.preventDefault()
    showTableToolbar.value[sectionId] = true
    toolbarPosition.value = {
      left: event.pageX + 'px',
      top: event.pageY + 'px'
    }
  }

  const hideTableToolbar = () => {
    Object.keys(showTableToolbar.value).forEach(key => {
      showTableToolbar.value[key] = false
    })
  }

  // 计算合并列的跨度
  const getMergedColspan = (sectionId, parentHeader, template) => {
    const section = template?.structure?.sections.find(s => s.id === sectionId)
    if (!section?.tableStructure?.headers) return 1
    
    let count = 0
    section.tableStructure.headers.forEach(header => {
      if (header.parentHeader && header.parentHeader.trim() === parentHeader) {
        count++
      }
    })
    
    return count > 0 ? count : 1
  }

  // 表格操作
  const addTableRow = (sectionId, template) => {
    const section = template.structure.sections.find(s => s.id === sectionId)
    if (!section) return
    
    const columnCount = section.tableStructure.headers.length
    const rows = getContentRows(sectionId)
    const dataRows = rows.filter(row => row.type === 'data')
    const nextRowNumber = dataRows.length + 1
    
    const newRow = {
      name: sectionId === 'default_header_table' ? '' : `行${nextRowNumber}`,
      type: 'data',
      cells: new Array(columnCount).fill(''),
      locked: false
    }
    
    // 添加到数据行区域的末尾
    rows.push(newRow)
    
    // 触发重新渲染
    contentRows.value.set(sectionId, [...rows])
    
    hideTableToolbar()
    ElMessage.success('已添加新行')
  }

  const addTableColumn = (sectionId, template) => {
    const section = template.structure.sections.find(s => s.id === sectionId)
    if (!section) return
    
    section.tableStructure.columnCount++
    
    // 添加新列到末尾
    const insertIndex = section.tableStructure.headers.length
    const newHeader = {
      index: insertIndex,
      name: sectionId === 'default_header_table' ? '' : `新列${insertIndex + 1}`,
      originalName: sectionId === 'default_header_table' ? '' : `新列${insertIndex + 1}`,
      parentHeader: '',
      dataType: 'FIXED',
      value: '',
      sqlQuery: '',
      dataSourceId: null
    }
    
    section.tableStructure.headers.push(newHeader)
    
    // 重新计算所有表头的index
    section.tableStructure.headers.forEach((header, index) => {
      header.index = index
    })
    
    // 更新内容行的单元格
    const rows = getContentRows(sectionId)
    if (rows) {
      rows.forEach(row => {
        if (!row.cells) {
          row.cells = new Array(section.tableStructure.headers.length).fill('')
        } else {
          row.cells.push('')
        }
      })
      // 触发重新渲染
      contentRows.value.set(sectionId, [...rows])
    }
    
    hideTableToolbar()
    ElMessage.success('已添加新列')
  }

  const deleteTableRow = (sectionId) => {
    if (selectedRows.value.size === 0) {
      ElMessage.warning('请先选择要删除的行')
      return
    }
    
    const rows = getContentRows(sectionId)
    
    // 获取选中行的索引并过滤可删除的行
    const rowIndices = Array.from(selectedRows.value)
      .filter(key => key.startsWith(sectionId + '_'))
      .map(key => {
        const parts = key.split('_')
        return parseInt(parts[parts.length - 1])
      })
      .filter(index => {
        // 检查行是否可以删除（不是locked的系统行）
        return index >= 0 && 
               index < rows.length && 
               rows[index] && 
               !rows[index].locked
      })
      .sort((a, b) => b - a) // 从后往前删除
    
    if (rowIndices.length === 0) {
      ElMessage.warning('无法删除系统配置行，只能删除数据行')
      return
    }
    
    rowIndices.forEach(index => {
      rows.splice(index, 1)
    })
    
    // 重新编号剩余的数据行
    const dataRows = rows.filter(row => row.type === 'data')
    dataRows.forEach((row, index) => {
      row.name = `行${index + 1}`
    })
    
    selectedRows.value.clear()
    hideTableToolbar()
    ElMessage.success(`已删除 ${rowIndices.length} 个数据行`)
  }

  const deleteTableColumn = (sectionId, template) => {
    const section = template.structure.sections.find(s => s.id === sectionId)
    if (!section) return
    
    if (selectedColumns.value.size === 0) {
      ElMessage.warning('请先选择要删除的列')
      return
    }
    
    const colIndices = Array.from(selectedColumns.value)
      .filter(key => key.startsWith(sectionId + '_'))
      .map(key => parseInt(key.split('_')[1]))
      .sort((a, b) => b - a) // 从后往前删除
    
    if (colIndices.length === 0) {
      return
    }
    
    // 删除列
    colIndices.forEach(index => {
      if (index >= 0 && index < section.tableStructure.headers.length) {
        section.tableStructure.headers.splice(index, 1)
        section.tableStructure.columnCount--
        
        // 更新内容行的单元格
        const rows = getContentRows(sectionId)
        rows.forEach(row => {
          if ((row.type === 'data' || row.type === 'custom') && row.cells && index < row.cells.length) {
            row.cells.splice(index, 1)
          }
        })
      }
    })
    
    // 重新计算列索引
    section.tableStructure.headers.forEach((header, index) => {
      header.index = index
    })
    
    selectedColumns.value.clear()
    hideTableToolbar()
    ElMessage.success('已删除选中的列')
  }

  return {
    selectedRows,
    selectedColumns,
    contentRows,
    showTableToolbar,
    toolbarPosition,
    getContentRows,
    onContentRowChange,
    selectRow,
    selectColumn,
    isRowSelected,
    isColumnSelected,
    showTableMenu,
    hideTableToolbar,
    getMergedColspan,
    addTableRow,
    addTableColumn,
    deleteTableRow,
    deleteTableColumn
  }
}