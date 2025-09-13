import { ref } from 'vue'

export function useGridHelpers() {
  
  // 检查是否有合并表头
  const hasMergedHeaders = (headers) => {
    if (!headers || headers.length === 0) return false
    return headers.some(header => header.parentHeader && header.parentHeader.trim() !== '')
  }

  // 获取合并表头分组
  const getMergedHeaderGroups = (headers) => {
    if (!headers || headers.length === 0) return []
    
    const groups = []
    const processedParents = new Set()
    
    let i = 0
    while (i < headers.length) {
      const header = headers[i]
      
      // 独立列（没有parentHeader或parentHeader为空）
      if (!header.parentHeader || header.parentHeader.trim() === '') {
        groups.push({
          name: header.name,
          colspan: 1,
          type: 'independent',
          startIndex: i  // 记录在headers数组中的索引
        })
        i++
      } else {
        // 合并列
        const parentName = header.parentHeader.trim()
        
        if (!processedParents.has(parentName)) {
          // 计算相同parentHeader的列数
          let childCount = 0
          let j = i
          while (j < headers.length && 
                 headers[j].parentHeader && 
                 headers[j].parentHeader.trim() === parentName) {
            childCount++
            j++
          }
          
          groups.push({
            name: parentName,
            colspan: childCount,
            type: 'merged',
            startIndex: i  // 记录开始索引
          })
          
          processedParents.add(parentName)
          i = j // 跳到下一个不同的父表头或独立列
        } else {
          i++
        }
      }
    }
    
    return groups
  }

  // RevoGrid 数据转换方法 - 支持多级表头
  const getGridColumns = (template, sectionId) => {
    const section = template?.structure?.sections.find(s => s.id === sectionId)
    if (!section?.tableStructure?.headers) return []
    
    const headers = section.tableStructure.headers
    
    // 检查是否有合并表头
    const hasMergedHeaders = headers.some(h => h.parentHeader && h.parentHeader.trim() !== '')
    
    if (hasMergedHeaders) {
      // 构建多级表头结构
      const columnGroups = []
      const processedParents = new Set()
      
      let i = 0
      while (i < headers.length) {
        const header = headers[i]
        
        if (!header.parentHeader || header.parentHeader.trim() === '') {
          // 独立列
          columnGroups.push({
            prop: `col_${i}`,
            name: header.name || `列${i + 1}`,
            size: 120,
            sortable: true,
            filter: true
          })
          i++
        } else {
          // 合并列组
          const parentName = header.parentHeader.trim()
          
          if (!processedParents.has(parentName)) {
            // 找到相同parentHeader的所有子列
            const childColumns = []
            let j = i
            
            while (j < headers.length && 
                   headers[j].parentHeader && 
                   headers[j].parentHeader.trim() === parentName) {
              childColumns.push({
                prop: `col_${j}`,
                name: headers[j].name || `子列${j + 1}`,
                size: 120,
                sortable: true,
                filter: true
              })
              j++
            }
            
            // 创建列组 - 使用正确的RevoGrid语法
            columnGroups.push({
              name: parentName,
              children: childColumns
            })
            
            processedParents.add(parentName)
            i = j
          } else {
            i++
          }
        }
      }
      
      return columnGroups
    }
    
    // 简单表头
    return headers.map((header, index) => ({
      prop: `col_${index}`,
      name: header.name || `列${index + 1}`,
      size: 120,
      sortable: true,
      filter: true
    }))
  }

  const getGridData = (template, sectionId, getContentRows) => {
    const rows = getContentRows(sectionId)
    const section = template?.structure?.sections.find(s => s.id === sectionId)
    if (!rows || !section?.tableStructure?.headers) return []
    
    return rows.map((row, rowIndex) => {
      const rowData = { _id: rowIndex }
      section.tableStructure.headers.forEach((header, colIndex) => {
        rowData[`col_${colIndex}`] = row.cells?.[colIndex] || ''
      })
      return rowData
    })
  }

  // Grid event handlers
  const onCellEdit = (event, contentRows, template, addTableRow) => {
    const { detail } = event
    if (!detail) return
    
    const { prop, rowIndex, val, oldVal } = detail
    const colIndex = parseInt(prop.replace('col_', ''))
    
    // 更新原始数据
    Object.keys(contentRows.value).forEach(sectionId => {
      const rows = contentRows.value.get(sectionId)
      if (rows && rows[rowIndex] && rows[rowIndex].cells) {
        rows[rowIndex].cells[colIndex] = val
        
        // 检查是否是最后一行，如果是则自动添加新行
        if (rowIndex === rows.length - 1 && val && val.trim() !== '') {
          console.log('Adding new row automatically for section:', sectionId)
          
          // 获取列数
          const section = template?.structure?.sections.find(s => s.id === sectionId)
          if (section?.tableStructure?.headers) {
            const columnCount = section.tableStructure.headers.length
            
            // 创建新行
            const newRow = {
              name: `行${rows.length + 1}`,
              type: 'data',
              cells: new Array(columnCount).fill(''),
              locked: false
            }
            
            // 添加到行数据中
            rows.push(newRow)
            
            // 触发重新渲染
            contentRows.value.set(sectionId, [...rows])
            
            console.log('New row added automatically:', newRow)
          }
        }
      }
    })
  }

  const onBeforeEdit = (event) => {
    const { detail } = event
    // 可以在此处添加编辑前的验证逻辑
    return true
  }

  const onRowHeaderClick = (event) => {
    const { detail } = event
    console.log('Row header clicked:', detail)
    // 选择整行
  }

  const onColHeaderClick = (event) => {
    const { detail } = event
    console.log('Column header clicked:', detail)
    // 选择整列
  }

  return {
    hasMergedHeaders,
    getMergedHeaderGroups,
    getGridColumns,
    getGridData,
    onCellEdit,
    onBeforeEdit,
    onRowHeaderClick,
    onColHeaderClick
  }
}