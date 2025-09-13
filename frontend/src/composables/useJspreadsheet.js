import { ref, nextTick } from 'vue'
import jspreadsheet from 'jspreadsheet-ce'
import 'jspreadsheet-ce/dist/jspreadsheet.css'

// 确保jspreadsheet正确加载
console.log('Jspreadsheet imported:', jspreadsheet)
console.log('Jspreadsheet type:', typeof jspreadsheet)
console.log('Jspreadsheet keys:', Object.keys(jspreadsheet || {}))

// 尝试不同的引用方式
let jss = jspreadsheet
if (jspreadsheet && jspreadsheet.default) {
  jss = jspreadsheet.default
  console.log('Using jspreadsheet.default:', jss)
}

// 检查全局window对象
if (typeof window !== 'undefined' && window.jspreadsheet) {
  jss = window.jspreadsheet
  console.log('Using window.jspreadsheet:', jss)
}

export function useJspreadsheet() {
  const spreadsheets = ref(new Map())

  // 动态构建嵌套表格配置
  const buildNestedHeaders = (section) => {
    const headers = section.tableStructure.headers
    console.log('Building nested headers for:', headers.map(h => ({ name: h.name, parentHeader: h.parentHeader })))
    
    // 检查是否有合并表头
    const hasMergedHeaders = headers.some(h => h.parentHeader && h.parentHeader.trim() !== '')
    
    if (!hasMergedHeaders) {
      // 简单表头：只有一行
      return {
        nestedHeaders: [
          [
            ...headers.map(header => ({
              title: header.name || '列',
              colspan: 1
            }))
          ]
        ]
      }
    }
    
    // 复杂表头：构建两层结构
    const firstRow = [] // 第一行：父表头
    const secondRow = [] // 第二行：子表头
    
    // 分组处理表头
    const processedParents = new Set()
    
    let i = 0
    while (i < headers.length) {
      const header = headers[i]
      
      if (!header.parentHeader || header.parentHeader.trim() === '') {
        // 独立列 - 跨两行（审计单位列）
        firstRow.push({ title: header.name || '列', rowspan: 2 })
        // 对于独立列，第二行不需要添加任何内容
        i++
      } else {
        // 合并列组
        const parentName = header.parentHeader.trim()
        
        if (!processedParents.has(parentName)) {
          // 计算同一父表头下的子列数量
          let childCount = 0
          let j = i
          const childHeaders = []
          
          while (j < headers.length && 
                 headers[j].parentHeader && 
                 headers[j].parentHeader.trim() === parentName) {
            childCount++
            childHeaders.push(headers[j].name || `子列${j + 1}`)
            j++
          }
          
          // 添加父表头到第一行
          firstRow.push({ title: parentName, colspan: childCount })
          
          // 添加子表头到第二行
          childHeaders.forEach(childName => {
            secondRow.push(childName)
          })
          
          processedParents.add(parentName)
          i = j
        } else {
          i++
        }
      }
    }
    
    console.log('First row:', firstRow)
    console.log('Second row:', secondRow)
    
    return {
      nestedHeaders: [firstRow, secondRow]
    }
  }

  // 创建Jspreadsheet实例
  const createSpreadsheet = (sectionId, template, getContentRows, onCellClick) => {
    console.log('Creating spreadsheet for section:', sectionId)
    const section = template.structure.sections.find(s => s.id === sectionId)
    if (!section?.tableStructure?.headers) {
      console.log('No headers found for section:', sectionId)
      return
    }
    
    const container = document.getElementById(`spreadsheet-${sectionId}`)
    if (!container) {
      console.log('Container not found for section:', sectionId)
      return
    }
    console.log('Container found, headers count:', section.tableStructure.headers.length)
    console.log('Headers:', section.tableStructure.headers)
    
    // 准备数据
    const headers = section.tableStructure.headers
    console.log('Section headers:', headers)
    
    let rows = []
    try {
      rows = getContentRows(sectionId) || []
    } catch (error) {
      console.error('Error getting content rows:', error)
      rows = []
    }
    console.log('Content rows:', rows)
    
    // 如果没有行数据，创建一些默认行
    if (rows.length === 0) {
      rows = [
        { name: '示例单位1', cells: new Array(headers.length - 1).fill('') },
        { name: '示例单位2', cells: new Array(headers.length - 1).fill('') }
      ]
    }
    
    // 使用原始headers构建列配置，不过滤
    const columns = headers.map(header => ({
      title: header.name || '列',
      width: header.name === '审计单位' ? 120 : 150,
      type: 'text',
      readOnly: false // 暂时设为可编辑，便于测试
    }))
    
    // 构建数据
    const data = rows.map(row => 
      headers.map((header, index) => {
        if (header.name === '审计单位') {
          return row.name || '单位名称'
        }
        return row.cells?.[index] || ''
      })
    )
    
    console.log('Final data:', data)
    console.log('Final columns:', columns)
    
    // 使用实际数据并配置首列只读，启用嵌套表头（若有）
    const columnsWithReadonly = columns.map((col, idx) => ({
      ...col,
      readOnly: idx === 0 ? true : false
    }))

    let config = {
      data: data,
      columns: columnsWithReadonly,
      allowInsertRow: true,
      allowInsertColumn: true,
      allowDeleteRow: true,
      allowDeleteColumn: true,
      allowRenameColumn: true,
      allowComments: false,
      csvHeaders: true,
      search: false,
      pagination: false,
      paginationOptions: false,
      contextMenu: true,
      toolbar: true,
      freezeColumns: 1
    }

    // 嵌套表头（如果识别到合并表头）
    try {
      const nested = buildNestedHeaders(section)
      if (nested && nested.nestedHeaders && nested.nestedHeaders.length > 0) {
        config.nestedHeaders = nested.nestedHeaders
      }
    } catch (e) {
      console.warn('Build nested headers failed:', e)
    }
    
    console.log('Config created with actual data:', config)
    console.log('Data rows:', data.length)
    console.log('Columns:', columns.length)
    
    console.log('Final jspreadsheet config:', config)
    
    // 切换回真实 jSpreadsheet
    const USE_HTML_TABLE = false
    
    if (USE_HTML_TABLE) {
      console.log('Using HTML table for section:', sectionId)
      
      // 创建HTML表格
      let tableHTML = '<div style="overflow: auto; max-height: 400px;">'
      tableHTML += '<table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">'
      
      // 检查是否需要合并表头
      const hasMergedHeaders = headers.some(h => h.parentHeader && h.parentHeader.trim() !== '')
      console.log('Has merged headers for section:', sectionId, hasMergedHeaders)
      
      // 构建表头
      tableHTML += '<thead style="position: sticky; top: 0; background: white; z-index: 10;">'
      
      if (hasMergedHeaders) {
        // 构建两行表头（合并表头）
        const headerGroups = []
        const subHeaders = []
        let currentGroup = null
        
        headers.forEach(header => {
          if (!header.parentHeader || header.parentHeader.trim() === '') {
            // 独立列（如"审计单位"），跨两行
            if (currentGroup) {
              headerGroups.push(currentGroup)
              currentGroup = null
            }
            headerGroups.push({
              title: header.name || '列',
              colspan: 1,
              rowspan: 2,
              isIndependent: true
            })
            // 第二行对应位置不需要单元格（因为被rowspan占用）
          } else {
            // 合并列
            const parentName = header.parentHeader.trim()
            if (!currentGroup || currentGroup.title !== parentName) {
              if (currentGroup) {
                headerGroups.push(currentGroup)
              }
              currentGroup = {
                title: parentName,
                colspan: 1,
                rowspan: 1,
                isIndependent: false,
                children: [header.name || '列']
              }
            } else {
              currentGroup.colspan++
              currentGroup.children.push(header.name || '列')
            }
          }
        })
        
        if (currentGroup) {
          headerGroups.push(currentGroup)
        }
        
        // 生成第一行（父表头）
        tableHTML += '<tr>'
        headerGroups.forEach(group => {
          const rowspan = group.rowspan ? ` rowspan="${group.rowspan}"` : ''
          const colspan = group.colspan > 1 ? ` colspan="${group.colspan}"` : ''
          tableHTML += `<th style="border: 1px solid #ddd; padding: 8px; background: #f0f0f0; font-weight: bold; text-align: center; min-width: 100px;"${rowspan}${colspan}>${group.title}</th>`
        })
        tableHTML += '</tr>'
        
        // 生成第二行（子表头）
        tableHTML += '<tr>'
        headerGroups.forEach(group => {
          if (!group.isIndependent && group.children) {
            group.children.forEach(child => {
              tableHTML += `<th style="border: 1px solid #ddd; padding: 8px; background: #f0f0f0; font-weight: bold; text-align: center; min-width: 100px;">${child}</th>`
            })
          }
        })
        tableHTML += '</tr>'
      } else {
        // 简单表头（单行）
        tableHTML += '<tr>'
        columns.forEach(col => {
          tableHTML += `<th style="border: 1px solid #ddd; padding: 8px; background: #f0f0f0; font-weight: bold; text-align: left; min-width: 100px;">${col.title || '列'}</th>`
        })
        tableHTML += '</tr>'
      }
      
      tableHTML += '</thead>'
      
      // 添加数据行
      tableHTML += '<tbody>'
      if (data && data.length > 0) {
        data.forEach((row, rowIndex) => {
          tableHTML += '<tr>'
          row.forEach((cell, cellIndex) => {
            const bgColor = cellIndex === 0 ? '#f8f9fa' : 'white'
            const fontWeight = cellIndex === 0 ? 'bold' : 'normal'
            const cellId = `cell_${sectionId}_${rowIndex}_${cellIndex}`
            const clickHandler = onCellClick ? `onclick="window.handleCellClick?.('${sectionId}', ${rowIndex}, ${cellIndex})"` : ''
            tableHTML += `<td id="${cellId}" ${clickHandler} style="border: 1px solid #ddd; padding: 8px; background: ${bgColor}; font-weight: ${fontWeight}; cursor: ${onCellClick ? 'pointer' : 'default'};">${cell || ''}</td>`
          })
          tableHTML += '</tr>'
        })
      } else {
        tableHTML += '<tr><td colspan="' + columns.length + '" style="text-align: center; padding: 20px; color: #999;">暂无数据</td></tr>'
      }
      tableHTML += '</tbody></table></div>'
      
      container.innerHTML = tableHTML
      console.log('HTML table created for section:', sectionId)
      return
    }
    
    // 使用jspreadsheet组件
    try {
      console.log('Creating jspreadsheet for section:', sectionId, 'with data rows:', data.length, 'columns:', columns.length)
      
      // 检查jss是否可用
      if (!jss || typeof jss !== 'function') {
        console.error('Jspreadsheet not available, jss:', jss, 'type:', typeof jss)
        throw new Error('Jspreadsheet library not loaded properly')
      }
      
      // 使用最简化配置
      const simpleConfig = {
        data: config.data,
        columns: config.columns
      }
      
      console.log('Using simple config:', simpleConfig)
      const spreadsheet = jss(container, simpleConfig)
      
      // 检查返回值
      if (!spreadsheet) {
        throw new Error('Jspreadsheet returned null/undefined')
      }
      
      console.log('Jspreadsheet created successfully for section:', sectionId)
      spreadsheets.value.set(sectionId, spreadsheet)
      
    } catch (error) {
      console.error('Error creating jspreadsheet for section:', sectionId, error)
      console.error('Error stack:', error.stack)
      
      // 如果jspreadsheet失败，使用简单的HTML表格作为后备方案
      console.log('Falling back to HTML table for section:', sectionId)
      
      try {
        let tableHTML = '<table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">'
        
        // 添加表头
        tableHTML += '<thead><tr>'
        columns.forEach(col => {
          tableHTML += `<th style="border: 1px solid #ddd; padding: 8px; background: #f0f0f0; font-weight: bold;">${col.title}</th>`
        })
        tableHTML += '</tr></thead>'
        
        // 添加数据行
        tableHTML += '<tbody>'
        data.forEach(row => {
          tableHTML += '<tr>'
          row.forEach(cell => {
            tableHTML += `<td style="border: 1px solid #ddd; padding: 8px;">${cell || ''}</td>`
          })
          tableHTML += '</tr>'
        })
        tableHTML += '</tbody></table>'
        
        container.innerHTML = tableHTML
        console.log('HTML table created as fallback for section:', sectionId)
      } catch (fallbackError) {
        console.error('Fallback HTML table also failed:', fallbackError)
        container.innerHTML = `
          <div style="border: 1px solid red; padding: 10px; background: #ffcccc;">
            <h4>表格创建失败 - ${sectionId}</h4>
            <p>错误: ${error.message}</p>
            <p>数据行数: ${data ? data.length : 'N/A'}</p>
            <p>列数: ${columns ? columns.length : 'N/A'}</p>
          </div>
        `
      }
    }
  }

  // 初始化所有表格
  const initializeSpreadsheets = (template, getContentRows, onCellClick) => {
    if (!template?.structure?.sections) return
    
    template.structure.sections.forEach(section => {
      // 对所有有表格的分节初始化动态表格
      if (section.hasTable && section.tableStructure?.headers) {
        console.log('Initializing spreadsheet for section:', section.id, 'type:', section.type)
        createSpreadsheet(section.id, template, getContentRows, onCellClick)
      }
    })
  }

  // 销毁表格实例
  const destroySpreadsheet = (sectionId) => {
    const spreadsheet = spreadsheets.value.get(sectionId)
    if (spreadsheet) {
      spreadsheet.destroy()
      spreadsheets.value.delete(sectionId)
    }
  }

  // 销毁所有表格
  const destroyAllSpreadsheets = () => {
    spreadsheets.value.forEach((spreadsheet, sectionId) => {
      spreadsheet.destroy()
    })
    spreadsheets.value.clear()
  }

  return {
    spreadsheets,
    createSpreadsheet,
    initializeSpreadsheets,
    destroySpreadsheet,
    destroyAllSpreadsheets,
    buildNestedHeaders
  }
}