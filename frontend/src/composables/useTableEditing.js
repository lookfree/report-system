import { ElMessage } from 'element-plus'

export function useTableEditing() {
  
  // 增强表格编辑功能
  const enhanceTableEditing = () => {
    const editorElement = document.getElementById('word-editor')
    if (!editorElement) return
    
    const tables = editorElement.querySelectorAll('table')
    tables.forEach(table => {
      if (!table.previousElementSibling?.classList.contains('table-toolbar')) {
        const toolbar = createTableToolbar()
        table.parentNode.insertBefore(toolbar, table)
      }
      
      const cells = table.querySelectorAll('td, th')
      cells.forEach(cell => {
        cell.addEventListener('contextmenu', showCellContextMenu)
      })
    })
  }

  // 创建表格工具栏
  const createTableToolbar = () => {
    const toolbar = document.createElement('div')
    toolbar.className = 'table-toolbar'
    toolbar.innerHTML = `
      <div class="table-controls">
        <button class="table-btn" onclick="window.addTableRow(this)" title="添加行">
          <span>+行</span>
        </button>
        <button class="table-btn" onclick="window.addTableColumn(this)" title="添加列">
          <span>+列</span>
        </button>
        <button class="table-btn" onclick="window.deleteTableRow(this)" title="删除行">
          <span>-行</span>
        </button>
        <button class="table-btn" onclick="window.deleteTableColumn(this)" title="删除列">
          <span>-列</span>
        </button>
        <button class="table-btn" onclick="window.adjustCellPadding(this, 'left')" title="文字左移">
          <span>←文</span>
        </button>
        <button class="table-btn" onclick="window.adjustCellPadding(this, 'right')" title="文字右移">
          <span>文→</span>
        </button>
      </div>
    `
    return toolbar
  }

  // 显示单元格右键菜单
  const showCellContextMenu = (event) => {
    event.preventDefault()
    ElMessage.info('右键菜单功能开发中，可使用工具栏进行表格操作')
  }

  // 设置全局表格操作函数
  const setupGlobalTableFunctions = () => {
    // 添加表格行
    window.addTableRow = function(button) {
      const table = button.closest('.table-toolbar').nextElementSibling
      if (table && table.tagName === 'TABLE') {
        const tbody = table.querySelector('tbody') || table
        const lastRow = tbody.lastElementChild
        if (lastRow) {
          const newRow = lastRow.cloneNode(true)
          const cells = newRow.querySelectorAll('td, th')
          cells.forEach(cell => {
            cell.innerHTML = '&nbsp;'
          })
          tbody.appendChild(newRow)
          
          // 触发内容变化
          const editorElement = document.getElementById('word-editor')
          if (editorElement) {
            editorElement.dispatchEvent(new Event('input', { bubbles: true }))
          }
        }
      }
    }

    // 添加表格列
    window.addTableColumn = function(button) {
      const table = button.closest('.table-toolbar').nextElementSibling
      if (table && table.tagName === 'TABLE') {
        const rows = table.querySelectorAll('tr')
        rows.forEach(row => {
          const lastCell = row.lastElementChild
          if (lastCell) {
            const newCell = document.createElement(lastCell.tagName.toLowerCase())
            newCell.innerHTML = '&nbsp;'
            newCell.style.cssText = lastCell.style.cssText
            row.appendChild(newCell)
          }
        })
        
        // 触发内容变化
        const editorElement = document.getElementById('word-editor')
        if (editorElement) {
          editorElement.dispatchEvent(new Event('input', { bubbles: true }))
        }
      }
    }

    // 删除表格行
    window.deleteTableRow = function(button) {
      const table = button.closest('.table-toolbar').nextElementSibling
      if (table && table.tagName === 'TABLE') {
        const tbody = table.querySelector('tbody') || table
        const rows = tbody.querySelectorAll('tr')
        if (rows.length > 1) {
          tbody.removeChild(rows[rows.length - 1])
          
          // 触发内容变化
          const editorElement = document.getElementById('word-editor')
          if (editorElement) {
            editorElement.dispatchEvent(new Event('input', { bubbles: true }))
          }
        }
      }
    }

    // 删除表格列
    window.deleteTableColumn = function(button) {
      const table = button.closest('.table-toolbar').nextElementSibling
      if (table && table.tagName === 'TABLE') {
        const rows = table.querySelectorAll('tr')
        let maxCells = 0
        rows.forEach(row => {
          maxCells = Math.max(maxCells, row.children.length)
        })
        
        if (maxCells > 1) {
          rows.forEach(row => {
            if (row.children.length > 0) {
              row.removeChild(row.lastElementChild)
            }
          })
          
          // 触发内容变化
          const editorElement = document.getElementById('word-editor')
          if (editorElement) {
            editorElement.dispatchEvent(new Event('input', { bubbles: true }))
          }
        }
      }
    }

    // 调整单元格内文字位置
    window.adjustCellPadding = function(button, direction) {
      const table = button.closest('.table-toolbar').nextElementSibling
      if (table && table.tagName === 'TABLE') {
        // 获取当前选中的单元格或第一个单元格
        let targetCell = null
        const selection = window.getSelection()
        
        if (selection.rangeCount > 0) {
          const range = selection.getRangeAt(0)
          let node = range.startContainer
          
          // 向上查找到td或th元素
          while (node && node.nodeType !== 1) {
            node = node.parentNode
          }
          while (node && !['TD', 'TH'].includes(node.tagName)) {
            node = node.parentNode
          }
          
          if (node && ['TD', 'TH'].includes(node.tagName)) {
            targetCell = node
          }
        }
        
        // 如果没有选中单元格，使用第一个单元格
        if (!targetCell) {
          targetCell = table.querySelector('td, th')
        }
        
        if (targetCell) {
          const currentStyle = window.getComputedStyle(targetCell)
          let currentPadding = parseInt(currentStyle.paddingLeft) || 5
          
          if (direction === 'left') {
            // 减少左padding，文字往左移
            currentPadding = Math.max(0, currentPadding - 3)
          } else if (direction === 'right') {
            // 增加左padding，文字往右移
            currentPadding = currentPadding + 3
          }
          
          targetCell.style.paddingLeft = currentPadding + 'px'
          
          // 触发内容变化
          const editorElement = document.getElementById('word-editor')
          if (editorElement) {
            editorElement.dispatchEvent(new Event('input', { bubbles: true }))
          }
          
          ElMessage.success(`单元格文字${direction === 'left' ? '左移' : '右移'}成功`)
        }
      }
    }
  }

  return {
    enhanceTableEditing,
    setupGlobalTableFunctions
  }
}