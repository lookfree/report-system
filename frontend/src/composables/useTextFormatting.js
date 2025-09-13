import { ref } from 'vue'
import { ElMessage } from 'element-plus'

export function useTextFormatting() {
  const currentFontSize = ref('12pt')
  let savedSelection = null
  
  // 保存当前选区
  const saveSelection = () => {
    const selection = window.getSelection()
    if (selection.rangeCount > 0) {
      savedSelection = selection.getRangeAt(0).cloneRange()
    }
  }
  
  // 恢复选区
  const restoreSelection = () => {
    if (savedSelection) {
      const selection = window.getSelection()
      selection.removeAllRanges()
      selection.addRange(savedSelection)
      return true
    }
    return false
  }
  
  // 格式化文本函数
  const formatText = (command) => {
    document.execCommand(command, false, null)
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      // 触发内容变化事件
      editorElement.dispatchEvent(new Event('input', { bubbles: true }))
    }
  }

  // 文本对齐函数
  const alignText = (alignment) => {
    let command = ''
    switch (alignment) {
      case 'left':
        command = 'justifyLeft'
        break
      case 'center':
        command = 'justifyCenter'
        break
      case 'right':
        command = 'justifyRight'
        break
      case 'justify':
        command = 'justifyFull'
        break
    }
    
    if (command) {
      document.execCommand(command, false, null)
      const editorElement = document.getElementById('word-editor')
      if (editorElement) {
        editorElement.dispatchEvent(new Event('input', { bubbles: true }))
      }
    }
  }

  // 调整缩进函数
  const adjustIndent = (direction) => {
    const command = direction === 'increase' ? 'indent' : 'outdent'
    document.execCommand(command, false, null)
    const editorElement = document.getElementById('word-editor')
    if (editorElement) {
      editorElement.dispatchEvent(new Event('input', { bubbles: true }))
    }
  }

  // 更改字体大小函数
  const changeFontSize = (size) => {
    let range = null
    
    // 首先尝试恢复保存的选区
    if (restoreSelection()) {
      const selection = window.getSelection()
      if (selection.rangeCount > 0) {
        range = selection.getRangeAt(0)
      }
    } else {
      // 如果没有保存的选区，使用当前选区
      const selection = window.getSelection()
      if (selection.rangeCount > 0) {
        range = selection.getRangeAt(0)
      }
    }
    
    if (!range || range.collapsed) {
      ElMessage.warning('请先选中要修改的文字')
      return
    }
    
    try {
      const selectedText = range.toString()
      if (!selectedText.trim()) {
        ElMessage.warning('请先选中要修改的文字')
        return
      }
      
      const span = document.createElement('span')
      span.style.fontSize = size
      span.style.fontFamily = 'inherit'
      
      try {
        range.surroundContents(span)
      } catch (e) {
        const contents = range.extractContents()
        span.appendChild(contents)
        range.insertNode(span)
      }
      
      // 清除选区并设置光标到修改后的文本之后
      const selection = window.getSelection()
      selection.removeAllRanges()
      const newRange = document.createRange()
      newRange.setStartAfter(span)
      newRange.collapse(true)
      selection.addRange(newRange)
      
      const editorElement = document.getElementById('word-editor')
      if (editorElement) {
        editorElement.dispatchEvent(new Event('input', { bubbles: true }))
      }
      
      // 清除保存的选区
      savedSelection = null
      ElMessage.success(`字体大小已设置为 ${size}`)
      
    } catch (error) {
      console.error('设置字体大小失败:', error)
      ElMessage.error('设置字体大小失败，请重试')
    }
  }

  // 字体大小增减函数
  const adjustFontSize = (direction) => {
    const selection = window.getSelection()
    if (selection.rangeCount === 0 || selection.getRangeAt(0).collapsed) {
      ElMessage.warning('请先选中要修改的文字')
      return
    }
    
    const fontSizes = ['9pt', '10pt', '10.5pt', '11pt', '12pt', '14pt', '16pt', '18pt', '20pt', '24pt']
    let newSize = '12pt'
    
    const range = selection.getRangeAt(0)
    const parentElement = range.commonAncestorContainer.nodeType === Node.TEXT_NODE 
      ? range.commonAncestorContainer.parentElement 
      : range.commonAncestorContainer
    
    const currentStyle = window.getComputedStyle(parentElement)
    const currentFontSize = currentStyle.fontSize
    
    const pxToPt = (px) => {
      const pxValue = parseFloat(px)
      const ptValue = Math.round(pxValue * 0.75)
      return `${ptValue}pt`
    }
    
    let currentPt = currentFontSize.includes('px') ? pxToPt(currentFontSize) : '12pt'
    
    let currentIndex = fontSizes.indexOf(currentPt)
    if (currentIndex === -1) {
      const currentNum = parseFloat(currentPt)
      currentIndex = fontSizes.findIndex(size => parseFloat(size) >= currentNum)
      if (currentIndex === -1) currentIndex = fontSizes.length - 1
    }
    
    if (direction === 'increase' && currentIndex < fontSizes.length - 1) {
      newSize = fontSizes[currentIndex + 1]
    } else if (direction === 'decrease' && currentIndex > 0) {
      newSize = fontSizes[currentIndex - 1]
    } else {
      ElMessage.info(direction === 'increase' ? '字体已是最大' : '字体已是最小')
      return
    }
    
    currentFontSize.value = newSize
    changeFontSize(newSize)
  }

  return {
    currentFontSize,
    formatText,
    alignText,
    adjustIndent,
    changeFontSize,
    adjustFontSize,
    saveSelection,
    restoreSelection
  }
}