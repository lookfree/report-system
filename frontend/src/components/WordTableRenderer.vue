<template>
  <div class="word-document-container">
    <!-- 文档标题 -->
    <div class="document-header" v-if="title">
      <h1 class="main-title">{{ title }}</h1>
      <div class="sub-title" v-if="subTitle">{{ subTitle }}</div>
      <div class="report-date" v-if="reportDate">（{{ reportDate }}）</div>
    </div>
    
    <!-- 文档内容 -->
    <div class="document-content" v-html="renderedHtml"></div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'

const props = defineProps({
  htmlContent: {
    type: String,
    default: ''
  },
  title: {
    type: String,
    default: ''
  },
  subTitle: {
    type: String,
    default: ''
  },
  reportDate: {
    type: String,
    default: ''
  }
})

// 渲染增强后的HTML
const renderedHtml = computed(() => {
  return enhanceWordHtml(props.htmlContent)
})

// 增强HTML以更好地模拟Word样式
function enhanceWordHtml(html) {
  if (!html) return ''
  
  let enhanced = html
  
  // 特殊处理"审计概述"这类两列表格
  enhanced = enhanced.replace(/<table([^>]*)>/gi, (match, attrs) => {
    return `<table class="word-style-table"${attrs}>`
  })
  
  // 处理表格行和单元格
  enhanced = enhanced.replace(/<tr>/gi, '<tr class="word-table-row">')
  enhanced = enhanced.replace(/<td/gi, '<td class="word-table-cell"')
  enhanced = enhanced.replace(/<th/gi, '<th class="word-table-header"')
  
  // 处理第一列（通常是标签列）
  enhanced = enhanced.replace(/<tr([^>]*)>\s*<td([^>]*)>/gi, 
    '<tr$1><td class="word-table-cell word-table-label"$2>')
  
  return enhanced
}

onMounted(() => {
  // 可以在这里添加额外的DOM操作
  processSpecialTables()
})

// 处理特殊表格格式
function processSpecialTables() {
  const tables = document.querySelectorAll('.word-style-table')
  tables.forEach(table => {
    const rows = table.querySelectorAll('tr')
    rows.forEach(row => {
      const cells = row.querySelectorAll('td, th')
      if (cells.length === 2) {
        // 两列表格特殊处理
        cells[0].classList.add('label-column')
        cells[1].classList.add('content-column')
      }
    })
  })
}
</script>

<style lang="scss">
.word-document-container {
  max-width: 21cm;
  margin: 0 auto;
  padding: 2.54cm;
  background: white;
  min-height: 29.7cm;
  box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
  /* 字体大小由Word模板数据决定 */
  line-height: 1.8;
  color: #000;
  
  // 文档标题样式
  .document-header {
    text-align: center;
    margin-bottom: 40px;
    
    .main-title {
      font-size: 22pt;
      font-weight: bold;
      font-family: '黑体', SimHei, sans-serif;
      letter-spacing: 3px;
      margin: 0 0 20px 0;
      line-height: 1.5;
    }
    
    .sub-title {
      font-size: 18pt;
      font-weight: bold;
      font-family: '黑体', SimHei, sans-serif;
      letter-spacing: 2px;
      margin: 0 0 15px 0;
    }
    
    .report-date {
      font-size: 14pt;
      font-family: '宋体', SimSun, serif;
      margin: 10px 0;
    }
  }
  
  // 文档内容样式
  .document-content {
    // Word风格表格
    .word-style-table, table {
      width: 100%;
      border-collapse: collapse;
      border: 1px solid #000;
      margin: 20px 0;
      table-layout: fixed;
      /* 字体大小由Word模板数据决定 */
      
      // 所有单元格基础样式
      td, th, .word-table-cell, .word-table-header {
        border: 1px solid #000;
        padding: 10px 12px;
        vertical-align: middle;
        line-height: 1.6;
        text-align: left;
        word-wrap: break-word;
        
        // 处理单元格内的段落
        p {
          margin: 0;
          text-indent: 0;
        }
        
        // 处理列表
        ul, ol {
          margin: 5px 0;
          padding-left: 20px;
          
          li {
            margin: 3px 0;
            line-height: 1.6;
          }
        }
      }
      
      // 标签列（第一列）特殊样式
      .word-table-label,
      .label-column,
      td:first-child,
      th:first-child {
        width: 100px;
        min-width: 100px;
        max-width: 150px;
        font-weight: normal;
        background-color: transparent;
        text-align: left;
        padding-left: 15px;
      }
      
      // 内容列（第二列及后续）
      .content-column {
        text-align: left;
        padding-left: 15px;
      }
      
      // 模拟Word的"审计概述"表格样式
      &.audit-overview-table {
        margin: 30px 0;
        
        tr {
          &:nth-child(odd) {
            background-color: transparent;
          }
          
          td:first-child {
            width: 120px;
            font-weight: normal;
            white-space: nowrap;
          }
          
          td:last-child {
            padding-left: 20px;
            line-height: 1.8;
          }
        }
      }
      
      // 处理合并单元格
      td[colspan], th[colspan] {
        text-align: center;
        font-weight: bold;
      }
      
      td[rowspan], th[rowspan] {
        vertical-align: middle;
      }
    }
    
    // 段落样式
    p {
      margin: 12px 0;
      text-indent: 2em;
      text-align: justify;
      line-height: 1.8;
      /* 字体大小由Word模板数据决定 */
      
      &.no-indent {
        text-indent: 0;
      }
      
      &.center {
        text-align: center;
        text-indent: 0;
      }
    }
    
    // 标题样式
    h1 {
      font-size: 22pt;
      font-weight: bold;
      text-align: center;
      margin: 30px 0;
      font-family: '黑体', SimHei, sans-serif;
      letter-spacing: 2px;
    }
    
    h2 {
      font-size: 16pt;
      font-weight: bold;
      margin: 20px 0 15px 0;
      font-family: '黑体', SimHei, sans-serif;
    }
    
    h3 {
      font-size: 14pt;
      font-weight: bold;
      margin: 15px 0 10px 0;
      font-family: '宋体', SimSun, serif;
    }
    
    // 列表样式
    ul, ol {
      margin: 10px 0;
      padding-left: 30px;
      
      li {
        margin: 5px 0;
        line-height: 1.8;
        text-align: justify;
      }
    }
    
    // 编号列表特殊样式
    ol {
      counter-reset: item;
      padding-left: 0;
      
      li {
        display: block;
        margin-left: 30px;
        
        &:before {
          content: counter(item) ". ";
          counter-increment: item;
          font-weight: normal;
          margin-left: -25px;
          display: inline-block;
          width: 25px;
        }
      }
    }
    
    // 特殊格式
    .bullet-list {
      list-style: none;
      padding-left: 0;
      
      li {
        position: relative;
        padding-left: 25px;
        
        &:before {
          content: "●";
          position: absolute;
          left: 0;
        }
      }
    }
  }
}

// 打印样式
@media print {
  .word-document-container {
    box-shadow: none;
    padding: 0;
    
    .document-content {
      page-break-inside: avoid;
      
      table {
        page-break-inside: auto;
        
        tr {
          page-break-inside: avoid;
        }
      }
    }
  }
}

// A4纸张模拟
@page {
  size: A4;
  margin: 2.54cm;
}
</style>