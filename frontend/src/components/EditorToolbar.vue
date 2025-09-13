<template>
  <div class="editor-toolbar">
    <div class="toolbar-group">
      <el-button type="primary" @click="$emit('save')">
        <el-icon><DocumentChecked /></el-icon>
        保存模板
      </el-button>
      <el-button @click="$emit('preview')">
        <el-icon><View /></el-icon>
        预览
      </el-button>
    </div>
    
    <el-divider direction="vertical" />
    
    <!-- 格式化工具组 -->
    <div class="toolbar-group">
      <el-tooltip content="粗体" placement="bottom">
        <el-button size="small" @click="$emit('format-text', 'bold')">
          <strong>B</strong>
        </el-button>
      </el-tooltip>
      <el-tooltip content="斜体" placement="bottom">
        <el-button size="small" @click="$emit('format-text', 'italic')">
          <em>I</em>
        </el-button>
      </el-tooltip>
      <el-tooltip content="下划线" placement="bottom">
        <el-button size="small" @click="$emit('format-text', 'underline')">
          <u>U</u>
        </el-button>
      </el-tooltip>
    </div>
    
    <el-divider direction="vertical" />
    
    <!-- 对齐工具组 -->
    <div class="toolbar-group">
      <el-tooltip content="左对齐" placement="bottom">
        <el-button size="small" @click="$emit('align-text', 'left')">
          <el-icon><Operation /></el-icon>
        </el-button>
      </el-tooltip>
      <el-tooltip content="居中对齐" placement="bottom">
        <el-button size="small" @click="$emit('align-text', 'center')">
          <el-icon><Aim /></el-icon>
        </el-button>
      </el-tooltip>
      <el-tooltip content="右对齐" placement="bottom">
        <el-button size="small" @click="$emit('align-text', 'right')">
          <el-icon><Right /></el-icon>
        </el-button>
      </el-tooltip>
      <el-tooltip content="两端对齐" placement="bottom">
        <el-button size="small" @click="$emit('align-text', 'justify')">
          <el-icon><Grid /></el-icon>
        </el-button>
      </el-tooltip>
    </div>
    
    <el-divider direction="vertical" />
    
    <!-- 缩进工具组 -->
    <div class="toolbar-group">
      <el-tooltip content="减少缩进" placement="bottom">
        <el-button size="small" @click="$emit('adjust-indent', 'decrease')">
          <el-icon><Back /></el-icon>
        </el-button>
      </el-tooltip>
      <el-tooltip content="增加缩进" placement="bottom">
        <el-button size="small" @click="$emit('adjust-indent', 'increase')">
          <el-icon><Right /></el-icon>
        </el-button>
      </el-tooltip>
    </div>
    
    <el-divider direction="vertical" />
    
    <!-- 字体大小控制组 -->
    <div class="toolbar-group font-size-group">
      <el-tooltip content="减小字体" placement="bottom">
        <el-button size="small" @click="$emit('adjust-font-size', 'decrease')">
          A⁻
        </el-button>
      </el-tooltip>
      <el-select 
        :model-value="currentFontSize" 
        @update:model-value="handleFontSizeChange"
        @mousedown="handleSelectMousedown"
        size="small" 
        style="width: 80px;"
        placeholder="字号"
        class="font-size-select"
      >
        <el-option label="9pt" value="9pt" />
        <el-option label="10pt" value="10pt" />
        <el-option label="10.5pt" value="10.5pt" />
        <el-option label="11pt" value="11pt" />
        <el-option label="12pt" value="12pt" />
        <el-option label="14pt" value="14pt" />
        <el-option label="16pt" value="16pt" />
        <el-option label="18pt" value="18pt" />
        <el-option label="20pt" value="20pt" />
        <el-option label="24pt" value="24pt" />
      </el-select>
      <el-tooltip content="增大字体" placement="bottom">
        <el-button size="small" @click="$emit('adjust-font-size', 'increase')">
          A⁺
        </el-button>
      </el-tooltip>
    </div>
    
    <el-divider direction="vertical" />
    
    <div class="toolbar-group">
      <el-button @click="$emit('insert-data')">
        <el-icon><Connection /></el-icon>
        插入数据
      </el-button>
    </div>
    
    <el-divider direction="vertical" />
    
    <div class="toolbar-group">
      <el-button @click="$emit('export-word')">
        <el-icon><Download /></el-icon>
        导出Word
      </el-button>
    </div>
  </div>
</template>

<script setup>
import { 
  DocumentChecked, 
  View, 
  Grid, 
  Connection, 
  Download,
  Operation,
  Aim,
  Right,
  Back
} from '@element-plus/icons-vue'

defineProps({
  currentFontSize: {
    type: String,
    default: '12pt'
  }
})

const emit = defineEmits([
  'save',
  'preview', 
  'format-text',
  'align-text',
  'adjust-indent',
  'change-font-size',
  'adjust-font-size',
  'insert-data',
  'export-word',
  'save-selection'
])

// 处理字体大小选择器的鼠标按下事件
const handleSelectMousedown = () => {
  // 先保存当前选区
  emit('save-selection')
}

// 处理字体大小变化
const handleFontSizeChange = (size) => {
  emit('change-font-size', size)
}
</script>

<style scoped lang="scss">
.editor-toolbar {
  background: white;
  padding: 12px 20px;
  border-bottom: 1px solid #e0e0e0;
  display: flex;
  align-items: center;
  gap: 10px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
  flex-wrap: wrap;
  
  .toolbar-group {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  
  .el-button {
    border-radius: 4px;
    
    &.is-plain {
      background: #f9f9f9;
    }
    
    &:hover {
      background-color: #e6f7ff;
      border-color: #91d5ff;
    }
  }
  
  .el-divider--vertical {
    height: 24px;
    margin: 0 8px;
  }
  
  .font-size-group {
    .el-button {
      font-weight: bold;
      font-size: 12px;
      min-width: 32px;
      
      &:first-child {
        border-radius: 4px 0 0 4px;
      }
      
      &:last-child {
        border-radius: 0 4px 4px 0;
      }
    }
    
    :deep(.el-select) {
      margin: 0 -1px;
      
      .el-input {
        .el-input__wrapper {
          border-radius: 0;
          border-left: none;
          border-right: none;
        }
      }
    }
  }
}
</style>