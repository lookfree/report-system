<template>
  <div class="word-cover-container">
    <div class="cover-page">
      <!-- 页面边框装饰 -->
      <div class="page-border">
        <!-- 左上角装饰 -->
        <div class="corner-decoration top-left">
          <div class="corner-line horizontal"></div>
          <div class="corner-line vertical"></div>
          <div class="corner-dots">
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
          </div>
        </div>
        
        <!-- 右上角装饰 -->
        <div class="corner-decoration top-right">
          <div class="corner-line horizontal"></div>
          <div class="corner-line vertical"></div>
        </div>
        
        <!-- 左下角装饰 -->
        <div class="corner-decoration bottom-left">
          <div class="corner-line horizontal"></div>
          <div class="corner-line vertical"></div>
          <div class="corner-dots">
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
            <span class="dot"></span>
          </div>
        </div>
        
        <!-- 右下角装饰 -->
        <div class="corner-decoration bottom-right">
          <div class="corner-line horizontal"></div>
          <div class="corner-line vertical"></div>
        </div>
        
        <!-- 封面内容 -->
        <div class="cover-content">
          <div class="title-group">
            <h1 class="main-title">{{ mainTitle || '省公司 API 接口' }}</h1>
            <h1 class="sub-title">{{ subTitle || '专题审计日报' }}</h1>
          </div>
          
          <div class="date-section">
            <p class="report-date">（{{ reportDate || '2025年5月28日' }}）</p>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { defineProps } from 'vue'

const props = defineProps({
  mainTitle: {
    type: String,
    default: '省公司 API 接口'
  },
  subTitle: {
    type: String,
    default: '专题审计日报'
  },
  reportDate: {
    type: String,
    default: '2025年5月28日'
  }
})
</script>

<style lang="scss" scoped>
.word-cover-container {
  width: 100%;
  height: 100%;
  display: flex;
  justify-content: center;
  align-items: center;
  background: #f5f5f5;
  
  .cover-page {
    width: 21cm;
    height: 29.7cm;
    background: white;
    position: relative;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
    display: flex;
    justify-content: center;
    align-items: center;
    
    .page-border {
      width: calc(100% - 3cm);
      height: calc(100% - 3cm);
      position: relative;
      
      // 角落装饰
      .corner-decoration {
        position: absolute;
        width: 60px;
        height: 60px;
        
        &.top-left {
          top: 0;
          left: 0;
          
          .corner-dots {
            position: absolute;
            top: -20px;
            left: -20px;
            display: grid;
            grid-template-columns: repeat(3, 8px);
            grid-template-rows: repeat(3, 8px);
            gap: 4px;
          }
        }
        
        &.top-right {
          top: 0;
          right: 0;
        }
        
        &.bottom-left {
          bottom: 0;
          left: 0;
          
          .corner-dots {
            position: absolute;
            bottom: -20px;
            left: -20px;
            display: grid;
            grid-template-columns: repeat(3, 8px);
            grid-template-rows: repeat(3, 8px);
            gap: 4px;
          }
        }
        
        &.bottom-right {
          bottom: 0;
          right: 0;
        }
        
        .corner-line {
          position: absolute;
          background: #666;
          
          &.horizontal {
            height: 1.5px;
            width: 60px;
          }
          
          &.vertical {
            width: 1.5px;
            height: 60px;
          }
        }
        
        // 定位每个角的线条
        &.top-left {
          .horizontal {
            top: 0;
            left: 0;
          }
          .vertical {
            top: 0;
            left: 0;
          }
        }
        
        &.top-right {
          .horizontal {
            top: 0;
            right: 0;
          }
          .vertical {
            top: 0;
            right: 0;
          }
        }
        
        &.bottom-left {
          .horizontal {
            bottom: 0;
            left: 0;
          }
          .vertical {
            bottom: 0;
            left: 0;
          }
        }
        
        &.bottom-right {
          .horizontal {
            bottom: 0;
            right: 0;
          }
          .vertical {
            bottom: 0;
            right: 0;
          }
        }
        
        // 点阵装饰
        .dot {
          width: 4px;
          height: 4px;
          background: #999;
          border-radius: 50%;
          display: block;
        }
      }
      
      // 封面内容
      .cover-content {
        width: 100%;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        position: relative;
        
        .title-group {
          text-align: center;
          margin-bottom: 80px;
          
          .main-title {
            font-size: 42pt;
            font-weight: bold;
            color: #1e88e5;
            font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;
            letter-spacing: 8px;
            margin: 0 0 30px 0;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
          }
          
          .sub-title {
            font-size: 42pt;
            font-weight: bold;
            color: #1e88e5;
            font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;
            letter-spacing: 8px;
            margin: 0;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
          }
        }
        
        .date-section {
          position: absolute;
          bottom: 150px;
          text-align: center;
          width: 100%;
          
          .report-date {
            font-size: 18pt;
            color: #333;
            font-family: '宋体', SimSun, serif;
            margin: 0;
            letter-spacing: 2px;
          }
        }
      }
    }
  }
}

// 打印样式
@media print {
  .word-cover-container {
    background: white;
    
    .cover-page {
      box-shadow: none;
      width: 100%;
      height: 100%;
      
      .page-border {
        width: calc(100% - 2.54cm);
        height: calc(100% - 2.54cm);
      }
    }
  }
}

// A4纸张设置
@page {
  size: A4;
  margin: 0;
}
</style>