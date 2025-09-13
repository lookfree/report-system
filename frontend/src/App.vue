<template>
  <el-container class="app-container">
    <el-header>
      <div class="header-content">
        <h1>报告管理系统</h1>
        <el-menu mode="horizontal" :default-active="activeMenu" router>
          <el-menu-item index="/templates">
            <el-icon><Document /></el-icon>
            <span>模板管理</span>
          </el-menu-item>
          <el-menu-item index="/datasources">
            <el-icon><Connection /></el-icon>
            <span>数据源</span>
          </el-menu-item>
          <el-menu-item index="/tasks">
            <el-icon><Clock /></el-icon>
            <span>报告任务</span>
          </el-menu-item>
          <el-menu-item index="/reports">
            <el-icon><Files /></el-icon>
            <span>报告结果</span>
          </el-menu-item>
        </el-menu>
      </div>
    </el-header>
    <el-main>
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
import { computed } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()
const activeMenu = computed(() => '/' + route.path.split('/')[1])
</script>

<style lang="scss">
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

.app-container {
  min-height: 100vh;
  background: #f5f7fa;
  
  .el-header {
    background: #fff;
    box-shadow: 0 1px 4px rgba(0, 21, 41, 0.08);
    padding: 0;
    border-bottom: 1px solid #e8eaec;
    
    .header-content {
      display: flex;
      align-items: center;
      padding: 0 24px;
      height: 100%;
      
      h1 {
        margin-right: 48px;
        font-size: 20px;
        color: #1f2329;
        font-weight: 500;
      }
      
      .el-menu {
        flex: 1;
        border: none;
        background: transparent;
        
        .el-menu-item {
          color: #4e5969;
          font-weight: 400;
          border-bottom: 2px solid transparent;
          margin: 0 4px;
          
          &:hover {
            background: #f2f3f5;
            color: #1677ff;
          }
          
          &.is-active {
            color: #1677ff;
            background: transparent;
            border-bottom-color: #1677ff;
          }
          
          .el-icon {
            margin-right: 8px;
            font-size: 16px;
          }
        }
      }
    }
  }
  
  .el-main {
    background: #f5f7fa;
    padding: 24px;
  }
}
</style>