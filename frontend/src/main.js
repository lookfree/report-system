import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import * as ElementPlusIconsVue from '@element-plus/icons-vue'
import zhCn from 'element-plus/dist/locale/zh-cn.mjs'

// Import jspreadsheet CSS
import 'jspreadsheet-ce/dist/jspreadsheet.css'
// Import jsuites CSS for UI components  
import 'jsuites/dist/jsuites.css'

// 尝试全局引入jspreadsheet
import jspreadsheet from 'jspreadsheet-ce'
console.log('Global jspreadsheet import:', jspreadsheet)
if (typeof window !== 'undefined') {
  window.jspreadsheet = jspreadsheet
  console.log('Set window.jspreadsheet:', window.jspreadsheet)
}

import App from './App.vue'
import router from './router'

const app = createApp(App)

// Register all icons
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}

app.use(ElementPlus, {
  locale: zhCn
})
app.use(router)
app.mount('#app')