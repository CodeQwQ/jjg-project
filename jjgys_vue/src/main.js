import Vue from 'vue'

import 'normalize.css/normalize.css' // A modern alternative to CSS resets
import "@/styles/all.css"
import ElementUI from 'element-ui'
import  Loading from 'element-ui';
import dataV from '@jiaminghi/data-view'
Vue.use(dataV)
import { borderBox1 } from '@jiaminghi/data-view'

Vue.use(borderBox1)

import * as echarts from 'echarts'
 Vue.prototype.$echarts = echarts
 import VScaleScreen from 'v-scale-screen'
 Vue.use(VScaleScreen)
import 'element-ui/lib/theme-chalk/index.css'
import locale from 'element-ui/lib/locale/lang/zh-CN' // lang i18n

import '@/styles/index.scss' // global css

import App from './App'
import store from './store'
import router from './router'

import '@/icons' // icon
import '@/permission' // permission control
// 引入Viewer插件
import VueViewer, { directive as viewerDirective } from 'v-viewer';
// 引入Viewer插件的图片预览器的样式
import 'viewerjs/dist/viewer.css'; 
// 使用Viewer图片预览器
Vue.use(VueViewer,{
  defaultOptions: {
    zIndex: 9999,
　　} 
})
// 用于图片预览的指令方式调用 在元素上加上会处理元素下所有的图片,为图片添加点击事件,点击即可预览
Vue.directive('viewer', viewerDirective({
  debug: false,

}));






Vue.prototype.$reset = function (formRef, ...excludeFields) {//弹窗关闭清空
  this.$refs[formRef].resetFields();
  console.log('清空',formRef)
  
};
//新增
import hasBtnPermission from '@/utils/btn-permission'
Vue.prototype.$hasBP = hasBtnPermission

/**
 * If you don't want to use mock-server
 * you want to use MockJs for mock api
 * you can execute: mockXHR()
 *
 * Currently MockJs will be used in the production environment,
 * please remove it before going online ! ! !
 */
if (process.env.NODE_ENV === 'production') {
  const { mockXHR } = require('../mock')
  mockXHR()
}

// set ElementUI lang to EN
Vue.use(ElementUI, { locale })
Vue.use(dataV)
// 如果想要中文版 element-ui，按如下方式声明
// Vue.use(ElementUI)


Vue.config.productionTip = false

var app=new Vue({
  el: '#app',
  router,
  store,
  render: h => h(App)
})
const errorHandler=(err,vm)=>{
  console.log('sssss1',err,vm);
  router.push({path:'/404',query:{error:err}})
 
}
Vue.config.errorHandler = errorHandler;
Vue.prototype.$throw = (error)=> errorHandler(error,this);
