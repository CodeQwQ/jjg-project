import request from '@/utils/request'

const api_name = '/jjg/fbgc/jagc'

export default {
  
  // 导出模板
  exportnew(){
    return request({
      url: `${api_name}/exportnew`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入文件
  importnew(params){
    return request({
      url: `${api_name}/importnew`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  }
  
  
}