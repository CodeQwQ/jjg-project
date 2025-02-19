import request from '@/utils/request'

const api_name = '/jjg/fbgc/jagc'

export default {
  
  // 导出模板
  exportold(){
    return request({
      url: `${api_name}/exportold`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入文件
  importold(params){
    return request({
      url: `${api_name}/importold`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  }
  
  
}