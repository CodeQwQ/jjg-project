import request from '@/utils/request'

const api_name = '/jjg/fbgc/sdgc'

export default {
  // 生成鉴定表
  scjdb(params){
    return request({
      url: `${api_name}/generateJdb`,
      method: 'post',
      responseType: 'blob',
      data: params// 使用blob下载
    })
  },
  // 下载鉴定表
  download(proname,htd){
    return request({
      url: `${api_name}/download?proname=`+proname+'&htd='+htd,
      method: 'get',
      responseType: 'blob',
      
      
    })
  },
  // 导出
 exportsdgc(){
    return request({
      url: `${api_name}/exportsdgc`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入
  importsdgc(params){
    return request({
      url: `${api_name}/importsdgc`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  },
  

}