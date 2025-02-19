import request from '@/utils/request'

const api_name = '/jjg/jgfc/lmgzsdsgpsf'

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
  download(proname){
    return request({
      url: `${api_name}/download?proname=`+proname,
      method: 'get',
      responseType: 'blob',
      
      
    })
  },
  // 导出模板
  exportgzsd(){
    return request({
      url: `${api_name}/exportlmgzsdsgpsf`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入文件
  importgzsd(params){
    return request({
      url: `${api_name}/importlmgzsdsgpsf`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  },
  
  
   // 批量删除
   batchRemove(idList) {
    return request({
      url: `${api_name}/removeBatch`,
      method: `delete`,
      data: idList
    })
  },
  //全部删除
  removeAll(params) {
    return request({
      url: `${api_name}/removeAll`,
      method: `delete`,
      data: params
    })
  },
    //分页查询
  pageList(current, limit, searchObj) {
    return request({
      url: `${api_name}/findQueryPage/${current}/${limit}`,
      method: 'post',
      data: searchObj
    })
  }

}