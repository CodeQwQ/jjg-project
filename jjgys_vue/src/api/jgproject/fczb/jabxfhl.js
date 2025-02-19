import request from '@/utils/request'

const api_name = '/jjg/fbgc/jtaqss/jabxfhljgfc'

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
  // 导出模板
  exportjabxfhl(){
    return request({
      url: `${api_name}/exportjabxfhl`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入文件
  importjabxfhl(params){
    return request({
      url: `${api_name}/importjabxfhl`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  },
  // 通过id查询
  getById(id){
    return request({
      url: `${api_name}/getJabxfhl/`+id,
      method: 'get',
      
    })
  },
   // 修改
   modify(data) {
    return request({
      url: `${api_name}/update`,
      method: `post`,
      data: data
    })
  },
  // 查看鉴定表
  lookjdb(params){
    return request({
      url: `${api_name}/lookJdbjg`,
      method: 'post',
      data:params, 
      
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