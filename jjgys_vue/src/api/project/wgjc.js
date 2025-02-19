import request from '@/utils/request'

const api_name = '/jjg/wgjc'

export default {
  
  // 导出模板文件
  exportwgjc(params){
    return request({
      url: `${api_name}/export?projectname=`+params,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入文件
  importwgjc(params){
    return request({
      url: `${api_name}/importwgjc`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  },
  
  // 通过id查询
  getById(id){
    return request({
      url: `${api_name}/getJabx/`+id,
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