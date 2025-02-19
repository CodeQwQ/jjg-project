import request from '@/utils/request'

const api_name = '/jjg/ljfrnum'

export default {
  export(projectname){
    return request({
      url: `${api_name}/export/`+projectname,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  import(params){
    return request({
      url: `${api_name}/importlj`,
      method: 'post',
      data:params // 使用blob下载
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

    //分页查询
  pageList(current, limit, searchObj) {
    return request({
      url: `${api_name}/findQueryPage/${current}/${limit}`,
      method: 'post',
      data: searchObj
    })
  },
  // 通过id查询
  getById(id){
    return request({
      url: `${api_name}/getxmjd/`+id,
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

}