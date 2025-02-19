import request from '@/utils/request'
import qs from 'qs'

const api_name = '/jjg/project/plan'

export default {

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
  // 查看项目进度
  ckxmjd(params) {
    return request({
      url: `${api_name}/lookplan`,
      method: 'post',
      data: params
    })
  },
  // 查看合同段信息
  ckhtd(params) {
    return request({
      url: `${api_name}/gethtd`,
      method: 'post',
      data: params
    })
  },
  // 导出
 exportxmjd(params){
  
  return request({
    url: `${api_name}/exportxmjd/`+params,
    method: 'get',
    
    responseType: 'blob', // 使用blob下载
  })
},
// 导入
importxmjd(params){
  
  return request({
    url: `${api_name}/importxmjd`,
    method: 'post',
    data:params, // 使用blob下载
    
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