import request from '@/utils/request'
import qs from 'qs'

const api_name = '/project/info/htd'

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
  // 添加合同段
  save(proHtd) {
    return request({
      url: `${api_name}/save`,
      method: 'post',
      data: proHtd
    })
  },
  // 检测合同段
  checkProname(params){
    return request({
      url: `${api_name}/checkHtd`,
      method: 'post',
      data:params
      
    })
  },
  // 导出
 exporthtd(proname,htds){
  let params={proname:proname,htd:htds}
  params=qs.stringify(params, { indices: false })
  
  console.log("paaa",params)
  return request({
    url: `${api_name}/exporthtd?`+params,
    method: 'get',
    
    responseType: 'blob', // 使用blob下载
  })
},
// 导入
importhtd(params){
  
  return request({
    url: `${api_name}/importhtd`,
    method: 'post',
    data:params, // 使用blob下载
    
  })
},
// 生成鉴定表
scjdb(params){
  params=qs.stringify(params,{indices:false})
  return request({
    url: `${api_name}/generateJdb`,
    method: 'post',
    responseType: 'blob',
    data: params// 使用blob下载
  })
},
// 下载鉴定表
download(proname,htd,userid){
  let params={proname:proname,htds:htd,userid:userid}
  params=qs.stringify(params, { indices: false })
  return request({
    url: `${api_name}/download?`+params,
    method: 'get',
    responseType: 'blob',
    
    
  })
},
// 通过id查询
getById(id){
  return request({
    url: `${api_name}/gethtd/`+id,
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