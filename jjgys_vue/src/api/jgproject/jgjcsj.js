import request from '@/utils/request'
import qs from 'qs'
const api_name = '/jjg/fbgc/jgjcsj'

export default {
  // 生成旧规范评定表
  scpdbold(proname){
    
    return request({
      url: `${api_name}/generatepdb?proname=`+proname,
      method: 'post',
      responseType: 'blob',
    })
  },
  // 生成新规范评定表
  scpdbnew(proname){
    return request({
      url: `${api_name}/generatepdbNew?proname=`+proname,
      method: 'post',
      responseType: 'blob',
      
    })
  },
  // 下载评定表
  download(proname){
    return request({
      url: `${api_name}/download?proname=`+proname,
      method: 'get',
      responseType: 'blob',
      
      
    })
  },
  // 导出模板
  exportjgsj(proname){
    return request({
      url: `${api_name}/exportjgjcdata?proname=`+proname,
      method: 'get',
      responseType: 'blob',
       // 使用blob下载
    })
  },
  // 导入文件
  importjgsj(params){
    return request({
      url: `${api_name}/importjgsj`,
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
  },
  // 通过id查询
getById(id){
  return request({
    url: `${api_name}/getjcsj/`+id,
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
// 生成报告

scbg(params){
  return request({
    url: `/jjg/fbgc/jgjcsj/generateword?proname=`+params,
    method: 'post',
    responseType: 'blob',
    //data: params// 使用blob下载
  })
},
// 下载报告
downloadbg(proname){
  return request({
    url: `/jjg/fbgc/jgjcsj/downloadjgbg?proname=`+proname,
    method: 'get',
    responseType: 'blob',
  })
},

}