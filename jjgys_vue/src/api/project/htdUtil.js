import request from '@/utils/request'
import qs from 'qs'
const api_name = '/jjg/fbgc/generate/table'

export default {

   
// 生成评定表
scpdb(params){
  //params=qs.stringify(params)
  return request({
    url: `${api_name}/generatePdb`,
    method: 'post',
    responseType: 'blob',
    contentType: "application/json",
    data: params// 使用blob下载
  })
},
// 下载评定表

downloadpdb(params){
  //let params={proname:proname,htds:htd}
  //params=qs.stringify(params, { indices: false })
  return request({
    url: `${api_name}/downloadpdb`,
    method: 'post',
    responseType: 'blob',
    data:params
   
    
  })
},
// 生成建设项目质量评定表
scjszlPdb(params){
    return request({
      url: `${api_name}/generateJSZLPdb`,
      method: 'post',
      responseType: 'blob',
      data: params// 使用blob下载
    })
  },
// 下载建设项目质量评定表
downloadjsxm(proname,htd){
  return request({
    url: `${api_name}/downloadjsxm?proname=`+proname,
    method: 'get',
    responseType: 'blob',
    
    
  })
},
// 报告中表格
scbgzBg(params){
    return request({
      url: `${api_name}/generateBGZBG`,
      method: 'post',
      responseType: 'blob',
      data: params// 使用blob下载
    })
  },
// 下载报告中表格
downloadbgzbg(proname,htd){
  return request({
    url: `${api_name}/downloadbgzbg?proname=`+proname,
    method: 'get',
    responseType: 'blob',
    
    
  })
},
// 生成报告

scbg(params){
  return request({
    url: `/jjg/fbgc/generate/word/generateword?proname=`+params,
    method: 'post',
    responseType: 'blob',
    //data: params// 使用blob下载
  })
},
// 下载报告
downloadbg(proname){
  return request({
    url: `/jjg/fbgc/generate/word/download?proname=`+proname,
    method: 'get',
    responseType: 'blob',
  })
},

    
    
  



}