import request from '@/utils/request'
import qs from 'qs'

const api_name = '/jjg/file/info'

export default {

   
  // 文件列表
 getfilelist(userid){
  return request({
    url: `${api_name}/filelist/${userid}`,
    method: 'get',
  })
},


// 下载
download(params){
  
  return request({
    url: `${api_name}/download`,
    method: 'post',
    responseType: 'blob',
    data:params
    
    
  })
},
// 上传
upload(params){
  return request({
    url: `${api_name}/upload`,
    method: 'post',
    data:params, // 使用blob下载
    
  })
},
// 查询
getxm(userid){
  return request({
    url: `${api_name}/getxm/${userid}`,
    method: 'post',
    
    
    
  })
},
// 删除文件
deleteFile(params){
  return request({
    url: `${api_name}/delete`,
    method: 'post',
    data:params
    
    
  })
}

}