import request from '@/utils/request'
import qs from 'qs'

const api_name = '/jjg/dpksh'

export default {

    // 获取合同段信息
    gethtdinfo(params){
        
        return request({
        url: `${api_name}/gethtd?proname=`+params,
        method: 'post',
        data:params, // 使用blob下载
        
        })
    },
    // 获取桥梁隧道等数量
    getqsnum(params){
  
        return request({
        url: `${api_name}/getnum?proname=`+params,
        method: 'post',
        data:params, // 使用blob下载
        
        })
    },
    // 获取单位工程
    getdwgc(params){
  
        return request({
        url: `${api_name}/getdwgc?proname=`+params,
        method: 'post',
        data:params, // 使用blob下载
        
        })
    },
    // 获取合同段数据
    gethtddata(params){
  
        return request({
        url: `${api_name}/gethtddata?proname=`+params,
        method: 'post',
        data:params, // 使用blob下载
        
        })
    },
    // 获取建设项目数据
    getjsxmdata(params){
  
        return request({
        url: `${api_name}/getjsxmdata?proname=`+params,
        method: 'post',
        data:params, // 使用blob下载
        
        })
    },
   


}