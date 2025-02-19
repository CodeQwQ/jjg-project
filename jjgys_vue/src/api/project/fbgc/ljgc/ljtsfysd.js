import request from '@/utils/request'

const api_name = '/jjg/fbgc/ljgc/ljtsfysd'

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
  // 导出模板文件
  exportysdht(){
    return request({
      url: `${api_name}/exportysdht`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  exportysdsl(){
    return request({
      url: `${api_name}/exportysdsl`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  // 导入文件
  importysdht(params){
    return request({
      url: `${api_name}/importysdht`,
      method: 'post',
      data:params, // 使用blob下载
      
    })
  },
  importysdsl(params){
    return request({
      url: `${api_name}/importysdsl`,
      method: 'post',
      data:params, // 使用blob下载
      
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
  // 通过id查询
  getById(id){
    return request({
      url: `${api_name}/getHt/`+id,
      method: 'get',
      
    })
  },
  getById1(id){
    return request({
      url: `${api_name}/getSl/`+id,
      method: 'get',
      
    })
  },
   // 修改
   modify(data) {
    return request({
      url: `${api_name}/updateHt`,
      method: `post`,
      data: data
    })
  },
  modify1(data) {
    return request({
      url: `${api_name}/updateSl`,
      method: `post`,
      data: data
    })
  },
   // 批量删除
   batchRemove(idList) {
    return request({
      url: `${api_name}/removeBeatchHt`,
      method: `delete`,
      data: idList
    })
  },
  batchRemove1(idList) {
    return request({
      url: `${api_name}/removeBeatchSL`,
      method: `delete`,
      data: idList
    })
  },
  //全部删除灰土
  removeAllHt(params) {
    return request({
      url: `${api_name}/removeAllHt`,
      method: `delete`,
      data: params
    })
  },
  //全部删除灰土
  removeAllSl(params) {
    return request({
      url: `${api_name}/removeAllSl`,
      method: `delete`,
      data: params
    })
  },

    //分页查询灰土
  pageList(current, limit, searchObj) {
    return request({
      url: `${api_name}/findQueryPageHt/${current}/${limit}`,
      method: 'post',
      data: searchObj
    })
  },
  pageList1(current, limit, searchObj) {
    return request({
      url: `${api_name}/findQueryPageSL/${current}/${limit}`,
      method: 'post',
      data: searchObj
    })
  }

}