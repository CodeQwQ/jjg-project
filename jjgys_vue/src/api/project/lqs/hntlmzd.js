import request from '@/utils/request'

const api_name = '/project/info/lqs/hntlmzd'

export default {
  exporthntlmzd(){
    return request({
      url: `${api_name}/export`,
      method: 'get',
      responseType: 'blob', // 使用blob下载
    })
  },
  importhntlmzd(params){
    return request({
      url: `${api_name}/importhntlmzd`,
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
  // 添加合同段
  save(proHtd) {
    return request({
      url: `${api_name}/save`,
      method: 'post',
      data: proHtd
    })
  },

}