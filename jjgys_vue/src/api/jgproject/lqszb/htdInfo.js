import request from '@/utils/request'
import qs from 'qs'

const api_name = '/project/jg/htd'

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
  // 删除
  remove(idList) {
    return request({
      url: `${api_name}/remove`,
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


}