import request from '@/utils/request'

const api_name = '/xmtj'

export default {

    // 查询所有项目信息
  list() {
    return request({
      url: `${api_name}/findAll`,
      method: `get`
    })
  },
     //分页查询
  pageList(current, limit) {
    return request({
      url: `${api_name}/findQueryPage/${current}/${limit}`,
      method: 'post',
      
    })
  },
}
// 查询所有项目信息
