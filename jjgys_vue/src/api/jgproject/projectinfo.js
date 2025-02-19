import request from '@/utils/request'

const api_name = '/jjg/jggl'

export default {
  

  // 项目条件查询分页
  // current当前页  limit每页记录数 searchObj条件对象
  pageList(current, limit, searchObj) {
    return request({
      url: `${api_name}/findQueryPage/${current}/${limit}`,
      method: 'post',
      // 使用json格式传递  写法 data:searchObj
      // 使用普通格式传递  写法 params:searchObj
      data: searchObj
    })
  },
  // 添加项目
  saveProject(project) {
    return request({
      url: `${api_name}/add`,
      method: 'post',
      data: project
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
  
}
