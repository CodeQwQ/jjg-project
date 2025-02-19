import request from '@/utils/request'

const api_name = '/admin/system/sysMenu'

export default {
    /*
  获取实测数据的权限(菜单/功能)列表
  */
  findNodesByName(projectname,menuname) {
    return request({
      url: `${api_name}/findNodesByName/${projectname}/${menuname}`,
      method: 'get'
    })
  }
}