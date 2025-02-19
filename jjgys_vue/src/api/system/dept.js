import request from '@/utils/request'

//常量
const api_name = '/dept'

export default {
    //公司部门树
    selectdeptTree() {
        return request({
            //接口路径
            url: `${api_name}/selectdeptTree`,
            method: 'get' //提交方式
        })
    },
    //列表
    selectdeptInfo() {
        return request({
            //接口路径
            url: `${api_name}/selectdeptinfo`,
            method: 'get' //提交方式
        })
    },
    // 添加公司
    addCompany(params){
        return request({
          url: `${api_name}/addcompany`,
          method: 'post',
          
          data: params// 使用blob下载
        })
      },
    // 查询公司
    selectCompany() {
        return request({
            //接口路径
            url: `${api_name}/selectcompany`,
            method: 'get' //提交方式
        })
    },
    // 校验
    checkDept(params){
      return request({
        url: `${api_name}/checkbm`,
        method: 'post',
        
        data: params// 使用blob下载
      })
    },
    // 添加部门
    addDept(params){
        return request({
          url: `${api_name}/add`,
          method: 'post',
          
          data: params// 使用blob下载
        })
      },
    // 修改部门
    update(params){
        return request({
          url: `${api_name}/updata`,
          method: 'put',
          
          data: params// 使用blob下载
        })
      },
    // 删除部门
   remove(id) {
    return request({
      url: `${api_name}/remove/`+id,
      method: `delete`,
      data: id
    })
  },
}