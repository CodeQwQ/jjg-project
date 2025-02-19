import request from '@/utils/request'

const api_name = '/admin/system/sysUser'

export default {
  selectdept(){
    return request({
      url: '/dept/selectdept',
      method: 'get'
    })
  },
  updatepwd(username,newpass){
    return request({
      url: `${api_name}/updatepwd/${username}/${newpass}`,
      method: 'post'
    })
  },
  checkSysUserPwd(oldpw,username){
    return request({
      url: `${api_name}/checkSysUserPwd/${oldpw}/${username}`,
      method: 'get'
    })
  },
//管理员账号查询全部用户信息
pageList(current, limit, searchObj) {
  return request({
    url: `${api_name}/useradmin/${current}/${limit}`,
    method: 'post',
    data: searchObj
  })
},
//添加
save(user) {
  return request({
      //接口路径
      url: `${api_name}/save`,
      method: 'post', //提交方式
      //参数
      data: user
  })
},
//根据id查询
getUserId(id) {
  return request({
      //接口路径
      url: `${api_name}/getUser/${id}`,
      method: 'get' //提交方式
  })
},
getUserByName(username){
  return request({
    //接口路径
    url: `${api_name}/getUserByName/${username}`,
    method: 'get' //提交方式
})
},
//修改
update(user) {
  return request({
      //接口路径
      url: `${api_name}/update`,
      method: 'post', //提交方式
      //参数
      data: user
  })
},
  //删除
  removeById(id) {
    return request({
        //接口路径
        url: `${api_name}/remove/${id}`,
        method: 'delete' //提交方式
    })
}
}
