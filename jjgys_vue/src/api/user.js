import request from '@/utils/request'

export function login(data) {
  return request({
    url: '/admin/system/index/login',// 后端URL
    method: 'post',
    data
  })
}

export function getInfo(token) {
  return request({
    url: '/admin/system/index/info',
    method: 'get',
    params: { token }
  })
}
export function register(userInfoEntity) {
  return request({
    url: '/user/register',
    method: 'post',
    data: userInfoEntity
  })
}

export function logout() {
  return request({
    url: '/vue-admin-template/user/logout',
    method: 'post'
  })
}
