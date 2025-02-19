import request from '@/utils/request'

//常量
const api_name = '/admin/system/sysUserMenu'

export default {
    //列表
    saveUserMenus(data) {
        return request({
            //接口路径
            url: `${api_name}/saveMenus`,
            method: 'post', //提交方式
            //参数
            data: data
        })
    },
}