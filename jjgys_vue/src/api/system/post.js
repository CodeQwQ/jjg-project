import request from '@/utils/request'

//常量
const api_name = '/admin/system/sysPost'

export default {
    //列表
    getPageList(page,limit,searchObj) {
        return request({
            //接口路径
            url: `${api_name}/${page}/${limit}`,
            method: 'get', //提交方式
            //参数
            params: searchObj
        })
    },
    //删除
    removeId(id) {
        return request({
            //接口路径
            url: `${api_name}/remove/${id}`,
            method: 'delete' //提交方式
        })
    },
    //添加
    save(sysPost) {
        return request({
            //接口路径
            url: `${api_name}/save`,
            method: 'post', //提交方式
            //传递json格式
            data: sysPost
        })
    },
    //根据id查询
    getPostId(id) {
        return request({
            //接口路径
            url: `${api_name}/findPostById/${id}`,
            method: 'get' //提交方式
        })
    },
    //修改的方法
    update(sysPost) {
        return request({
            //接口路径
            url: `${api_name}/update`,
            method: 'post', //提交方式
            data: sysPost
        })
    },
    //批量删除
    batchRemove(idList) {
        return request({
            //接口路径
            url: `${api_name}/batchRemove`,
            method: 'delete', //提交方式
            data: idList
        })
    }
}