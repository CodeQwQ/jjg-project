<template>
  <div class="app-container">
    <!--查询表单-->
    <div class="search-div">
      <el-form label-width="70px" size="small">
        <el-row>
          <el-col :span="24">
            <el-input style="width: 20%" v-model="searchObj.username" placeholder="用户名"></el-input>
            <el-button type="primary" icon="el-icon-search" size="mini"  @click="fetchData()">搜索</el-button>
            <el-button icon="el-icon-refresh" size="mini" @click="resetData">重置</el-button>
          </el-col>
        </el-row>
        <el-row style="display:flex">
          <!-- <el-button type="primary" icon="el-icon-search" size="mini"  @click="fetchData()">搜索</el-button>
          <el-button icon="el-icon-refresh" size="mini" @click="resetData">重置</el-button> -->
        </el-row>
      </el-form>
    </div>
    <!-- 工具条 -->
  <div class="tools-div">
    <el-button class="btn-add" size="mini" @click="batchRemove()" >批量删除</el-button>
  </div>
    <!-- 表格 -->
    <el-table
      v-loading="listLoading"
      :data="list"
      stripe
      border
      style="width: 100%;margin-top: 10px;"
      @selection-change="handleSelectionChange">

      <el-table-column type="selection"/>

      <el-table-column
        label="序号"
        width="70"
        align="center">
        <template slot-scope="scope">
          {{ (page - 1) * limit + scope.$index + 1 }}
        </template>
      </el-table-column>

      <el-table-column prop="username" label="用户名" />
      <el-table-column prop="ipaddr" label="登录IP地址" />
      <el-table-column prop="msg" label="提示信息" width="160"/>
      <el-table-column prop="createTime" label="访问时间" width="160"/>
      
    </el-table>

  <!-- 分页组件 -->
    <el-pagination
      :current-page="page"
      :total="total"
      :page-size="limit"
      style="padding: 30px 0; text-align: center;"
      layout="total, prev, pager, next, jumper"
      @current-change="fetchData"
    />
  </div>
</template>
<script>
//引入定义接口的js文件
import api from '@/api/system/loginLog'

export default {
  //定义初始值
  data() {
    return {
      listLoading:false,//是否显示加载
      list:[],//操作日志列表
      total:0,//总记录数
      page:1,//当前页
      limit:10,//每页显示记录数
      searchObj:{},//条件查询封装对象

      dialogVisible: false,//弹出框
      sysRole:{}, //封装添加表单数据
      selectValue:[] //复选框选择内容封装数组
    }
  },
  //页面渲染之前执行
  created() {
    //调用列表方法
    this.fetchData()
  },
  methods:{//具体方法
    //复选框发生变化执行方法
    handleSelectionChange(selection) {
      this.selectValue = selection
      //console.log(this.selectValue)
    },
    //批量删除
    batchRemove() {
      //判断
      if(this.selectValue.length==0) {
        this.$message.warning('请选择要删除的记录！')
        return
      }
      this.$confirm('此操作将永久删除该角色, 是否继续?', '提示', {
          confirmButtonText: '确定',
          cancelButtonText: '取消',
          type: 'warning'
        }).then(() => {
          //数组
          var idList = []
          //获取多个复选框对应id，封装到数组里面
          // [1,2,3]
          for(var i=0;i<this.selectValue.length;i++) {
            var obj = this.selectValue[i]
            //id值
            var id = obj.id
            //放到数组里面
            idList.push(id)
          }
          api.batchRemove(idList).then(response => {
              //提示
              this.$message({
                type: 'success',
                message: '删除成功!'
              });
              //刷新
              this.fetchData()
          })
        })
    },
    //重置
    resetData() {
      //清空表单
      this.searchObj = {}
      //查询所有数据
      this.fetchData()
    },
    // 条件分页查询列表  pageNum 查询页数
    fetchData(pageNum=1) {
      //页数赋值
      this.page = pageNum
      //ajax
      api.getPageList(this.page,this.limit,this.searchObj)
        .then(response => {
          //this.listLoading = false
          //console.log(response)
          //每页数据列表
          this.list = response.data.records
          //总记录数
          this.total = response.data.total
        })
    }
  }
}
</script>