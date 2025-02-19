<template>
    <div class="app-container">
      
      <!-- 工具按钮 -->
      <div class="tools-div" v-show="!(username=='admin'&&type==1  )">
      <div class="anniu">
        <el-button
          class="btn-add"
          type="success"
          @click="add()"
          style="margin-left: 0px"
          icon="el-icon-plus"
          size="mini"
          
          >新增</el-button
        >
        <!--:disabled="$hasBP('bnt.project.add')  === true" -->
        <el-button
          class="btn-add"
          type="danger"
          size="mini"
          @click="batchRemove()"
          style="margin-left: 1px"
          icon="el-icon-delete"
          
          >删除</el-button
        >
        
        <!--:disabled="$hasBP('bnt.project.remove')  === true"-->
      </div>
      </div>
      <!-- 表格 -->
      <el-table
        :data="list"
        border
        stripe
        :header-cell-style="{ 'text-align': 'center' }"
        :cell-style="{ 'text-align': 'center' }"
        @selection-change="handleSelectionChange"
      >
        <el-table-column type="selection" />
        <el-table-column label="序号" width="102">
          <template slot-scope="scope">
            {{ (page - 1) * limit + scope.$index + 1 }}
          </template>
        </el-table-column>
        <el-table-column prop="proname" label="项目名称" width="162" />
        <el-table-column prop="xmqc" label="项目全称"  />
        <el-table-column prop="grade" label="公路等级" width="162" />
        <el-table-column prop="createTime" label="创建时间" />
  
        <el-table-column label="操作"  align="center" fixed="right" width="400">
          <template slot-scope="scope">
            <el-button type="primary"  size="mini" @click="gotoNew(scope.row)" >JTG F80/1-2017规范</el-button>
            <el-button type="primary"  size="mini" @click="gotoOld(scope.row)" >JTG F80/1-2004规范</el-button>
            
          </template>
        </el-table-column>
      </el-table>
      <!-- 分页组件 -->
      <el-pagination
        :current-page="page"
        :total="total"
        :page-size="limit"
        :page-sizes="[5, 10, 20, 30, 40, 50, 100]"
        style="padding: 30px 0; text-align: center"
        layout=" ->,total, sizes, prev, pager, next, jumper"
        @size-change="changePageSize"
        @current-change="changeCurrentPage"
      />
  
      <el-dialog title="添加" :visible.sync="dialogVisible" width="40%" >
        <el-form ref="dataForm" :model="sysproject"  label-width="100px" size="small" style="padding-right: 40px;" :rules="rules">
          <el-form-item label="项目名称" prop="proname">
            <el-input v-model="sysproject.proname" />
          </el-form-item>
          <el-form-item label="项目全称" prop="xmqc">
            <el-input v-model="sysproject.xmqc" />
          </el-form-item>
          <el-form-item label="公路等级" prop="grade">
            <el-input v-model="sysproject.grade"/>
          </el-form-item>
          <el-form-item label="创建时间" prop="createTime">
            <el-date-picker
            v-model="sysproject.createTime"
            value-format="yyyy-MM-dd HH:mm:ss"
          />
          </el-form-item>
        </el-form>
        <span slot="footer" class="dialog-footer">
          <el-button @click="dialogVisible = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
          <el-button type="primary" icon="el-icon-check" @click="save()" size="small">确 定</el-button>
        </span>
      </el-dialog>
    </div>
  </template>
  <script>
  // 引入定义接口js文件
  
  import api from '@/api/jgproject/projectinfo.js'
  import Cookies from 'js-cookie'
  
  export default {
    data() {
      // 初始值
      return {
        list: [], // 项目列表
        total: 0, // 总记录数
        page: 1, // 当前页码
        limit: 10, // 每页记录数
        searchObj: {}, // 查询条件
        multipleSelection: [], // 批量删除选中的记录列表
        dialogVisible: false,
        username:Cookies.get('username'),
        userid:Cookies.get('userid'),
        type:Cookies.get('type'),
        sysproject: {},
        rules:{
            proname:[
              {
                required:true,
                message:"项目名称不能为空！"
              },
              
            ],
            xmqc:[
              {
                required:true,
                message:"项目全称不能为空！"
              },
              
            ],
            
            grade:[
              {
                required:true,
                message:"公路名称不能为空！"
              },
              
            ],
            createTime:[
              {
                required:true,
                message:"创建时间不能为空！"
              },
              
            ]
          }

      }
    },
    created() {
      // 页面渲染之前
      
      this.fetchData()
      
    },
    methods: {
      // 批量删除
      batchRemove() {
        // 判断非空
        if (this.multipleSelection.length === 0) {
          this.$message.warning('请选择要删除的记录！')
          return
        }
        this.$confirm('此操作将删除该项目的所有信息, 是否继续?', '提示', {
          confirmButtonText: '确定',
          cancelButtonText: '取消',
          type: 'warning'
        }).then(() => {
          var idList = [] // [1,2,3 ]
          // 遍历数组
          for (var i = 0; i < this.multipleSelection.length; i++) {
            var obj = this.multipleSelection[i]
            var id = obj.id
            idList.push(id)
          }
          // 调用接口删除
          api.batchRemove(idList).then((response) => {
            // 提示信息
            this.$message({
              type: 'success',
              message: '删除成功!'
            })
            // 刷新页面
            location.reload();
            //this.$router.go(0)
            this.fetchData()
          })
        })
      },
      // 复选框发生变化，调用方法，选中复选框行内容传递
      handleSelectionChange(selection) {
        this.multipleSelection = selection
      },
      // 跳转到添加的表单页面
      // add() {
      //   this.$router.push({ path: '/project/create' })
      // },
      add() {
        this.dialogVisible = true
        this.sysproject = {}
      },
      save() {
        api.saveProject(this.sysproject).then((response) => {
          this.list = response.data
          // 提示
          this.$message({
            type: 'success',
            message: '添加成功!'
          })
          // 关闭弹框
            this.dialogVisible = false
            // 刷新
            this.fetchData()
            location.reload()
        })
      },
      // 新规范
      gotoNew(row){
        this.$router.push({path:"/jgproject/guifan/new",query:{projecttitle:row.proname}});
      },
      // 旧规范
      gotoOld(row){
        this.$router.push({path:"/jgproject/guifan/old",query:{projecttitle:row.proname}});
      },
      fetchData() {
        this.searchObj.userid=this.userid
        this.searchObj.username=this.username
        // 调用api
        api
          .pageList(this.page, this.limit, this.searchObj)
          .then((response) => {
            this.list = response.data.records
            this.total = response.data.total
            console.log("ssqs",this.list)
          })
      },
      // 改变每页显示的记录数
      changePageSize(size) {
        this.limit = size
        this.fetchData()
      },
      // 改变页码数
      changeCurrentPage(page) {
        this.page = page
        this.fetchData()
      },
      // 清空
      resetData() {
        this.searchObj = {}
        this.fetchData()
      }
    }
  }
  </script>
  