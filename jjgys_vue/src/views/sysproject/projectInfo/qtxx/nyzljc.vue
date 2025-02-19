<template>
    <div class="app-container">
      <!--查询表单
      <div class="search-div">
        <el-form :inline="true" class="demo-form-inline">
          <el-form-item label="合同段名称:">
            <el-input v-model="searchObj.name" />
          </el-form-item>
          <el-button
            type="primary"
            icon="el-icon-search"
            size="mini"
            @click="fetchData()"
            >搜索</el-button
          >
          <el-button icon="el-icon-refresh" size="mini" @click="resetData"
            >重置</el-button
          >
        </el-form>
      </div>-->
      <!-- 工具按钮 -->
      <div class="tools-div" v-show="!(username=='admin'&&type==1)">
        <div class="anniu">
          
          
          <!--:disabled="$hasBP('bnt.htd.remove') === true"-->
          <el-button
            class="btn-add"
            type="primary"
            size="mini"
            @click="exportnyzljc()"
            style="margin-left: 1px"
            icon="el-icon-bottom"
            
            >导出内页资料检查模板文件</el-button
          >
  
          <el-button
            class="btn-add"
            type="primary"
            size="mini"
            @click="importnyzljc()"
            style="margin-left: 1px"
            icon="el-icon-top"
            
            >导入内页资料检查数据文件</el-button
          >
            <el-button
            class="btn-add"
            type="primary"
            size="mini"
            @click="modify()"
            style="margin-left: 1px"
            icon="el-icon-edit"
            >修改</el-button>
            <el-button
            class="btn-add"
            type="danger"
            size="mini"
            @click="batchRemove()"
            style="margin-left: 1px"
            icon="el-icon-delete"
            >删除</el-button>
            
          
          <!--:disabled="$hasBP('bnt.htd.remove') === true"-->
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
        <el-table-column label="序号">
          <template slot-scope="scope">
            {{ (page - 1) * limit + scope.$index + 1 }}
          </template>
        </el-table-column>
        <el-table-column prop="htd" label="合同段" />
        <el-table-column prop="wt" label="存在问题" />
        <el-table-column prop="kf" label="扣分" />
        
        <el-table-column prop="createTime" label="创建时间" />
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
  
      <el-dialog title="修改" :visible.sync="dialogVisible" width="40%">
        <el-form
          ref="dataForm"
          :model="nyzljc"
          label-width="100px"
          size="small"
          :rules="rules"
          style="padding-right: 40px"
        >
          <el-form-item label="合同段" prop="htd">
            <el-input v-model="nyzljc.htd" />
          </el-form-item>
          <el-form-item label="存在问题" prop="wt">
            <el-input v-model="nyzljc.wt" />
          </el-form-item>
          <el-form-item label="扣分" prop="kf">
            <el-input v-model="nyzljc.kf" />
          </el-form-item>
          
          <el-form-item label="创建时间" prop="createTime">
            <el-date-picker
              v-model="nyzljc.createTime"
              value-format="yyyy-MM-dd HH:mm:ss"
            />
          </el-form-item>
        </el-form>
        <span slot="footer" class="dialog-footer">
          <el-button
            @click="dialogVisible = false"
            size="small"
            icon="el-icon-refresh-right"
            >取 消</el-button
          >
          <el-button
            type="primary"
            icon="el-icon-check"
            @click="update()"
            size="small"
            >确 定</el-button
          >
        </span>
      </el-dialog>
      <el-dialog title="导入" :visible.sync="dialogImportVisible" width="480px">
        <el-form label-position="right" label-width="170px">
          <el-form-item label="文件">
            <el-upload
              :multiple="false"
              :on-change="fileChange"
              :auto-upload="false"
              :file-list="fileList"
              :limit="1"
              action="/"
              class="upload-demo"
            >
              <el-button size="small" type="primary">点击上传</el-button>
              <div slot="tip" class="el-upload__tip">只能上传xls文件</div>
            </el-upload>
          </el-form-item>
        </el-form>
        <div slot="footer" class="dialog-footer">
          <el-button @click="onUploadSuccess">提交</el-button>
          <el-button @click="dialogImportVisible = false,fileList=[]">取消</el-button>
        </div>
      </el-dialog>
    </div>
  </template>
  <script>
  // 引入定义接口js文件
  import api from "@/api/project/nyzljc.js";
  
  
  import FileSaver from "file-saver"
  import { Loading } from 'element-ui';
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
        dialogImportVisible:false,
        username:Cookies.get('username'),
        userid:Cookies.get('userid'),
        type:Cookies.get('type'),

        file:'',// 待上传文件
        fileList:[],
        rules:{
           htd:[
            {
                required:true,message:'合同段不能为空'
            }
           ],
           
           
        },
        nyzljc: {
          
        },
        
      }
    },
    created() {
      
      this.fetchData();
    },
    methods: {
        modify(){
          if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要修改的记录！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能修改一条记录！");
            return;
          }
          this.nyzljc={}
          api.getById(this.multipleSelection[0].id).then((res)=>{
            if(res.message=='成功'){
              this.dialogVisible=true
              this.nyzljc=res.data
            }
            else{
              this.$alert('出错')
            }
          })
          
        },
        update(){
          this.$refs["dataForm"].validate((valid)=>{
              if(valid){
                api.modify(this.nyzljc).then((res)=>{
                  if(res.message=='成功'){
                    this.$message({
                      type: "success",
                      message: "修改成功!",
                    });
                    // 刷新页面
                    this.dialogVisible=false
                    
                    this.fetchData();
                  }
                  else{
                    this.$alert('出错')
                  }
                })
              }
            })
          
        },
      // 批量删除
      batchRemove() {
        // 判断非空
        if (this.multipleSelection.length === 0) {
          this.$message.warning("请选择要删除的记录！");
          return;
        }
        this.$confirm("此操作将删除该项目的所有信息, 是否继续?", "提示", {
          confirmButtonText: "确定",
          cancelButtonText: "取消",
          type: "warning",
        }).then(() => {
          var idList = []; // [1,2,3 ]
          // 遍历数组
          for (var i = 0; i < this.multipleSelection.length; i++) {
            var obj = this.multipleSelection[i];
            var id = obj.id;
            idList.push(id);
          }
          // 调用接口删除
          api.batchRemove(idList).then((response) => {
            // 提示信息
            this.$message({
              type: "success",
              message: "删除成功!",
            });
            // 刷新页面
            location.reload();
            this.fetchData();
          });
        });
      },
       //全部删除
        removeAll() {
          
          this.$confirm("此操作将删除该项目的所有信息, 是否继续?", "提示", {
            confirmButtonText: "确定",
            cancelButtonText: "取消",
            type: "warning",
          }).then(() => {
            let commonInfoVo={}
            commonInfoVo.proname=this.proname
            commonInfoVo.userid=this.userid
            // 调用接口删除
            api.removeAll(commonInfoVo).then((response) => {
              // 提示信息
              this.$message({
                type: "success",
                message: "删除成功!",
              });
              // 刷新页面
              this.fetchData();
            });
          });
        },
      // 复选框发生变化，调用方法，选中复选框行内容传递
      handleSelectionChange(selection) {
        this.multipleSelection = selection;
      },
      // 跳转到添加的表单页面
      // add() {
      //   this.$router.push({ path: '/project/create' })
      // },
     
      exportnyzljc() {
          
            let projectname = this.$route.query.projecttitle;
            
            
            
            let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
            api.exportnyzljc().then((response) => {
                var temp = response.headers["content-disposition"]
                var filename=temp.split('=')[1]
                filename=filename.replace(/['"]/g, '')
                
                FileSaver.saveAs(response.data,decodeURI(filename))
                load.close()
            }).catch((err)=>{
                
                load.close()
              });
      },
      
      importnyzljc(){
        this.dialogImportVisible=true
      },
       // 上传文件触发
       fileChange(file,fileList){
          this.fileList=fileList
          
          
        },
       onUploadSuccess(response, file) {
          let fd=new FormData()
          let projectname = this.$route.query.projecttitle;
          
          
  
          fd.append('file',this.fileList[0].raw)
          fd.append('proname',projectname)
          fd.append('userid',this.userid)
          
          
          
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
          api.importnyzljc(fd).then((res)=>{
            console.log('res',res)
            if(res.message=='成功'){
              
              this.$message({
                type: "success",
                message: "上传成功!",
              });
              this.dialogImportVisible = false;
              this.fileList=[]
              this.fetchData()
            }
            else{
              this.$alert('上传失败！')
            }
            load.close()
          }).catch((err)=>{
                
                load.close()
              });
      },
      
      fetchData() {
        
        // 调用api
        this.searchObj.proname = this.$route.query.projecttitle
        this.searchObj.userid=this.userid
        api.pageList(this.page, this.limit, this.searchObj).then((response) => {
          this.list = response.data.records;
          this.total = response.data.total;
          console.log("sq",this.list)
        });
      },
      // 改变每页显示的记录数
      changePageSize(size) {
        this.limit = size;
        this.fetchData();
      },
      // 改变页码数
      changeCurrentPage(page) {
        this.page = page;
        this.fetchData();
      },
      // 清空
      resetData() {
        this.searchObj = {};
        this.fetchData();
      },
    },
  };
  </script>
  