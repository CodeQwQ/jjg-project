<template>
    <div class="app-container">
      <!--查询表单-->
      <div class="search-div">
        <el-form :inline="true" class="demo-form-inline">
          <el-form-item label="匝道收费站名称:">
            <el-input v-model="searchObj.zdsfzname" />
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
      </div>
      <!-- 工具按钮 -->
      <div class="tools-div" v-show="!(username=='admin'&&type==1  )">
        <div class="anniu">
          <el-button
            class="btn-add"
            type="success"
            @click="exportsfz()"
            style="margin-left: 0px"
            icon="el-icon-bottom"
            size="mini"
            >导出收费站清单模板文件</el-button
          >
          <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
          <el-button
            class="btn-add"
            type="success"
            @click="importsfz()"
            style="margin-left: 0px"
            icon="el-icon-top"
            size="mini"
            >导入收费站清单数据</el-button
          >
          <!-- :disabled="$hasBP('bnt.ql.import')  === true" -->
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
            >删除</el-button
          >
          <!-- <el-button
            class="btn-add"
            type="danger"
            size="mini"
            @click="removeAll()"
            style="margin-left: 1px"
            icon="el-icon-delete"
            >全部删除</el-button> -->
          <el-button class="shuoming" style="float：right" type="text" @click="dialogVisible = true"
            >填写说明</el-button
          >
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
        <el-table-column prop="zdsfzname" label="匝道收费站名称" />
        <el-table-column prop="htd" label="合同段" />
        <el-table-column prop="lf" label="路幅" />
        <el-table-column prop="zhq" label="桩号起" />
        <el-table-column prop="zhz" label="桩号止" />
        <el-table-column prop="pzlx" label="铺筑类型" />
        
        <el-table-column prop="sszd" label="所属匝道(填字母A/B/M)" width="180px"/>
        <el-table-column prop="sshtmc" label="所属互通名称" />
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
  
      <el-dialog
        title="填写说明"
        :visible.sync="dialogVisible"
        width="30%"
        
      >
        <span style="white-space: pre-wrap;">{{ instruction }}</span>
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/路桥隧/收费站清单.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
        <span slot="footer" class="dialog-footer">
          <el-button type="primary" @click="dialogVisible = false"
            >确 定</el-button
          >
        </span>
      </el-dialog>
      <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm" :model="modifyData"  label-width="200px" size="small" style="padding-right: 40px;">
            <el-form-item label="匝道收费站名称" prop="zdsfzname">
            <el-input v-model="modifyData.zdsfzname"></el-input>
            </el-form-item>
            <el-form-item label="合同段" prop="htd">
            <el-input v-model="modifyData.htd"></el-input>
            </el-form-item>
            <el-form-item label="路幅" prop="lf">
            <el-input v-model="modifyData.lf"></el-input>
            </el-form-item>
            <el-form-item label="桩号起" prop="zhq">
            <el-input v-model="modifyData.zhq"></el-input>
            </el-form-item>
            <el-form-item label="桩号止" prop="zhz">
            <el-input v-model="modifyData.zhz"></el-input>
            </el-form-item>
            <el-form-item label="铺筑类型" prop="pzlx">
            <el-input v-model="modifyData.pzlx"></el-input>
            </el-form-item>
            <el-form-item label="所属匝道(填字母A/B/M)" prop="sszd">
            <el-input v-model="modifyData.sszd"></el-input>
            </el-form-item>
            <el-form-item label="所属互通名称" prop="sshtmc">
            <el-input v-model="modifyData.sshtmc"></el-input>
            </el-form-item>
            
           
            
            
            <el-form-item label="创建时间" prop="createtime">
            <el-input v-model="modifyData.createtime"></el-input>
            </el-form-item>
          </el-form>
          <span slot="footer" class="dialog-footer">
            <el-button @click="dialogVisible2 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
            <el-button type="primary" icon="el-icon-check" @click="update()" size="small">确 定</el-button>
          </span>
        </el-dialog>
    </div>
  </template>
  <script>
  // 引入定义接口js文件
  import api from "@/api/jgproject/lqszb/sfzqd.js";
  import { Loading } from 'element-ui';
  import Cookies from 'js-cookie'
  export default {
    data() {
      // 初始值
      let value = this.$route.query.name;
      return {
        data: {
          proname: value,
        },
        dialogImportVisible: false,
        list: [], // 项目列表
        total: 0, // 总记录数
        page: 1, // 当前页码
        limit: 10, // 每页记录数
        searchObj: {}, // 查询条件
        multipleSelection: [], // 批量删除选中的记录列表
        dialogVisible: false,
        dialogVisible2:false,
        modifyData:{},
        username:Cookies.get('username'),
        userid:Cookies.get('userid'),
        type:Cookies.get('type'),

        file:'',// 待上传文件
        fileList:[],
        sfz: {},
        instruction:"1.日期格式：yyyy-MM-dd，例如：2021-05-23"+"\n"+"2.左右幅时，请填写左幅或右幅，请勿填写左线或右线",

      };
    },
    created() {
      // 页面渲染之前
      console.log(this.$route.query.name);
      //alert(this.$route.matched[1].meta.title)
      this.fetchData();
    },
    methods: {
      handleClose(){
  
      },
      importsfz() {
          this.dialogImportVisible = true;
        },
        // 上传文件触发
        fileChange(file,fileList){
          this.fileList=fileList
          
        },
        onUploadSuccess() {
          
          //debugger
          let fd=new FormData()
          
          
    
          fd.append('file',this.fileList[0].raw)
          fd.append('proname',this.$route.query.projecttitle)
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importsfz(fd).then((res)=>{
            console.log('res',res)
            if(res.message=='成功'){
              this.fetchData();
              this.$message({
                type: "success",
                message: "上传成功!",
              });
              this.dialogImportVisible = false;
              this.fileList=[]
            }
            else{
              this.$alert('上传失败！')
            }
            load.close()
          }).catch((err)=>{
              
              load.close()
            })
          
          
          
        },
      exportsfz() {
        let projectname = this.$route.query.projecttitle;
        api.exportsfz().then((res) => {
          const objectUrl = URL.createObjectURL(
            new Blob([res.data], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            })
          );
          const link = document.createElement("a");
          // 设置导出的文件名称
          link.download = projectname + `收费站清单` + ".xlsx";
          link.style.display = "none";
          link.href = objectUrl;
          link.click();
          document.body.appendChild(link);
        });
        //window.open("http://localhost:8800/project/info/lqs/exportQL")
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
      fetchData() {
        this.searchObj.proname = this.$route.query.projecttitle;
        this.searchObj.userid=this.userid
        // 调用api
        api.pageList(this.page, this.limit, this.searchObj).then((response) => {
          this.list = response.data.records;
          this.total = response.data.total;
        });
      },
       // 修改
       modify(){
          if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要修改的记录！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能修改一条记录！");
            return;
          }
          this.modifyData={}
          api.getById(this.multipleSelection[0].id).then((res)=>{
            if(res.message=='成功'){
              this.dialogVisible2=true
              this.modifyData=res.data
            }
            else{
              this.$alert('出错')
            }
          })
          
        },
        update(){
          api.modify(this.modifyData).then((res)=>{
            if(res.message=='成功'){
              this.$message({
                type: "success",
                message: "修改成功!",
              });
              // 刷新页面
              this.dialogVisible2=false
              
              this.fetchData();
            }
            else{
              this.$alert('出错')
            }
          })
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
  
  <style lang="scss" scoped>
  .tools-div{
    .anniu{
      .shuoming{
        float: right;
      }
    }
  }
  </style>
  