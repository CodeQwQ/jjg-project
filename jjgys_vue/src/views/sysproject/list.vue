<template>
  <div class="app-container">
    <!--查询表单-->
    <div class="search-div">
    <el-form :inline="true" class="demo-form-inline">
      <el-form-item label="项目名称:">
        <el-input v-model="searchObj.proName" />
      </el-form-item>
      <!-- <el-button type="primary" icon="el-icon-search" @click="fetchData()">查询</el-button> -->
      <el-button type="primary" icon="el-icon-search" size="medium"  @click="fetchData()">搜索</el-button>
      <el-button icon="el-icon-refresh" size="medium" @click="resetData">重置</el-button>

      <!-- <el-button
        type="primary"
        icon="el-icon-refresh"
        @click="resetData()"
        style="margin-left: 1px; width: 89px"
        >清空</el-button
      > -->
    </el-form>
  </div>
    <!-- 工具按钮 -->
    <div class="tools-div">
    <div class="anniu">
      <el-button
        class="btn-add"
        type="success"
        @click="add()"
        style="margin-left: 0px"
        icon="el-icon-plus"
        size="mini"
        v-show="type==2 || type == 4"
        >添加</el-button
      >
      <!--:disabled="$hasBP('bnt.project.add')  === true"-->
      <el-button
          class="btn-add"
          type="primary"
          size="mini"
          @click="modify()"
          style="margin-left: 1px"
          icon="el-icon-edit"
          v-show="type==2 || type == 4"
          >修改</el-button>
      <el-button
        class="btn-add"
        type="danger"
        size="mini"
        @click="batchRemove()"
        style="margin-left: 1px"
        icon="el-icon-delete"
        v-show="type==2 || type == 4"
        >删除</el-button
      >
      <el-button
        class="btn-add"
        type="primary"
        size="mini"
        @click="scbgzBg()"
        style="margin-left: 1px"
        v-show="type==2 || type == 4"
        >生成报告中表格</el-button
      >
      <el-button
        class="btn-add"
        type="primary"
        size="mini"
        @click="scjszlPdb()"
        style="margin-left: 1px"
        v-show="type==2 || type == 4"
        >生成建设项目质量评定表</el-button
      >
      <el-button
        class="btn-add"
        type="primary"
        size="mini"
        @click="scBg()"
        style="margin-left: 1px"
        v-show="type==2 || type == 4"
        >生成报告</el-button
      >
      <!--:<el-button
        class="btn-add"
        type="success"
        size="mini"
        @click="gotoxm()"
        style="margin-left: 1px"
        icon="el-icon-link"
        
        >项目详情</el-button
      >
      disabled="$hasBP('bnt.project.remove')  === true"-->
    </div>
    </div>
    <!-- 表格 -->
    <el-table
      :data="list"
      border
      stripe
      :header-cell-style="{ 'text-align': 'center','color':'black' }"
      :cell-style="{ 'text-align': 'center' }"
      @selection-change="handleSelectionChange"
    >
      <el-table-column type="selection" />
      <el-table-column label="序号" width="102">
        <template slot-scope="scope">
          {{ (page - 1) * limit + scope.$index + 1 }}
        </template>
      </el-table-column>
      <el-table-column prop="proName" label="项目名称" width="162" />
      <el-table-column prop="xmqc" label="项目全称" width="162" />
      <el-table-column prop="grade" label="公路等级" />
      <el-table-column prop="participant" label="参与人员" />
      <el-table-column prop="tze" label="投资额（万元）" />
      <el-table-column prop="lxcd" label="路线长度（KM）" />
      <el-table-column prop="createTime" label="创建时间" />
      <el-table-column label="操作" align="center" fixed="right" >
          <template slot-scope="scope">
            <el-button type="warning" round  size="mini" @click="gotoxm(scope.row)" >项目详情</el-button>            
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
      <el-form ref="dataForm" :model="sysproject"  label-width="110px" size="small" style="padding-right: 40px;" :rules="rules">
        <el-form-item label="项目名称" prop="proName">
          <el-input v-model="sysproject.proName" />
        </el-form-item>
        <el-form-item label="项目全称" prop="xmqc">
          <el-input v-model="sysproject.xmqc" />
        </el-form-item>
        <el-form-item label="公路等级" prop="grade">
          <!-- <el-input v-model="sysproject.grade"/> -->
          <el-select style="width: 100%" v-model="sysproject.grade" placeholder="请选择">
              
              <el-option v-for="item in optionsdj"
                :key="item.id"
                :value="item.type"
                :label="item.type"></el-option>
                </el-select>
        </el-form-item>
        <el-form-item label="投资额(万元)" prop="tze">
          <el-input v-model="sysproject.tze"/>
        </el-form-item>
        <el-form-item label="路线长度(KM)" prop="lxcd">
          <el-input v-model="sysproject.lxcd"/>
        </el-form-item>
        <el-form-item label="参与人员" prop="participant" >
          <!-- <el-input v-model="htd.lx"/> -->
          <el-select
            style="width: 100%"
            v-model="sysproject.participant"
            multiple
            
            allow-create
            default-first-option
            placeholder="请选择参与人员"
          >
            <el-option
              v-for="item in options"
              :key="item.id"
              :label="item.name"
              :value="item.username"
            >
            </el-option>
          </el-select>
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
    <el-dialog title="修改" :visible.sync="dialogVisible1" width="40%" >
      <el-form ref="dataForm" :model="modifyData"  label-width="110px" size="small" style="padding-right: 40px;" :rules=rules1>
        <el-form-item label="项目全称" prop="xmqc">
         <el-input v-model="modifyData.xmqc"></el-input>
        </el-form-item>
        <el-form-item label="公路等级" prop="grade">
          
          <el-select style="width: 100%" v-model="modifyData.grade" placeholder="请选择">
            <el-option v-for="item in optionsdj"
                :key="item.id"
                :value="item.type"
                :label="item.type"></el-option>
            </el-select>
        </el-form-item>
        <el-form-item label="投资额(万元)" prop="tze">
         <el-input v-model="modifyData.tze"></el-input>
        </el-form-item>
        <el-form-item label="路线长度(KM)" prop="lxcd">
          <el-input v-model="modifyData.lxcd"/>
        </el-form-item>
        
      </el-form>
      <span slot="footer" class="dialog-footer">
        <el-button @click="dialogVisible1 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
        <el-button type="primary" icon="el-icon-check" @click="update()" size="small">确 定</el-button>
      </span>
    </el-dialog>
  </div>
</template>
<script>
// 引入定义接口js文件
import projectApi from '@/api/project/projectInfo.js'
import api from '@/api/project/htdUtil.js'
import FileSaver from "file-saver"
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
  data() {
    var validateProname = (rule, value, callback) => {
      
      //调用后台接口查询旧密码是否正确
      this.checkProname(
        value,
        function () {
          callback();
        },function (err) {
          callback(new Error("已存在的项目！"));
        }
      );
  };
    // 初始值
    return {
      list: [], // 项目列表
      total: 0, // 总记录数
      page: 1, // 当前页码
      limit: 10, // 每页记录数
      searchObj: {}, // 查询条件
      
      multipleSelection: [], // 批量删除选中的记录列表
      dialogVisible: false,
      dialogVisible1:false,
      sysproject: {
        
      },
      modifyData:{},

      username:Cookies.get('username'),
      userid:Cookies.get('userid'),
      type:Cookies.get('type'),
      companyId:Cookies.get('companyId'),
      optionsdj: [
          {
            id: '1',
            type: '高速公路'
          }, {
            id: '2',
            type: '一级公路'
          }, {
            id: '3',
            type: '二级公路'
          }, {
            id: '4',
            type: '三级公路'
          }, {
            id: '5',
            type: '四级公路'
          }
        ], //列表数据
      options:[],
      rules:{
            proName:[
              {
                required:true,
                message:"项目名称不能为空！"
              },
              { validator: validateProname, trigger: "blur" }
              
            ],
            xmqc:[
              {
                required:true,
                message:"项目全称不能为空！"
              },
              
            ],
            tze:[
              {
                required:true,
                message:"投资额不能为空！"
              },
              
            ],
            lxcd:[
              {
                required:true,
                message:"路线长度不能为空！"
              },
              
            ],
            grade:[
              {
                required:true,
                message:"公路名称不能为空！"
              },
              
            ],
            participant:[
              {required:true,message:"参与人员不能为空！"}
            ],
            createTime:[
              {
                required:true,
                message:"创建时间不能为空！"
              },
              
            ]
          },
      rules1:{
        xmqc:[
              {
                required:true,
                message:"项目全称不能为空！"
              },
              
            ],
            tze:[
              {
                required:true,
                message:"投资额不能为空！"
              },
              
            ],
            lxcd:[
              {
                required:true,
                message:"路线长度不能为空！"
              },
              
            ],
            grade:[
              {
                required:true,
                message:"公路名称不能为空！"
              },
              
            ],
      }
    }
  },
  created() {
    // 页面渲染之前
    
    
    this.fetchData()
    
  },
  methods: {
    formatType: function (row, column, cellValue) {
                var ret = ''  //你想在页面展示的值
                if (cellValue=='1') {
                    ret = "高速公路"  //根据自己的需求设定
                } else if(cellValue=='2') {
                    ret = "一级公路"
                }
                else if(cellValue=='3') {
                    ret = "二级公路"
                }
                else if(cellValue=='4') {
                    ret = "三级公路"
                }
                else if(cellValue=='5') {
                    ret = "四级公路"
                }
                return ret;
            },
    checkProname(proname,success, error) {
      //checkIsRepeat(oldpw) {
      
      if (!proname) return;
      
      
      
      projectApi.checkProname(proname).then((response) => {
        
          if(response.message == "校验成功"){
            success()
            
          }else if(response.message== "校验失败"){
            error()
            
            
          }
        })

    },
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
        projectApi.batchRemove(idList).then((response) => {
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
      this.getParticipant()

      
    },
    getParticipant(){
      projectApi.selectUser(this.companyId).then((response) => {
          this.options=response.data
        })
    },
    save() {
      this.$refs["dataForm"].validate((valid)=>{
              if(valid){
                let str='';
                for(let i=0;i<this.sysproject.participant.length;i++){
                  str+=this.sysproject.participant[i]+','
                }
                str=str.substring(0,str.length-1);
                this.sysproject.participant=str
                this.sysproject.userid=this.userid
                projectApi.saveProject(this.sysproject).then((response) => {
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
              }})
      
    },
    // 跳转
    gotoxm(row){
     
      let menus1=JSON.parse(JSON.stringify(global.antRouter))
      let str=[]
      if(menus1.map(v=>v.meta.title).includes("交工管理")){
            
            str=menus1[menus1.map(v=>v.meta.title).indexOf("交工管理")].children.filter(item=>item.meta.title==row.proName)
            
          }
      console.log("srt",menus1)
      console.log("str",str)
      this.$router.push({path:'/serviceProject/'+str[0].path+'/htdInfo',query:{projecttitle:row.proName}});
    },
    // 生成报告中表格
    scbgzBg(){
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择项目！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能选择一个项目！");
            return;
          }
        let commonInfoVo={}
        commonInfoVo.proname=this.multipleSelection[0].proName
        commonInfoVo.userid = this.userid;
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})

        api.scbgzBg(commonInfoVo).then((res) => {
          api.downloadbgzbg(commonInfoVo.proname).then((res) => {
            var temp = res.headers["content-disposition"]
            var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
            var matches = filenameRegex.exec(temp);
            var filename=''
            if (matches != null && matches[1]) { 
              filename = matches[1].replace(/['"]/g, '');
            }
            console.log('鉴定表',res.data)
            FileSaver.saveAs(res.data,decodeURI(filename))
            load.close()
          }).catch((err)=>{
              
              load.close()
            })
          
        }).catch((err)=>{
              
              load.close()
            })
    },
    // 生成建设项目质量评定表
    scjszlPdb(){
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择项目！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能选择一个项目！");
            return;
          }
      let commonInfoVo={}
        commonInfoVo.proname=this.multipleSelection[0].proName
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.scjszlPdb(commonInfoVo).then((res) => {
          api.downloadjsxm(commonInfoVo.proname).then((res) => {
            var temp = res.headers["content-disposition"]
            var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
            var matches = filenameRegex.exec(temp);
            var filename=''
            if (matches != null && matches[1]) { 
              filename = matches[1].replace(/['"]/g, '');
            }
            console.log('建设项目质量评定表',decodeURI(filename))
            FileSaver.saveAs(res.data,decodeURI(filename))
            load.close()
          }).catch((err)=>{
              
              load.close()
            })
          
            
            
          
          
        }).catch((err)=>{
              
              load.close()
            })
    },
    // 生成报告
    scBg(){
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择项目！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能选择一个项目！");
            return;
          }
        let commonInfoVo={}
        commonInfoVo.proname=this.multipleSelection[0].proName
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.scbg(commonInfoVo.proname).then((response) => {
          return response.data.text() 
        }).then(res=>{
                  
                  let code = JSON.parse(res).code
                  if(code == 200){
                    api.downloadbg(commonInfoVo.proname).then((res) => {
                        var temp = res.headers["content-disposition"]
                        var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                        var matches = filenameRegex.exec(temp);
                        var filename=''
                        if (matches != null && matches[1]) { 
                          filename = matches[1].replace(/['"]/g, '');
                        }
                        console.log('鉴定表',decodeURI(filename))
                        FileSaver.saveAs(res.data,decodeURI(filename))
                        load.close()
                      }).catch((err)=>{
                          
                          load.close()
                        })
                  }
                  else{
                    load.close()
                    let s=JSON.parse(res)
                    
                    this.$message.warning(s.message)
                    
                  }
                }).catch((err)=>{
              
                   load.close()
                })
          
            
            
          
          
        
    },
    
    fetchData() {
      // 调用api
      this.searchObj.username=this.username
      this.searchObj.userid=this.userid
      
      projectApi
        .pageList(this.page, this.limit, this.searchObj)
        .then((response) => {
          this.list = response.data.records
          this.total = response.data.total
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
      projectApi.getById(this.multipleSelection[0].id).then((res)=>{
         if(res.message=='成功'){
          this.dialogVisible1=true
          this.modifyData=res.data
          
         }
         else{
          this.$alert('出错')
         }
      })
      
    },
    update(){
      this.$refs["dataForm"].validate((valid)=>{
              if(valid){
                projectApi.modify(this.modifyData).then((res)=>{
                  if(res.message=='成功'){
                    this.$message({
                      type: "success",
                      message: "修改成功!",
                    });
                    // 刷新页面
                    this.dialogVisible1=false
                    
                    this.fetchData();
                  }
                  else{
                    this.$alert('出错')
                  }
                })
              }
            })
      
    },
  }
}
</script>
