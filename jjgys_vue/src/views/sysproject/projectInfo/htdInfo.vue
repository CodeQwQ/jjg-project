<template>
  <div class="app-container">
    <!--查询表单-->
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
    </div>
    <!-- 工具按钮 -->
    <div class="tools-div" v-show="!(username=='admin'&&type==1  )" >
      <div class="anniu">
        <el-button
          class="btn-add"
          type="success"
          @click="add()"
          style="margin-left: 0px"
          icon="el-icon-plus"
          size="mini"
          
          >添加</el-button
        >
        <!--:disabled="$hasBP('bnt.htd.add') === true"-->
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
        <!--:disabled="$hasBP('bnt.htd.remove') === true"-->
        <el-button
          class="btn-add"
          type="primary"
          size="mini"
          @click="exporthtd()"
          style="margin-left: 1px"
          >批量导出实测数据模板文件</el-button
        >

        <el-button
          class="btn-add"
          type="primary"
          size="mini"
          @click="checkSelect()"
          style="margin-left: 1px"          
          >批量导入实测数据文件</el-button
        >
        <el-button
          class="btn-add"
          type="primary"
          size="mini"
          @click="scjdb()"
          style="margin-left: 1px"
          >生成鉴定表</el-button
        >
         <el-button
          class="btn-add"
          type="primary"
          size="mini"
          @click="scpdb()"
          style="margin-left: 1px"          
          >生成评定表</el-button
        >
        
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
      <el-table-column prop="name" label="合同段名称" />
      <el-table-column prop="sgdw" label="施工单位" />
      <el-table-column prop="jldw" label="监理单位" />
      <el-table-column prop="gls" label="公里数(KM)" />
      <el-table-column prop="tze" label="投资额(万元)" />
      <!-- <el-table-column prop="zy" label="ZY" /> -->
      <el-table-column prop="zhq" label="左幅桩号起" />
      <el-table-column prop="zhz" label="左幅桩号止" />
      <el-table-column prop="yfzhq" label="右幅桩号起" />
      <el-table-column prop="yfzhz" label="右幅桩号止" />
      <el-table-column prop="lx" label="合同段类型" />
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

    <el-dialog title="添加" :visible.sync="dialogVisible" width="40%">
      <el-form
        ref="dataForm"
        :model="htd"
        label-width="100px"
        size="small"
        :rules="htdrules"
        style="padding-right: 40px"
      >
        <el-form-item label="合同段名称" prop="name">
          <el-input v-model="htd.name" />
        </el-form-item>
        <el-form-item label="施工单位" prop="sgdw">
          <el-input v-model="htd.sgdw" />
        </el-form-item>
        <el-form-item label="监理单位" prop="jldw">
          <el-input v-model="htd.jldw" />
        </el-form-item>
        <!-- <div style="font-size: 11px;color:blue;margin-left: 100px;margin-bottom: 10px;">若左右幅桩号起止相同，选择Z&&Y，左右幅的起止桩号填相同即可；否则选择Z||Y，分别填写不同的起止桩号。</div>
        <el-form-item label="ZY" prop="zy">
          <el-select v-model="zy" style="width: 100%">
            <el-option v-for="item in zys"
            :key="item.value"
            :label="item.label"
            :value="item.value"
            >
            </el-option>
          </el-select>
        </el-form-item> -->
        <el-form-item label="左幅桩号起" prop="zhq">
          <el-input v-model="htd.zhq" />
        </el-form-item>

        <el-form-item label="左幅桩号止" prop="zhz">
          <el-input v-model="htd.zhz" />
        </el-form-item>
        <el-form-item label="右幅桩号起" prop="yfzhq">
          <el-input v-model="htd.yfzhq" />
        </el-form-item>

        <el-form-item label="右幅桩号止" prop="yfzhz">
          <el-input v-model="htd.yfzhz" />
        </el-form-item>

        <el-form-item label="公里数(KM)" prop="gls">
          <el-input v-model="htd.gls" />
        </el-form-item>
        <el-form-item label="投资额(万元)" prop="tze">
          <el-input v-model="htd.tze" />
        </el-form-item>
        
        <el-form-item label="合同段类型" prop="lx" >
          <!-- <el-input v-model="htd.lx"/> -->
          <el-select
            style="width: 100%"
            v-model="htd.lx"
            multiple
            allow-create
            default-first-option
            placeholder="请选择合同段类型"
          >
            <el-option
              v-for="item in options"
              :key="item.value"
              :label="item.label"
              :value="item.value"
            >
            </el-option>
          </el-select>
        </el-form-item>
        <el-form-item label="创建时间" prop="createTime">
          <el-date-picker
            v-model="htd.createTime"
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
          @click="save()"
          size="small"
          >确 定</el-button
        >
      </span>
    </el-dialog>
    <el-dialog title="导入" :visible.sync="dialogImportVisible" width="480px">
      <el-form label-position="right" label-width="110px">
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
            <div slot="tip" class="el-upload__tip" style=color:red>只能上传zip文件</div>
            <div slot="tip" class="el-upload__tip" style=color:red>请确保你的压缩包层级是：实测数据\合同段名称\分部工程</div>
            <div slot="tip" class="el-upload__tip" style=color:red>例如：京昆高速实测数据模板文件\JK-C01</div>
          </el-upload>
        </el-form-item>
      </el-form>
      <div slot="footer" class="dialog-footer">
        <el-button @click="onUploadSuccess">提交</el-button>
        <el-button @click="dialogImportVisible = false,fileList=[]">取消</el-button>
      </div>
    </el-dialog>
    <el-dialog title="修改" :visible.sync="dialogVisible1" width="40%" >
      <el-form ref="dataForm" :model="modifyData"  label-width="100px" size="small" style="padding-right: 40px;" :rules=rules1>
        <el-form-item label="施工单位" prop="sgdw">
         <el-input v-model="modifyData.sgdw"></el-input>
        </el-form-item>
        <el-form-item label="监理单位" prop="jldw">
         <el-input v-model="modifyData.jldw"></el-input>
        </el-form-item>
        <!-- <div style="font-size: 11px;color:blue;margin-left: 100px;margin-bottom: 10px;">若左右幅桩号起止相同，选择Z&&Y，左右幅的起止桩号填相同即可；否则选择Z||Y，分别填写不同的起止桩号。</div>
        <el-form-item label="ZY" prop="zy">
          
          <el-select v-model="modifyData.zy" style="width: 100%">
            <el-option v-for="item in zys"
            :key="item.value"
            :label="item.label"
            :value="item.value"
            >

            </el-option>
          </el-select>
        </el-form-item> -->
        <el-form-item label="左幅桩号起" prop="zhq">
          <el-input v-model="modifyData.zhq"/>
        </el-form-item>
        <el-form-item label="左幅桩号止" prop="zhz">
          <el-input v-model="modifyData.zhz"/>
        </el-form-item>

        <el-form-item label="右幅桩号起" prop="yfzhq">
          <el-input v-model="modifyData.yfzhq"/>
        </el-form-item>
        <el-form-item label="右幅桩号止" prop="yfzhz">
          <el-input v-model="modifyData.yfzhz"/>
        </el-form-item>

        <el-form-item label="投资额(万元)" prop="tze">
         <el-input v-model="modifyData.tze"></el-input>
        </el-form-item>
        <el-form-item label="公里数(KM)" prop="gls">
          <el-input v-model="modifyData.gls"/>
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
import api from "@/api/project/htd.js";
import api1  from "@/api/project/htdUtil.js";

import FileSaver from "file-saver"
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
  data() {
    var validateHtdName = (rule, value, callback) => {
      
        //调用后台接口查询旧密码是否正确
        this.checkHtdName(
          value,
          function () {
            callback();
          },function (err) {
            callback(new Error("已存在的合同段！"));
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
      dialogImportVisible:false,
      username:Cookies.get("username"),
      userid:Cookies.get("userid"),
      type:Cookies.get("type"),
      companyId:Cookies.get("companyId"),
      modifyData:{},
      file:'',// 待上传文件
      fileList:[],
      fileflag:false,
      htdrules:{
         name: [
         { required: true, message: '请输入合同段名称' },
         { validator: validateHtdName, trigger: "blur" }
         
        ],
        // zys: [
        //   { required: true, message: '请选择左右幅' }
        // ],
        sgdw: [
          { required: true, message: '请输入施工单位' }
        ],
        jldw: [
          { required: true, message: '请输入监理单位' }
        ],
        zhq: [
         { required: true, message: '请输入左幅桩号起' }
         
        ],
        zhz: [
         { required: true, message: '请输入左幅桩号止' }
         
        ],
        yfzhq: [
         { required: true, message: '请输入右幅桩号起' }
         
        ],
        yfzhz: [
         { required: true, message: '请输入右幅桩号止' }
         
        ],
        // gls: [
        //   { required: true, message: '请输入公里数' }
        // ],
        // tze: [
        //   { required: true, message: '请输入投资额' }
        // ],
        zh: [
          { required: true, message: '请输入桩号' }
        ],
        lx: [
          { required: true, message: '请选择合同段类型' }
        ],
        createTime: [
          { required: true, message: '请选择创建时间' }
        ]
      },
      rules1:{
        // zy: [
        //   { required: true, message: '请选择左右幅' }
        // ],
        sgdw: [
          { required: true, message: '请输入施工单位' }
        ],
        jldw: [
          { required: true, message: '请输入监理单位' }
        ],
        zhq: [
         { required: true, message: '请输入左幅桩号起' }
         
        ],
        zhz: [
         { required: true, message: '请输入左幅桩号止' }
         
        ],
        yfzhq: [
         { required: true, message: '请输入右幅桩号起' }
         
        ],
        yfzhz: [
         { required: true, message: '请输入右幅桩号止' }
         
        ],
        // gls: [
        //   { required: true, message: '请输入公里数' }
        // ],
      },
      htd: {
        name: '',
        sgdw: '',
        jldw: '',
        zh: '',
        createTime: '',
        lx: '',
        proname:'',
        //zy:""
      },
      // zy:"",
      // zys:[
      //   {
      //     value: 'Z||Y',
      //     label: "Z||Y"
      //   },
      //   {
      //     value: "Z&&Y",
      //     label: "Z&&Y"
      //   }
      // ],
      options: [
        {
          value: "路基工程",
          label: "路基工程",
        },
        {
          value: "路面工程",
          label: "路面工程",
        },
        {
          value: "交安工程",
          label: "交安工程",
        },
        {
          value: "桥梁工程",
          label: "桥梁工程",
        },
        {
          value: "隧道工程",
          label: "隧道工程",
        }
      ]
    }
  },
  created() {
    
    this.fetchData();
  },
  methods: {
    checkHtdName(htd,success, error) {
      //checkIsRepeat(oldpw) {
      
      if (!htd) return;
      let commonInfoVo={}
      commonInfoVo.proname=this.$route.query.projecttitle
      commonInfoVo.htd=htd
      api.checkProname(commonInfoVo).then((response) => {
        
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
    // 复选框发生变化，调用方法，选中复选框行内容传递
    handleSelectionChange(selection) {
      this.multipleSelection = selection;
    },
    // 跳转到添加的表单页面
    // add() {
    //   this.$router.push({ path: '/project/create' })
    // },
    add() {
      this.dialogVisible = true;
      this.htd = {};
    },
    save() {
      
      this.$refs["dataForm"].validate((valid)=>{
              if(valid){
                this.htd.lx = JSON.stringify(this.htd.lx)
                this.htd.zy=this.zy
                this.htd.proname = this.$route.query.projecttitle
                this.htd.companyId=this.companyId
                api.save(this.htd).then((response) => {
                this.list = response.data;
                // 提示
                this.$message({
                  type: "success",
                  message: "添加成功!",
                });
                // 关闭弹框
                this.dialogVisible = false;
                // 刷新
                this.fetchData();
                location.reload();
              });
              }
              
            
            })
     
    },
    exporthtd() {
        if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要导出模板的合同段！");
            return;
          }
          let projectname = this.$route.query.projecttitle;
          let htds=[];
          for(let i=0;i<this.multipleSelection.length;i++){
            //console.log('test',this.multipleSelection[i])
            htds.push(this.multipleSelection[i].name);
          }
          /*let htd =this.htdname*/
          
          
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
          api.exporthtd(projectname,htds).then((response) => {
              var temp = response.headers["content-disposition"]
              var filename=temp.split('=')[1]
              filename=filename.replace(/['"]/g, '')
              
              FileSaver.saveAs(response.data,decodeURI(filename))
              load.close()
          }).catch((err)=>{
              
              load.close()
            });
    },
    // 检查复选框
    checkSelect(){
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要导入数据的合同段！");
            return;
          }
          this.dialogImportVisible=true
    },

     // 上传文件触发
     fileChange(file,fileList){
          let filename=file.name
          let pos=filename.lastIndexOf(".")
          let lastname=filename.substring(pos,filename.length)
          if(lastname==".zip"){
            
            this.fileflag=true
            this.fileList=fileList
          }
          else{
            this.fileflag=false
            this.fileList=[]
          }
        
        
      },
     onUploadSuccess(response, file) {
        if(this.fileflag==false){
            this.$message("只能上传zip文件！")
        }
        else{
          let fd=new FormData()
        let projectname = this.$route.query.projecttitle;
        console.log(this.fileList[0],"ddd")

        fd.append('file',this.fileList[0].raw)
        fd.append('proname',projectname)
        fd.append('userid',this.userid)
        let htds=[];
        for(let i=0;i<this.multipleSelection.length;i++){
            //console.log('test',this.multipleSelection[i])
            htds.push(this.multipleSelection[i].name);
            //fd.append('htd',htds[i])
        }
        
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.importhtd(fd).then((res)=>{
          console.log('res',res)
          if(res.message=='成功'){
            
            this.$message({
              type: "success",
              message: "上传成功!",
            });
            this.dialogImportVisible = false;
            this.fileList=[]
            //this.fetchData()
          }
          else{
            this.$alert('上传失败！')
          }
          load.close()
        }).catch((err)=>{
              
              load.close()
            });
        }
        
    },
    scjdb(){
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要生成鉴定表的合同段！");
            return;
          }
        
        let commonInfoVo={}
        commonInfoVo.proname=this.$route.query.projecttitle
        commonInfoVo.userid=this.userid
        let htds=[];
        for(let i=0;i<this.multipleSelection.length;i++){
            
            htds.push(this.multipleSelection[i].name);
            
        }
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
        api.scjdb({proname:commonInfoVo.proname,htds:htds,userid:this.userid}).then((response) => {
                  return response.data.text()       
                })
                .then(res=>{
                  
                  let code = JSON.parse(res).code
                  if(code == 200){
                    api.download(commonInfoVo.proname,htds,this.userid).then((response)=>{
                    
                    var temp = response.headers["content-disposition"]
                    var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    var matches = filenameRegex.exec(temp);
                    var filename=''
                    if (matches != null && matches[1]) { 
                      filename = matches[1].replace(/['"]/g, '');
                    }
                    console.log('鉴定表',response)
                    FileSaver.saveAs(response.data,decodeURI(filename))
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
                })
                .catch((err)=>{
                        console.log(err)
                        load.close()
                      });
        
      },
      scpdb(){
        if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要生成鉴定表的合同段！");
            return;
        }
        
        let commonInfoVo={}
        let htds=[]
        commonInfoVo.proname=this.$route.query.projecttitle
        commonInfoVo.userid=this.userid
        for(let i=0;i<this.multipleSelection.length;i++){
            
            htds.push(this.multipleSelection[i].name);
            
        }
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api1.scpdb({proname:commonInfoVo.proname,htds:htds}).then((response) => {
          
          api1.downloadpdb({proname:commonInfoVo.proname,htds:htds}).then((res=>{
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
          })).catch((err)=>{
              
              load.close()
            })
            
          
          
        }).catch((err)=>{
              
              //load.close()
            });;

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
                api.modify(this.modifyData).then((res)=>{
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
  },
};
</script>

