<template>
  <div class="app-container">
      <div class="tools-div">
      <div class="anniu">
              <el-button
              class="btn-add"
              type="success"
              @click="download()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >下载选中文件</el-button>
              <el-button
              class="btn-add"
              type="success"
              @click="getxm()"
              style="margin-left: 5px"
              icon="el-icon-top"
              size="mini"
              >上传文件</el-button>
              <el-button
              class="btn-add"
              type="warning"
              @click="deleteFile()"
              style="margin-left: 5px"
              icon="el-icon-delete"
              size="mini"
              >删除文件</el-button>     
      </div>  
       
    </div>
    <el-row>
      
        <el-card shadow="never"  class="content" align="middle" > 所有文件 </el-card>
      
    </el-row>
    <el-row>
      <!--<el-col :span="3">
        <el-input
          placeholder="输入关键字搜索"
          v-model="filterText">
        </el-input>
      </el-col>-->
      <el-col :span="24">
        <div class="tree-main">
          <el-tree
          ref="tree"
          style="font-size: 20px;"
          :data="fileList"
          :props="defaultProps"
          show-checkbox
          check-strictly
          :router="true"
          @node-click='changtree'
          current-node-key="id"
          highlight-current
          :filter-node-method="filterNode"
          >
          <span class="custom-tree-node" slot-scope="{ node, data }">
              <i v-if="data.isDir==true&&!node.expanded"  class="el-icon-folder"></i>
              <i v-else-if="node.expanded" class="el-icon-folder-opened"></i>
              <i v-else  class="el-icon-document"></i>
              <span >{{node.label}}</span>
          </span>

        </el-tree>
         
          
          
      </div>
      </el-col>
    </el-row>
    <el-dialog title="上传" :visible.sync="dialogImportVisible" width="480px">
        <el-form ref="dataForm" :model="fileproject" label-position="right"   label-width="170px" :rules="rules">
          <el-form-item label="选择项目" prop="proname" >
            <!-- <el-input v-model="htd.lx"/> -->
            <el-select
              style="width: 100%"
              v-model="fileproject.proname"
             
              default-first-option
              placeholder="请选择要导入的项目"
              
            >
              <el-option
                v-for="item in options"
                :key="item.id"
                :label="item.proName"
                :value="item.proName"
              >
              </el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="文件">
            <el-upload
              :multiple="false"
              :on-change="fileChange"
              :auto-upload="false"
              :file-list="fileList1"
              :limit="1"
              action="/"
              class="upload-demo"
            >
              <el-button size="small" type="primary">点击上传</el-button>
              <!--<div slot="tip" class="el-upload__tip">只能上传xls文件</div>-->
            </el-upload>
          </el-form-item>

        </el-form>
        <div slot="footer" class="dialog-footer">
          <el-button @click="onUploadSuccess">提交</el-button>
          <el-button @click="dialogImportVisible = false,fileList1=[],proname=''">取消</el-button>
        </div>
      </el-dialog>
    
    
    

    
  </div>
</template>
<script>
// 引入定义接口js文件
import api from "@/api/project/file/file.js";
import { Loading } from 'element-ui';
import FileSaver from "file-saver"
import Cookies from 'js-cookie'
export default {
  data() {
    // 初始值
    let value = this.$route.query.name;
    return {
      fileproject:{},
      dialogImportVisible: false,
      fileList: [], // 项目列表
      defaultProps: {
        children: 'children',
        label: 'fileName'
      },
      userid:Cookies.get('userid'),
      multipleSelection:[],
      filterText:"",
      file:'',// 待上传文件
      fileList1:[],
      options:[
        
      ],
      rules:{
        proname:[
          {
            required:true,
            message:"请选择要导入的项目！"
          }
        ]
      }

    };
  },
  created() {
    // 页面渲染之前
    console.log(this.$route.query.name);
    //alert(this.$route.matched[1].meta.title)
    this.fetchData();
  },
  watch: {
    filterText(val) {
      this.$refs.tree.filter(val);
    },
  },
  methods: {
    changtree(){

    },

    getxm(){
      this.dialogImportVisible=true
      api.getxm(this.userid).then((response) => {
                  console.log("sqwe",response)
                  this.options=response.data
              })
    },
    // 上传文件触发
    fileChange(file,fileList){
        this.fileList1=fileList
        
      },
    onUploadSuccess() {
      this.$refs["dataForm"].validate((valid)=>{
            if(valid){
              let fd=new FormData()
        
        
  
              fd.append('file',this.fileList1[0].raw)
              fd.append('proName',this.fileproject.proname)
              
              
              let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
              api.upload(fd).then((res)=>{
                console.log('res',res)
                if(res.message=='成功'){
                  this.fetchData();
                  this.$message({
                    type: "success",
                    message: "上传成功!",
                  });
                  this.dialogImportVisible = false;
                  this.fileList1=[]
                  this.proname=''
                }
                else{
                  this.$alert('上传失败！')
                }
                load.close()
              }).catch((err)=>{
                  this.dialogImportVisible=false
                  load.close()
                })
            }
            else{
              console.log("error")
            }
          })
        
        
        
        
        
      },
    filterNode(value,data){
      if(!value) this.fetchData();
      
      return  data.name.indexOf(value)!==-1;
    },
    download() {
      this.multipleSelection=this.$refs.tree.getCheckedNodes()
      console.log("文件",this.multipleSelection)
      
      if (this.multipleSelection.length === 0) {
          this.$message.warning("请选择要下载的文件！");
          return;
      }

      else if(this.multipleSelection.length>1){
          var flag=true;
          for(let i=0;i<this.multipleSelection.length;i++){
              if(this.multipleSelection[i].children!=null){
                  flag=false;
                  break;
              }
          }
          if( flag==false){
              this.$message.warning("一次只能选择一个文件夹或者多个文件下载！");
              return;
          }
          else{
              api.download(this.multipleSelection).then((response) => {
                  var temp = response.headers["content-disposition"]
                  var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                  var matches = filenameRegex.exec(temp);
                  var filename=''
                  if (matches != null && matches[1]) { 
                      filename = matches[1].replace(/['"]/g, '');
                  }
                  console.log('鉴定表',decodeURI(filename))
                  FileSaver.saveAs(response.data,decodeURI(filename))
                  
              })
          }
      }
      else if(this.multipleSelection.length==1){
          api.download(this.multipleSelection).then((response) => {
                  var temp = response.headers["content-disposition"]
                  var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                  var matches = filenameRegex.exec(temp);
                  var filename=''
                  if (matches != null && matches[1]) { 
                      filename = matches[1].replace(/['"]/g, '');
                  }
                  console.log('鉴定表',response)
                  FileSaver.saveAs(response.data,decodeURI(filename))
                  
              })
      }
          
      
      
    },
    deleteFile(){
      this.multipleSelection=this.$refs.tree.getCheckedNodes()
     
      
      if (this.multipleSelection.length === 0) {
          this.$message.warning("请选择要删除的文件！");
          return;
      }

      else if(this.multipleSelection.length>1){
          var flag=true;
          for(let i=0;i<this.multipleSelection.length;i++){
              if(this.multipleSelection[i].children!=null){
                  flag=false;
                  break;
              }
          }
          if( flag==false){
              this.$message.warning("一次只能选择一个文件夹或者多个文件删除！");
              return;
          }
          else{
            this.$confirm("此操作将删除选中的所有文件, 是否继续?", "提示", {
              confirmButtonText: "确定",
              cancelButtonText: "取消",
              type: "warning",
            }).then(()=>{
              api.deleteFile(this.multipleSelection).then((response) => {
                
                this.$message({
                  type: "success",
                  message: "删除成功!",
                });
                  location.reload()
                })
            })
              
          }
      }
      else if(this.multipleSelection.length==1){
        this.$confirm("此操作将删除选中的所有文件, 是否继续?", "提示", {
              confirmButtonText: "确定",
              cancelButtonText: "取消",
              type: "warning",
            }).then(()=>{
              api.deleteFile(this.multipleSelection).then((response) => {
                this.$message({
                  type: "success",
                  message: "删除成功!",
                });
                location.reload()
                })
            })
      }
    },
    
    fetchData() {
      
      api.getfilelist(this.userid).then((response) => {
        this.fileList = response.data
        
        console.log("sqsd",this.fileList)
      });
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
.tree-main{
  font-size: 50px;
}
</style>
