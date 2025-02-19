<template>
  <div class="app-container">
    <div class="tools-div" v-show="!(username=='admin'&&type==1  )">
      <el-button size="small" type="primary" @click="exportlqs()" icon="el-icon-bottom" 
        >导出路桥隧模板文件</el-button
      >
      <el-button size="small" icon="el-icon-top" type="primary" @click="importlqs()" >导入路桥隧信息文件</el-button>
      
      <!-- <el-button size="small" type="primary"
        >删除所有导入的路桥隧数据</el-button
      > -->
    </div>

    <div class="tool-div">
      <el-col :span="8">
        <el-card shadow="hover" @click.native="ql()"> 桥梁清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="sd()"> 隧道清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="fhlm()"> 复合路面清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="hntlmjzd()"> 混凝土路面及匝道清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="sfz()"> 收费站清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="ljx()"> 连接线清单 </el-card>
      </el-col>
      <!-- <el-col :span="8">
        <el-card shadow="hover" @click.native="ljxql()"> 连接线桥梁清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="ljxsd()"> 连接线隧道清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="ljxhntlm()"> 连接线混凝土路面清单 </el-card>
      </el-col> 
      <el-col :span="8">
        <el-card shadow="hover" @click.native="xmjd()"> 项目进度 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="janum()"> 交安数量清单 </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" @click.native="ljnum()"> 路基附属清单 </el-card>
      </el-col>-->
    </div>

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
import projectApi from "@/api/project/projectInfo.js";
import api from "@/api/project/xmjd.js";
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
  data() {
    let projecttitle = this.$route.query.projecttitle;
    // 初始值
    return {
      data: {
        projectname: projecttitle,
      },
      list: [], // 项目列表
      dialogImportVisible: false,
      total: 0, // 总记录数
      page: 1, // 当前页码
      limit: 10, // 每页记录数
      searchObj: {}, // 查询条件
      multipleSelection: [], // 批量删除选中的记录列表
      dialogVisible: false,
      file:'',// 待上传文件
      fileList:[],
      sysproject: {},
      username:Cookies.get('username'),
      userid:Cookies.get('userid'),
      type:Cookies.get('type'),
    };
  },
  created() {
    // 页面渲染之前
    //this.fetchData();
  },
  methods: {
    importlqs() {
      this.dialogImportVisible = true;
    },
    fileChange(file,fileList){
        this.fileList=fileList
        
      },
      onUploadSuccess() {
        
        //debugger
        let fd=new FormData()
        
        
  
        fd.append('file',this.fileList[0].raw)
        fd.append('projectname',this.$route.query.projecttitle)
        fd.append('userid',this.userid)
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
        projectApi.importlqs(fd).then((res)=>{
          console.log('res',res)
          if(res.message=='成功'){
            
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
            });
        
        
        
      },
    ql() {
      let name = this.$route.query.projecttitle;
      //this.$router.push({ path:'/sysproject/projectInfo/lqs',name:name })
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/ql?name=" + name);
    },
    sd() {
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/sd?name=" + name);
    },
    fhlm() {
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/fhlm?name=" + name);
    },
    hntlmjzd(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/hntlmjzd?name=" + name);
    },
    sfz(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/sfz?name=" + name);
    },
    ljx(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/ljx?name=" + name);
    },
    ljxql(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/ljxql?name=" + name);
    },
    ljxsd(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/ljxsd?name=" + name);
    },
    ljxhntlm(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/ljxhntlm?name=" + name);
    },
    xmjd(){
      let name = this.$route.query.projecttitle;
      this.$router.push({path:this.$route.matched[1].path+"/xmjdmanage",query:{projecttitle:name}});
    },
    janum(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/janum?name=" + name);
    },
    ljnum(){
      let name = this.$route.query.projecttitle;
      this.$router.push(this.$route.matched[1].path+"/projectInfo/lqs/ljnum?name=" + name);
    },
    // 导出路桥隧模板文件
    exportlqs() {
      let projectname = this.$route.query.projecttitle;
      projectApi.exportAll(projectname).then((res) => {
        const objectUrl = URL.createObjectURL(
          new Blob([res.data], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          })
        );
        const link = document.createElement("a");
        // 设置导出的文件名称
        link.download = projectname+`路桥隧数据文件` + '.xlsx'
        link.style.display = "none";
        link.href = objectUrl;
        link.click();
        document.body.appendChild(link);
      })
      // let projectname = this.$route.matched[1].meta.title;
      // window.open(
      //   "http://localhost:8800/project/info/lqs/exportLqs?projectname=" +
      //     projectname
    }
  }
}
</script>

<style lang="scss">
.upload-demo {
  display: inline-block;
  margin-left: 3px;
  margin-right: 3px;
}
.margin-change {
  display: inline-block;
  margin-left: 10px;
}
.tool-div {
  height: 200px;
}

</style>
