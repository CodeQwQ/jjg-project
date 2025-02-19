<template>
    <div class="app-container">
      <div class="tools-div" >
        <el-button size="small" type="primary" @click="back()" 
          >返回</el-button>
        <el-button size="small" type="primary" v-show="!(username=='admin'&&type==1  )" @click="exportlqs()" icon="el-icon-bottom"
          >导出路桥隧模板文件</el-button
        >
        <el-button size="small" icon="el-icon-top" v-show="!(username=='admin'&&type==1  )" type="primary" @click="importlqs()">导入路桥隧信息文件</el-button>
        
        <!-- <el-button size="small" type="primary"
          >删除所有导入的路桥隧数据</el-button
        > -->
      </div>
  
      <div class="tool-div">
      <el-row >  
          <el-card  class="row-title"  > 路桥隧信息 </el-card> 

          
        </el-row>
        <el-row>
          <el-col :span="8">
                <el-card shadow="hover" @click.native="htd()"> 合同段信息 </el-card>
            </el-col>
            
            <el-col :span="8">
                <el-card shadow="hover" @click.native="ql()"> 桥梁清单 </el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="sd()"> 隧道清单 </el-card>
            </el-col>
            
        </el-row>
        <el-row>
          <el-col :span="8">
                <el-card shadow="hover" @click.native="ljx()"> 连接线清单 </el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="hntlmjzd()"> 混凝土路面及匝道清单 </el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="sfz()"> 收费站清单 </el-card>
            </el-col>
            
            
        </el-row>

         <el-row >  
          <el-card  class="row-title"  > 其他信息 </el-card> 
        </el-row>

        <el-row>

            <el-col :span="8">
                <el-card shadow="hover" @click.native="ryxxjgfc()"> 人员信息 </el-card>
            </el-col>
        </el-row>
        <el-row >  
          <el-card  class="row-title"  > 复测指标 </el-card> 
        </el-row>
        <el-row>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="cz()"> 车辙 </el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="xlbgc()"> 混凝土路面相邻板高差 </el-card>
            </el-col>
            
            <el-col :span="8">
                <el-card shadow="hover" @click.native="mcxs()"> 摩擦系数 </el-card>
            </el-col>
            
        </el-row>
        <el-row>
          <el-col :span="8">
                <el-card shadow="hover" @click.native="gzsdsgpsf()"> 构造深度(手工铺砂法) </el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="gzsdjqjcf()"> 构造深度(机器检测法) </el-card>

            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="pzd()"> 平整度 </el-card>
            </el-col>
        </el-row>
        <el-row>
          <el-col :span="8">
                <el-card shadow="hover" @click.native="jglmwcbkml()"> 路面弯沉(贝克曼梁法) </el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="jglmwclcf()"> 路面弯沉(落锤法)</el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="jabx()"> 交安标线</el-card>
            </el-col>
            <el-col :span="8">
                <el-card shadow="hover" @click.native="jafhl()"> 交安防护栏</el-card>
            </el-col>
        </el-row>
        <el-row>
            <div class="tools-div1">
              <el-button size="medium" type="primary"  @click="jgjcsj()" >交工检测数据</el-button>
            
            
        </div>
        </el-row>
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
  
  import api from "@/api/jgproject/new.js";
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
        proname:this.$route.query.projecttitle,
        username:Cookies.get('username'),
        userid:Cookies.get('userid'),
        type:Cookies.get('type'),
        file:'',// 待上传文件
        fileList:[],
        sysproject: {},
      };
    },
    created() {
      // 页面渲染之前
      //this.fetchData();
    },
    methods: {
      back(){
        this.$router.go(-1)
      },
      
      htd() {
        let proname = this.$route.query.projecttitle;
        
        this.$router.push({path:"/jgproject/guifan/htdinfo",query:{projecttitle:proname}});
      },
      dwgctze() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/dwgctze",query:{projecttitle:proname}});
      },
      ql() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/qlqd",query:{projecttitle:proname}});
      },
      sd() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/sdqd",query:{projecttitle:proname}});
      },
      ljx(){
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/ljxqd",query:{projecttitle:proname}});
      },
      hntlmjzd(){
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/hntlmjzd",query:{projecttitle:proname}});
      },
      sfz(){
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/sfzqd",query:{projecttitle:proname}});
      },
      wgkf() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/wgkf",query:{projecttitle:proname}});
      },
      nyzlkf() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/nyzlkf",query:{projecttitle:proname}});
      },
      cz() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/cz",query:{projecttitle:proname}});
      },
      xlbgc() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/hntlmxlbgc",query:{projecttitle:proname}});
      },
      mcxs() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/mcxs",query:{projecttitle:proname}});
      },
      gzsdsgpsf() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/gzsdsgpsf",query:{projecttitle:proname}});
      },
      gzsdjqjcf() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/gzsdjqjcf",query:{projecttitle:proname}});
      },
      pzd() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/pzd",query:{projecttitle:proname}});
      },

      jglmwcbkml() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/jglmwcbkml",query:{projecttitle:proname}});
      },
      jglmwclcf() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/jglmwclcf",query:{projecttitle:proname}});
      },
      jabx() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/jabx",query:{projecttitle:proname}});
      },
      jafhl() {
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/jafhl",query:{projecttitle:proname}});
      },
      
      
      
      jgjcsj(){
        let proname = this.$route.query.projecttitle;
        this.$router.push({path:"/jgproject/guifan/jgjcsj",query:{projecttitle:proname,guifan:'new'}});
      },
      importlqs() {
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
          fd.append('projectname',this.proname)
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importnew(fd).then((res)=>{
            console.log('res',res)
            if(res.status=='200'){
              
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
      // 导出路桥隧模板文件
      exportlqs() {
        let projectname = this.$route.query.projecttitle;
        api.exportnew().then((res) => {
          const objectUrl = URL.createObjectURL(
            new Blob([res.data], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            })
          );
          const link = document.createElement("a");
          // 设置导出的文件名称
          link.download = projectname+`路桥隧数据文件(新)` + '.xlsx'
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
  .tools-div1{
    margin-top: 20px;
  }
  .row-title {
      margin-top: 20px;
      background: rgb(210, 238, 238);
      font-weight: bold;
      
      
    }
  </style>
  