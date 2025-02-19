<template>
<div class="app-container">
    <!--<div class="tools-div" v-show="!(username=='admin'&&type==1)">
       <el-button size="small" type="primary" @click="exportlmgc()" icon="el-icon-bottom"
        >导出路面工程模板文件</el-button
      > -->
      <!-- <el-button size="small" icon="el-icon-top" type="primary" @click="importlmgc()">导入路面工程信息文件</el-button>
      <el-button size="small" type="primary" icon="el-icon-bottom" @click="scjdb()"
        >生成路面工程所有鉴定表</el-button
      > 
    </div>-->

    <div class="tool-div">
      <el-row >  
        <el-card  class="row-title"  > 路面面层 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="5">
          <el-card shadow="hover"  class="content" align="middle" @click.native="lqlmysd()"> 沥青路面压实度 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover"  class="content" align="middle" @click.native="lmwcbkmlf()"> 路面弯沉贝克曼梁法 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover"  class="content" align="middle" @click.native="lmwclcf()"> 路面弯沉落锤法 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover"  class="content" align="middle" @click.native="lqlmssxs()"> 沥青路面渗水系数 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover"  class="content" align="middle" @click.native="hntlmqd()"> 混凝土路面强度 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover"  class="content" align="middle" @click.native="tlmxlbgc()"> 砼路面相邻板高差 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover"  class="content" align="middle" @click.native="lmgzsd()"> 路面构造深度手工铺砂法 </el-card>
        </el-col>
      <el-col :span="6">
        <el-card shadow="hover"  class="content" align="middle" @click.native="gslqlmhdzx()"> 高速沥青路面厚度钻芯法 </el-card>
      </el-col>
      <el-col :span="7">
        <el-card shadow="hover"  class="content" align="middle" @click.native="hntlmhdzx()"> 混凝土路面厚度钻芯法 </el-card>
      </el-col>
      <el-col :span="5">
        <el-card shadow="hover"  class="content" align="middle" @click.native="lmhp()"> 路面横坡 </el-card>
      </el-col>
      </el-row>
     <el-row>
      
      <!-- <el-col :span="4">
        <el-card shadow="hover"  class="content" align="middle" @click.native="xywzx()"> 芯样完整性 </el-card>
      </el-col>
      <el-col :span="4">
        <el-card shadow="hover"  class="content" align="middle" @click.native="hd()"> 厚度 </el-card>
      </el-col> -->
     </el-row>
     <el-row >  
        <el-card  class="row-title"  > 自动化 </el-card> 
      </el-row>
      <el-row>
      <el-col :span="4">
          <el-card shadow="hover"  class="content" align="middle" @click.native="mcxs()"> 摩擦系数 </el-card>
        </el-col>
      <el-col :span="4">
        <el-card shadow="hover"  class="content" align="middle" @click.native="gzsd()"> 构造深度 </el-card>
      </el-col>
      <el-col :span="4">
        <el-card shadow="hover"  class="content" align="middle" @click.native="pzd()"> 平整度 </el-card>
      </el-col>
      <el-col :span="4">
        <el-card shadow="hover"  class="content" align="middle" @click.native="cz()"> 车辙 </el-card>
      </el-col>
      <el-col :span="4">
        <el-card shadow="hover"  class="content" align="middle" @click.native="ldhd()"> 雷达厚度 </el-card>
      </el-col>
      
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
            <div slot="tip" class="el-upload__tip">只能上传zip文件</div>
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
import api from '@/api/project/fbgc/lmgc.js'
import FileSaver from "file-saver"
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
    name:'lmlist',
    data() {
        let projecttitle = this.$route.matched[1].meta.title;
        return {
            data: {
        projectname: projecttitle,
      },
            dialogImportVisible: false,
            htdname:this.$route.query.htdname,
            fbgcName:this.$route.query.fbgcName,
            file:'',// 待上传文件
            fileList:[],
            username:Cookies.get('username'),
            userid:Cookies.get('userid'),
            type:Cookies.get('type'),
        }
    },
    methods: {
    lqlmysd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/lqlmysd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lmwcbkmlf() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/lmwcbkmlf",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lmwclcf() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/lmwclcf",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lqlmssxs() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/lqlmssxs",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    hntlmqd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/hntlmqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    tlmxlbgc() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/tlmxlbgc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lmgzsd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/lmgzsd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    gslqlmhdzx() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/gslqlmhdzx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    hntlmhdzx() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/hntlmhdzx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lmhp() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/lmhp",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    xywzx(){

    },
    hd(){

    },
    mcxs() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/mcxs",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    gzsd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/gzsd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    pzd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/pzd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    cz() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/cz",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    ldhd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/lmgc/ldhd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    importlmgc() {
      this.dialogImportVisible = true;
    },
    // 上传文件触发
    fileChange(file,fileList){
        this.fileList=fileList
        
        
      },
     onUploadSuccess(response, file) {
        let fd=new FormData()
        let projectname = this.$route.query.projecttitle;
        let htd =this.htdname
        

        fd.append('file',this.fileList[0].raw)
        fd.append('proname',projectname)
        fd.append('htd',htd)
        
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.importlmgc(fd).then((res)=>{
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
    exportlmgc() {
          let projectname = this.$route.query.projecttitle;
          let htd =this.htdname
          
          console.log('test')
          
          api.exportlmgc().then((response) => {
              var temp = response.headers["content-disposition"]
              var filename=temp.split('=')[1]
              filename=filename.replace(/['"]/g, '')
              
              FileSaver.saveAs(response.data,projectname+htd+decodeURI(filename))
          });
    },
    scjdb(){
        let commonInfoVo={}
        commonInfoVo.proname=this.$route.query.projecttitle
        commonInfoVo.htd=this.htdname
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.scjdb(commonInfoVo).then((response) => {
                  return response.data.text()       
                })
                .then(res=>{
                  
                  let code = JSON.parse(res).code
                  if(code == 200){
                    api.download(commonInfoVo.proname,commonInfoVo.htd).then((response)=>{
                    
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

    }
    
}
</script>
<style>
.row-title {
    margin-top: 20px;
    background: rgb(240, 238, 238);
    
    
  }
.content{
  cursor: pointer;
}
</style>