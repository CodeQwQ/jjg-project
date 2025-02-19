<template>
<div class="app-container">
    <!--<div class="tools-div" v-show="!(username=='admin'&&type==1)">
       <el-button size="small" type="primary" @click="exportqlgc()" icon="el-icon-bottom"
        >导出桥梁工程模板文件</el-button
      > 
      <el-button size="small" icon="el-icon-top" type="primary" @click="importqlgc()">导入桥梁工程信息文件</el-button>
      <el-button size="small" type="primary" icon="el-icon-bottom" @click="scjdb()"
        >生成桥梁工程所有鉴定表</el-button
      >
    </div> -->

    <div class="tool-div">
      <el-row >  
        <el-card  class="row-title"  > 桥面 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" @click.native="qmpzd()"> 桥面平整度三米直尺法 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="qmhp()"> 桥面横坡 </el-card>
        </el-col> 
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" @click.native="qmgzsd()"> 桥面构造深度手工铺砂法 </el-card>
        </el-col>
      </el-row>

      <el-row >  
        <el-card  class="row-title"  > 上部 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="sbtqd()"> 桥梁上部砼强度 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover"  class="content" align="middle" @click.native="sbjgcc()"> 桥梁上部结构尺寸 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover"  class="content" align="middle" @click.native="sbbhchd()"> 桥梁上部保护层厚度 </el-card>
        </el-col>
      </el-row>
     
      <el-row >  
        <el-card  class="row-title"  > 下部 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="5">
          <el-card shadow="hover" class="content" align="middle" @click.native="xbdttqd()"> 桥梁下部墩台砼强度 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover" class="content" align="middle" @click.native="xbjgcc()"> 桥梁下部结构尺寸 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover" class="content" align="middle" @click.native="xbbhchd()"> 桥梁下部保护层厚度 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="xbszd()"> 桥梁下部竖直度 </el-card>
        </el-col>
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

import api from '@/api/project/fbgc/qlgc.js'
import FileSaver from "file-saver"
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
    name:"qllist",
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
      importqlgc() {
        this.dialogImportVisible = true;
        
      },
       // 上传文件触发
      fileChange(file,fileList){
        this.fileList=fileList
        
        
      },
      onUploadSuccess() {
        
        //debugger
        let fd=new FormData()
        let projectname = this.$route.query.projecttitle;
        let htd =this.htdname
        

        fd.append('file',this.fileList[0].raw)
        fd.append('proname',projectname)
        fd.append('htd',htd)
        fd.append('userid',this.userid)
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.importqlgc(fd).then((res)=>{
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
            })
        
        
        
      },
    exportqlgc() {
          let projectname = this.$route.query.projecttitle;
          let htd =this.htdname
          
          console.log('test')
          
          api.exportqlgc().then((response) => {
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
        commonInfoVo.userid=this.userid
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
    
    qmpzd() {
      let projecttitle = this.$route.query.projecttitle;
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/qmpzd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      //this.$router.push({path:"/jjgfbgc/ljgc/psdmcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    qmhp() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/qmhp",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    qmgzsd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/qmgzsd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    sbtqd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/sbtqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    sbjgcc() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/sbjgcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    sbbhchd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/sbbhchd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    xbdttqd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/xbdttqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    xbjgcc() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/xbjgcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    xbbhchd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/xbbhchd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    xbszd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/xbszd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    mcxs() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/mcxs",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    gzsd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/gzsd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    pzd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/qlgc/pzd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
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