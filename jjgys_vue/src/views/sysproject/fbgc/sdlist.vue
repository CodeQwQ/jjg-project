<template>
<div class="app-container">
    <!-- <div class="tools-div" v-show="!(username=='admin'&&type==1)">
      <el-button size="small" type="primary" @click="exportsdgc()" icon="el-icon-bottom"
        >导出隧道工程模板文件</el-button
      > 
       <el-button size="small" icon="el-icon-top" type="primary" @click="importsdgc()">导入隧道工程信息文件</el-button>
      <el-button size="small" type="primary" icon="el-icon-bottom" @click="scjdb()"
        >生成隧道工程所有鉴定表</el-button
      > 
    </div>-->

    <div class="tool-div">
      <el-row >  
        <el-card  class="row-title"  > 衬砌 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="sdcqtqd()"> 隧道衬砌砼强度 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="sdcqhd()"> 隧道衬砌厚度 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="sddmpzd()"> 隧道大面平整度 </el-card>
        </el-col>
      </el-row>
     
      <el-row >  
        <el-card  class="row-title"  > 隧道路面 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="lqlmysd()"> 隧道沥青路面压实度 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="lqlmssxs()"> 隧道沥青路面渗水系数 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="hntlmqd()"> 隧道混凝土路面强度 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="tlmxlbgc()"> 隧道砼路面相邻板高差 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="lmgzsd()">  隧道路面构造深度手工铺砂法 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="lqlmhdzx()"> 隧道沥青路面厚度钻芯法 </el-card>
        </el-col>
        <el-col :span="6">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="hntlmhdzx()"> 隧道混凝土路面厚度钻芯法 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="sdhp()"> 隧道横坡 </el-card>
        </el-col>
      </el-row>
      <el-row>
       
        
      </el-row>
      
      <el-row >  
        <el-card  class="row-title"  > 总体 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="sdztkd()"> 隧道总体宽度 </el-card>
        </el-col> 
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" style="font-size: 15px;" @click.native="jk()"> 净空 </el-card>
        </el-col> 
      </el-row>
      <el-row >  
        <el-card  class="row-title"  > 自动化 </el-card> 
      </el-row>
      <el-row>
      <el-col :span="4">
          <el-card shadow="hover"  class="content" align="middle"  @click.native="mcxs()"> 摩擦系数 </el-card>
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

import api from '@/api/project/fbgc/sdgc.js'
import FileSaver from "file-saver"
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
    name:'sdlist',
    data() {
        let projecttitle = this.$route.matched[1].meta.title;
        return {
            data: {
        projectname: projecttitle,
      },
      username:Cookies.get('username'),
      userid:Cookies.get('userid'),
      type:Cookies.get('type'),
            dialogImportVisible: false,
            htdname:this.$route.query.htdname,
            fbgcName:this.$route.query.fbgcName,
            file:'',// 待上传文件
            fileList:[],
        }
    },
    methods: {

    sdcqtqd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/cqtqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    sdcqhd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/cqhd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    sddmpzd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/dmpzd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lqlmysd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/lqlmysd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lqlmssxs() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/lqlmssxs",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    hntlmqd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/hntlmqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    tlmxlbgc() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/tlmxlbgc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lmgzsd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/lmgzsd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    lqlmhdzx() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/lqlmhdzx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    hntlmhdzx() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/hntlmhdzx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },


    sdhp() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/sdhp",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    sdztkd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/ztkd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    jk() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/jk",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    mcxs() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/mcxs",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    gzsd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/gzsd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    pzd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/pzd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    cz() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/cz",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    ldhd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path+"/sdgc/ldhd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    importsdgc() {
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
        fd.append('userid',this.userid)
        
        let load=Loading.service({fullscreen:true,text:'加载中,请稍后'})
        api.importsdgc(fd).then((res)=>{
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
    exportsdgc() {
      let projectname = this.$route.query.projecttitle;
          let htd =this.htdname
          
          console.log('test')
          
          api.exportsdgc().then((response) => {
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