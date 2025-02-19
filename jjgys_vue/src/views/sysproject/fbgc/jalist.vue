<template>
<div class="app-container">
    <!--<div class="tools-div" v-show="!(username=='admin'&&type==1)">
       <el-button size="small" type="primary" @click="exportjagc()" icon="el-icon-bottom"
        >导出交安工程模板文件</el-button
      > -->
      <!-- <el-button size="small" icon="el-icon-top" type="primary" @click="importjagc()">导入交安工程信息文件</el-button>
      <el-button size="small" type="primary" icon="el-icon-bottom" @click="scjdb()"
        >生成交安工程所有鉴定表</el-button
      > 
    </div>-->

    <div class="tool-div">
      <el-row >  
        <el-card  class="row-title"  > 标线 </el-card> 
      </el-row>
      <el-row>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="jabx()"> 交安标线 </el-card>
        </el-col>
      </el-row>
      
      <el-row >  
        <el-card  class="row-title"  > 标志</el-card> 
      </el-row>
      <el-row>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="jabz()"> 交安标志 </el-card>
        </el-col>
      </el-row>
     
      <el-row >  
        <el-card  class="row-title"  > 防护栏</el-card> 
      </el-row>
      <el-row>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="jabxfhl()"> 交安波形防护栏 </el-card>
        </el-col>
        <el-col :span="4">
          <el-card shadow="hover" class="content" align="middle" @click.native="jathlqd()"> 交安砼护栏强度 </el-card>
        </el-col>
        <el-col :span="5">
          <el-card shadow="hover" class="content" align="middle" @click.native="jathldmcc()"> 交安砼护栏断面尺寸 </el-card>
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
import api from '@/api/project/fbgc/jagc.js'
import FileSaver from "file-saver"
import { Loading } from 'element-ui';
import Cookies from 'js-cookie'
export default {
    name:'jalist',
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
    jabx() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path +"/jagc/jabx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    jabz() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path +"/jagc/jabz",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    jabxfhl() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path +"/jagc/jabxfhl",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    jathlqd() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path +"/jagc/jathlqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    jathldmcc() {
      let projecttitle = this.$route.query.projecttitle;
      
      this.$router.push({path:this.$route.matched[1].path +"/jagc/jathldmcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
    },
    importjagc() {
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
        api.importjagc(fd).then((res)=>{
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
    exportjagc() {
      let projectname = this.$route.query.projecttitle;
          let htd =this.htdname
          
          console.log('test')
          
          api.exportjagc().then((response) => {
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