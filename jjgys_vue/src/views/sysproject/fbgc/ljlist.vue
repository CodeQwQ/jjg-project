<template>

  <div class="app-container">
  
      <!-- <div class="tools-div" v-show="!(username=='admin'&&type==1)">
        <el-button size="small" type="primary" @click="exportljgc()" icon="el-icon-bottom" 
          >导出路基工程模板文件</el-button
        > -->
        <!-- <el-button size="small" icon="el-icon-top" type="primary" @click="importljgc()" >导入路基工程信息文件</el-button>
        <el-button size="small" type="primary" icon="el-icon-bottom" @click="scjdb()" 
          >生成路基工程所有鉴定表</el-button
        > 
      </div>-->
  
      <div class="tool-div"  >
        
        
  
  
        <el-row >  
          <el-card  class="row-title"  > 涵洞 </el-card> 
        </el-row>
        <el-row>
          <el-col :span="4">
            <el-card class="content" shadow="hover" align="middle"   @click.native="hdgsd()" >涵洞砼强度</el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" class="content" align="middle" @click.native="hdjgcc()"> 涵洞结构尺寸</el-card>
          </el-col>
        </el-row>
  
        <el-row >  
          <el-card  class="row-title"  > 排水工程 </el-card> 
        </el-row>
        <el-row>
          <el-col :span="4">
            <el-card shadow="hover" align="middle" class="content" @click.native="psdmcc()" >排水断面尺寸</el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" align="middle" class="content" @click.native="pspqhd()"> 排水铺砌厚度</el-card>
          </el-col>
        </el-row>
  
        <el-row >  
          <el-card  class="row-title"  > 路基土石方 </el-card> 
        </el-row>
        <el-row>
          <el-col :span="5">
            <el-card shadow="hover" align="middle" class="content" @click.native="ljtsfysd()"> 路基土石方压实度 </el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" align="middle" class="content" @click.native="ljysdcj()"> 路基沉降量 </el-card>
          </el-col>
          <el-col :span="5">
            <el-card shadow="hover" align="middle"  class="content" @click.native="ljwcbkmlf()"> 路基弯沉贝克曼梁法 </el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" align="middle"  class="content" @click.native="ljwclcf()"> 路基弯沉落锤法 </el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" align="middle"  class="content" @click.native="ljbp()"> 路基边坡 </el-card>
          </el-col>
        </el-row>
       
        <el-row >  
          <el-card  class="row-title"  > 小桥工程 </el-card> 
        </el-row>
        <el-row>
          <el-col :span="4">
            <el-card shadow="hover" align="middle" class="content" @click.native="xqjgcc()"> 小桥结构尺寸 </el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" align="middle" class="content" @click.native="xqgqd()"> 小桥砼强度 </el-card>
          </el-col>
        </el-row>
  
        <el-row >  
          <el-card  class="row-title"  > 支挡工程 </el-card> 
        </el-row>
        <el-row>
          <el-col :span="4" >
            <el-card shadow="hover"   align="middle" class="content" @click.native="zddmcc()"> 支挡断面尺寸 </el-card>
          </el-col>
          <el-col :span="4">
            <el-card shadow="hover" align="middle" class="content" @click.native="zdgqd()"> 支挡砼强度 </el-card>
          </el-col>
          <!-- <el-col :span="4">
            <el-card shadow="hover"   align="middle" class="content" @click.native="zdbmpzd()"> 支挡表面平整度 </el-card>
          </el-col> -->
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
  import projectApi from '@/api/project/projectInfo.js'
  import api from '@/api/project/fbgc/ljgc.js'
  import FileSaver from "file-saver"
  import { Loading } from 'element-ui';
  import Cookies from 'js-cookie'
  export default {
      
      data() {
           
          return {
              data: {
              
              },
              username:Cookies.get('username'),
              userid:Cookies.get('userid'),
              type:Cookies.get('type'),
              dialogImportVisible: false,
              projecttitle : this.$route.query.projecttitle,
              htdname:this.$route.query.htdname,
              fbgcName:this.$route.query.fbgcName,
              file:'',// 待上传文件
              fileList:[],
          }
      },
      methods: {
        
        hdjgcc(){
          let projecttitle = this.$route.query.projecttitle;
          
          this.$router.push({path:this.$route.matched[1].path+"/ljgc/hdjgcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
  
        },
      hdgsd() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/hdgqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      psdmcc() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/psdmcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      pspqhd() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/pspqhd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      ljtsfysd() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/ljtsfysd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      ljysdcj() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/ljysdcj",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      ljwcbkmlf() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/ljwcbkmlf",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      ljwclcf() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/ljwclcf",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      ljbp() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/ljbp",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
  
  
      xqjgcc() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/xqjgcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      xqgqd() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/xqgqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      zddmcc() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/zddmcc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      zdgqd() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/zdgqd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      zdbmpzd() {
        let projecttitle = this.$route.query.projecttitle;
        
        this.$router.push({path:this.$route.matched[1].path+"/ljgc/zdbmpzd",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
      },
      importljgc() {
        this.dialogImportVisible = true;
        console.log("qqqqqq")
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
        api.importljgc(fd).then((res)=>{
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
      exportljgc() {
          let projectname = this.$route.query.projecttitle;
          let htd =this.htdname
          
          console.log('test')
          
          api.exportljgc().then((response) => {
              var temp = response.headers["content-disposition"]
              var filename=temp.split('=')[1]
              filename=filename.replace(/['"]/g, '')
              console.log('鉴定表',filename)
              console.log('鉴定表',decodeURI(filename))
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