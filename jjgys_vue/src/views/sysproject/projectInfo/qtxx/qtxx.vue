<template>
    <div class="app-container">
        
    
        <div class="tool-div">
          
         
          <el-row >  
            <el-card  class="row-title"  > </el-card> 
          </el-row>
          <el-row>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="xmjd()"> 项目进度 </el-card>
            </el-col>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="janum()"> 交安数量清单 </el-card>
            </el-col>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="ljnum()"> 路基附属清单 </el-card>
            </el-col>
            </el-row>
            <el-row >  
            <el-card  class="row-title"  > </el-card> 
          </el-row>
            <el-row>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="wgjc()"> 外观检查 </el-card>
            </el-col>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="nyzljc()"> 内业资料检查 </el-card>
            </el-col>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="ryxx()"> 人员信息 </el-card>
            </el-col>
            <el-col :span="6">
              <el-card shadow="hover" class="content" align="middle" @click.native="yqxx()"> 仪器信息 </el-card>
            </el-col>
          </el-row>
          
          
        </div>
    
    </div>
    </template>
    
    <script>
    
    import { Loading } from 'element-ui';
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
            }
        },
        methods: {
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
        wgjc() {
          let projecttitle = this.$route.query.projecttitle;
          console.log("path",this.$route.matched[1].path)
          this.$router.push({path:this.$route.matched[1].path +"/qtxx/wgjc",query:{projecttitle:projecttitle}});
        },
        nyzljc() {
          let projecttitle = this.$route.query.projecttitle;
          
          this.$router.push({path:this.$route.matched[1].path +"/qtxx/nyzljc",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
        },
        yqxx() {
          let projecttitle = this.$route.query.projecttitle;
          
          this.$router.push({path:this.$route.matched[1].path +"/qtxx/yqxx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
        },
        ryxx() {
          let projecttitle = this.$route.query.projecttitle;
          
          this.$router.push({path:this.$route.matched[1].path +"/qtxx/ryxx",query:{projecttitle:projecttitle,htdname:this.htdname,fbgcName:this.fbgcName}});
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