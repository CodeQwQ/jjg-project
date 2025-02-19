<template>
    <div class="app-container1">
        
        <div  style="margin-top: 10px;font-size: 50px; margin-left: 600px;">{{ proname }}</div>
        <el-select @change="ckxmjd($event)" v-model="value" placeholder="请选择" style="margin-bottom: 20px;margin-left: 150px;">
            <el-option
                v-for="item in options"
                :key="item.value"
                :label="item.label"
                :value="item.value"
                >
            </el-option>
        </el-select>
        <el-row v-for="item in progresses"
                :key="item.now"
                :value="item.now"
                style="margin-bottom: 20px;">
            <el-col :span="5" style="margin-right: 10px;margin-left: 150px;">
                <el-card align="middle" >{{ item.name }}</el-card>
            </el-col>
            <el-col :span="8" style="margin-right: 10px;">
                <el-progress :text-inside="true" :stroke-width="40" :percentage=" towNumber( item.jdl*100)" style="margin-top: 10px;"></el-progress>
            </el-col>
            <el-col :span="2">
                <el-card align="middle">{{ item.now+"/"+item.all }}</el-card>
            </el-col>
        </el-row>
        
        
        
    </div>
    
  </template>
  <script>
  // 引入定义接口js文件
  import api from "@/api/project/xmjd.js";
  import FileSaver from "file-saver"
  import { Loading } from 'element-ui';
  export default {
    data() {
      // 初始值
      return {
        value:0,
        proname:this.$route.query.projecttitle,
        options:[],
        progresses:[]
      }
    },
    created() {
      
      this.ckxmjd(0)
      
      
    },
    methods: {
      // 保留两位小时
      towNumber(val) {
        
        val = Math.floor(val * 100) / 100;
        return val

      },
      
      // 查看项目进度
      ckxmjd(htd){
        let commonInfoVo={};
        commonInfoVo.proname=this.proname;
        if(this.options.length==0){
          api.ckhtd(commonInfoVo).then((response)=>{
              
              
              for(let i=0;i<response.data.length;i++){
                this.options.push({value:i,label:response.data[i]})
              }
              commonInfoVo.htd=this.options[htd].label;
              this.progresses=[]
              api.ckxmjd(commonInfoVo).then((res)=>{
                    console.log("xmjd",res.data)
                    for(let i=0;i<res.data.length;i++){
                      this.progresses.push({now:res.data[i].jcs,all:res.data[i].num,jdl:res.data[i].jdl,name:res.data[i].zb})
                    }
                    
                  })
            })
        }
        else{
          commonInfoVo.htd=this.options[htd].label;
          this.progresses=[]
          api.ckxmjd(commonInfoVo).then((res)=>{
                
                for(let i=0;i<res.data.length;i++){
                  this.progresses.push({now:res.data[i].jcs,all:res.data[i].num,jdl:res.data[i].jdl,name:res.data[i].zb})
                }
                
              })
          }
        
      },
      // 查看合同段
      ckhtd(){
        let commonInfoVo={};
        commonInfoVo.proname=proname;
        this.options=[]
        api.ckhtd(commonInfoVo).then((res)=>{
              
              console.log('sss',res)
              for(let i=0;i<res.data.length;i++){
                this.options.push({value:i,label:res.data[i]})
              }
            })
      }
     
    },
  };
  </script>
<style>
 .app-container1{
    height: 100%;
    width: 100%;
    position: fixed;
    overflow: auto;
    
    background: rgb(48, 104, 136);
 }
 .logo1{
    position: relative;
    width: 520px;
    max-width: 100%;
    margin: 0 auto;
    margin-top: 50px;
    margin-bottom: 10px;
}
</style>
  