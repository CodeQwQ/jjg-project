<template>
        <v-scale-screen width="1920" height="1080" :autoScale={x:false,y:true} style="overflow: hidden;" :delay=10 >
        <div id="index" ref="appRef" :key="key">
          

          <div class="bg">
            
            

              <el-button @click="fullScreen()" type="primary" style="width: 60px;"> {{ full }}</el-button>
              <div  class="host-body">
                <div class="d-flex jc-center">
                  <dv-decoration-10 class="dv-dec-10" />
                  <div class="d-flex jc-center">
                    <dv-decoration-8 class="dv-dec-8"  />
                    
                    
                    <div class="title">
                      <div class='logo' style="width:60px;height:60px;;border-radius:50%;overflow:hidden;margin-left: 30px;">
                        <img src="@/assets/login/logo.png" style="clear: both;display: block;margin: auto; width:60px;height: 60px">
                      </div>
                      <span class="title-text"  style="font-size: 25px;width: 300px; margin-bottom: 30px;margin-left: 20px;" > 陕西交控工程技术有限公司</span>
                      <span  style="font-size: 20px" > {{proname}}项目交工验收质量检测可视化数据展示平台</span>
                      <dv-decoration-6
                        class="dv-dec-6"
                        :reverse="true"
                        :color="['#50e3c2', '#67a1e5']"
                      />
                    </div>
                    <dv-decoration-8
                      class="dv-dec-8"
                      :reverse="true"
                      
                    />
                  </div>
                  <dv-decoration-10 class="dv-dec-10-s" />
                </div>
        
                <!-- 第二行 -->
                <!-- <div class="d-flex jc-between px-2">
                  <div class="d-flex aside-width">
                    <div class="react-left ml-4 react-l-s">
                      <span class="react-before"></span>
                      <span class="text"> 合同段信息 </span>
                    </div>
                    <div class="react-left ml-3">
                      <span class="text">项目总览</span>
                    </div>
                  </div>
                  <div class="d-flex aside-width">
                    <div class="react-right bg-color-blue mr-3">
                      <span class="text fw-b"> 工程可视化数据 </span>
                    </div>
                    <div class="react-right mr-4 react-l-s">
                      <span class="react-after"></span>
                      <span class="text">
                        数据分析
                      </span>
                    </div>
                  </div>
                </div> -->
        
                <div class="body-box">
                  
                  <div class="content-boxleft">
                    <div >
                      <dv-border-box-12 ref="borderBox1">
                        <center-left1 />
                      </dv-border-box-12>
                      
                    
                    </div>
                    <div>
                      <dv-border-box-12 ref="borderBox2">
                        <center-left2 />
                      </dv-border-box-12>
                    </div>

                  </div>
                  <div class="content-box">
                    
                    
                    
                    
                    <div>
                      <dv-border-box-2 ref="borderBox3">
                        <center-right-1/>
                      </dv-border-box-2>
                      
                    </div>
                    <div>
                      <dv-border-box-7 ref="borderBox4">
                        <center-right2 />
                      </dv-border-box-7>
                      
                    </div>
                    <div>
                      <dv-border-box-12 ref="borderBox5">
                        <center/>
                      </dv-border-box-12>
                      
                    </div>
                    
                  </div>
                  
                  
                  
                </div>
              </div>
            
            
          </div>
        
        </div>
      </v-scale-screen>
      
    
    
  
  
  </template>
<script >
import CenterLeft1 from './centerleft1.vue'
import CenterLeft2 from './centerleft2.vue'
import Center from './center.vue'
import CenterRight1 from './centerright1.vue'
import CenterRight2 from './centerright2.vue'
import BottomLeft from './bottomleft.vue'
import BottomRight from './bottomright.vue'
import screenfull from 'screenfull'
import ScaleBox from '@/components/ScaleBox.vue'




export default{
    components: {
        Center,
        CenterLeft1,
        CenterLeft2,
        CenterRight1,
        CenterRight2,
        BottomLeft,
        BottomRight,
        ScaleBox,
        ScaleBox,
       
        
    },
    data() {
      // 初始值
      return {
        full:"按我全屏",
        proname:this.$route.query.projecttitle,
        key:1,
        open:this.$store.state.app.sidebar.opened,
        width:"",
        newWidth:""
      }
    },
    created() {
      this.$store.state.app.sidebar.opened=false
      
      
      
      
    },
    mounted(){
      let that = this;
      
      setTimeout(() => {
        this.$refs.borderBox1.initWH()
        }, 100)
        setTimeout(() => {
        this.$refs.borderBox2.initWH()
        }, 100)
        setTimeout(() => {
        this.$refs.borderBox3.initWH()
        }, 100)
        setTimeout(() => {
        this.$refs.borderBox4.initWH()
        }, 100)
        setTimeout(() => {
        this.$refs.borderBox5.initWH()
        }, 100)
      //let MutationObserver = window.MutationObserver || window.WebKitMutationObserver || window.MozMutationObserver;
      var element=document.getElementById("index");
      this.width=element.clientWidth
      var element1=document.getElementsByClassName("app-main")[0]
      console.log("sssss",element1)

     
      window.onresize = function(){
				if(!that.checkFull()){
					// 退出全屏后要执行的动作
					
					that.full = "按我全屏";
				}
        
			}
      
    },
    watch:{
      "$store.state.app.sidebar.opened":{
        handler:function(newVal,oldVal){
          console.log(newVal,oldVal);
          this.resize(newVal)
        }
      }
    },
    methods:{
      resize(val){
        if(val){
          console.log("缩放",this.width)
          var element=document.getElementById("index");
          //console.log("缩放1",element.clientWidth)
          this.newWidth=element.clientWidth
          var ww=this.width/this.newWidth
          
          console.log("缩放1",this.newWidth)
          element.style.transform = "scale(" + ww + "," + 1 + ")"; 
        }
        else{
          
          var element=document.getElementById("index");
          
          element.style.transform = "scale(" + 1 + "," + 1 + ")"; 
        }
      },
      checkFull(){
        //判断浏览器是否处于全屏状态 （需要考虑兼容问题）
        //火狐浏览器
        var isFull = document.mozFullScreen||
        document.fullScreen ||
        //谷歌浏览器及Webkit内核浏览器
        document.webkitIsFullScreen ||
        document.webkitRequestFullScreen ||
        document.mozRequestFullScreen ||
        document.msFullscreenEnabled
        if(isFull === undefined) {
            isFull = false
        }
        return isFull;
      },


      fullScreen() {
        

        const element = document.getElementById('index');
        screenfull.toggle(element);
        if(this.full=="退出全屏"){
          this.full="按我全屏"
        }
        else{
          this.full="退出全屏"
        }
        
        
      },
    }
}
</script>
<style lang="scss" scoped>
@import '@/assets/index.scss' ;
@import '@/assets/_variables.scss';
@import '@/assets/style.scss';
</style>