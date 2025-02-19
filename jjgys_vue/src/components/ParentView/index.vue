<template>
  <div class="tree_htd">
    <div class="tree-main">
      <el-tree
      ref="tree"
      style="padding-bottom: 10px; margin-left: 5px;"
      :data="sysMenuList"
      :props="defaultProps"
     
      
      @node-click='changtree'
      current-node-key="id"
      highlight-current
    ></el-tree>
    <!-- <el-menu>
      <router-link to="htdInfo">
        <el-menu-item index="1">合同段信息</el-menu-item>
      </router-link> -->
      
    
    </div>
    <div class="tree-content">
      <router-view></router-view>
    </div>
    
    
    
  </div>
</template>

<script>
import api from '@/api/project/shice'
import {constantRoutes} from '@/router'
import router from '@/router'
import  Router  from 'vue-router'
import store from '@/store'
export default {
  
  data() {
    return {
      loading: false, // 用来标识是否正在保存请求中的标识, 防止重复提交   
      sysMenuList: [],
       defaultProps: {
          children: 'children',
          label: 'label'
        },
      flag:0,
      htdname:'',
      fbgcName:'',
    }
  },
  watch: {
    $route: function(to, from) {
      console.log(to,from,"route 更新")
      if(this.$route.path.split("/").length === 3){
        setTimeout(() => {
          this.$nextTick(()=>{
            this.changtree(this.sysMenuList[0],{parent:{data:{label:undefined}}});
          })
        }, 0);
        // this.$router.push({path:`${this.$route.path}/htdInfo`,query:{projecttitle:this.$route.matched[1].meta.title}})
      }
      this.fetchData();
    },
  },
  //当页面加载时获取数据
  created() {
    console.log('ssssss')
    if(this.$route.path.split("/").length === 3){
        setTimeout(() => {
          this.$nextTick(()=>{
            this.changtree(this.sysMenuList[0],{parent:{data:{label:undefined}}});
          })
        }, 0);
      // this.$router.push({path:`${this.$route.path}/htdInfo`,query:{projecttitle:this.$route.matched[1].meta.title}})
    }
    //页面渲染之前执行
    this.fetchData();
  },
  methods: {
    getFlag(value){
      this.flag=value
    },
    changtree(data,node){
      let projecttitle = this.$route.matched[1].meta.title;//陕西高速
      let htdname = node.parent.data.label //LJ-1
      let fbgcName = data.meta.title//路基工程......
      if(fbgcName == "合同段信息"){
       
        this.$router.push({path:this.$route.matched[1].path+"/htdInfo",query:{projecttitle:projecttitle}})
      }
      else if(fbgcName == "路桥隧信息"){
        
        this.$router.push({path:this.$route.matched[1].path+"/lqsInfo",query:{projecttitle:projecttitle}})
      }
      else if(fbgcName == "项目进度"){
        
        this.$router.push({path:this.$route.matched[1].path+"/xmjd",query:{projecttitle:projecttitle}})
      }
      else if(fbgcName == "数据可视化"){
        /*let routeData = this.$router.resolve({
            query: {projecttitle:projecttitle},
            path:'/gaosuInfo/dp'
        });
        window.open(routeData.href, '_blank');*/

        
        this.$router.push({path:this.$route.matched[1].path+"/dp",query:{projecttitle:projecttitle}})
        
      }
      else if(fbgcName == "其他信息"){
        
        this.$router.push({path:this.$route.matched[1].path+"/qtxx",query:{projecttitle:projecttitle}})
      }
      
      else if(fbgcName == "路基工程"){
        this.$router.push({path:this.$route.matched[1].path+"/fbgc/ljlist",query:{projecttitle:projecttitle,htdname:htdname,fbgcName:fbgcName}})
        
        this.flag=1
      }else if(fbgcName == "路面工程"){
        this.$router.push({path:this.$route.matched[1].path+"/fbgc/lmlist",query:{projecttitle:projecttitle,htdname:htdname,fbgcName:fbgcName}})
        //this.$router.push("/sysproject/fbgc/lmlist")
        this.flag=2
      }else if(fbgcName == "交安工程"){
        this.$router.push({path:this.$route.matched[1].path+"/fbgc/jalist",query:{projecttitle:projecttitle,htdname:htdname,fbgcName:fbgcName}})
        //this.$router.push("/sysproject/fbgc/jalist")
        this.flag=5
      }else if(fbgcName == "桥梁工程"){
        this.$router.push({path:this.$route.matched[1].path+"/fbgc/qllist",query:{projecttitle:projecttitle,htdname:htdname,fbgcName:fbgcName}})
        //this.$router.push("/sysproject/fbgc/qllist")
        this.flag=3
      }else if(fbgcName == "隧道工程"){
        this.$router.push({path:this.$route.matched[1].path+"/fbgc/sdlist",query:{projecttitle:projecttitle,htdname:htdname,fbgcName:fbgcName}})
        //this.$router.push("/sysproject/fbgc/sdlist")
        this.flag=4
      }
      
      else{
        this.flag=0
      }
    },
    
    fetchData() {
      let title = this.$route.matched[1].meta.title//项目名称 陕西高速 //
      console.log(JSON.stringify(this.$store.getters.tablemenus)+"cccc")
      this.sysMenuList =this.setdata((this.$store.getters.tablemenus.filter(item=>item.title==title)[0].data)||[])
      console.log('cccccccc',this.sysMenuList)
      
      
      /*const selfRoute =store.getters.menus
      
      console.log('add1',selfRoute)
      const route={
        
        path: 'fbgc/ljlist',
        hidden: true,
        component: () => import('@/views/sysproject/fbgc/ljlist'),
        meta: { title: '路基工程' },
        children:[
          {
            path: 'hdgqd',
            hidden: true,
            component: () => import('@/views/jjgfbgc/ljgc/hdgqd'),
            meta: { title: '涵洞砼强度' }
            },
            {
              path: 'hdjgcc',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/hdjgcc'),
              meta: { title: '涵洞结构尺寸' }
            },
            {
              path: 'ljgc/psdmcc',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/psdmcc'),
              meta: { title: '排水断面尺寸' }
            },
            {
              path: 'ljgc/pspqhd',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/pspqhd'),
              meta: { title: '排水铺砌厚度' }
            },
            {
              path: 'ljgc/xqjgcc',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/xqjgcc'),
              meta: { title: '小桥结构尺寸' }
            },
            {
              path: 'ljgc/xqgqd',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/xqgqd'),
              meta: { title: '小桥砼强度' }
            },
            {
              path: 'ljgc/zddmcc',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/zddmcc'),
              meta: { title: '支挡断面尺寸' }
            },
            {
              path: 'ljgc/zdgqd',
              hidden: true,
              component: () => import('@/views/jjgfbgc/ljgc/zdgqd'),
              meta: { title: '支挡砼强度' }
            }
        ]
      } 
      
      const routeBase=this.pushRoute( route, selfRoute)
      
      //router.matcher = new Router({ mode: "hash" }).matcher
      router.addRoutes(routeBase)
      //router.options.routes = store.getters.menus
      //global.antRouter = routeBase
      
      console.log('test',router.options.routes)
      
     
    */
      //debugger
    },
    setdata(data){
      data.forEach(item => {
        item.label = item.meta.title
        if(item.children) this.setdata(item.children)
      });
      return data
    }
  }
}
</script>


<style lang='scss' scoped>
.tree_htd{
  width: 100%;
  height: 100%;
  display: flex;
  justify-content: space-between;
  .tree-main{
    // flex: 1;//父元素采用flex布局
    width:200px;//设置最小宽度
    min-height: 850px;
    background:rgb(231, 232, 233);
    border-radius:3px;
    border:1px solid rgba(211,219,222,1);
    margin-left: 5px;
    padding: 0px;
    .el-tree{
     
        background: rgb(231, 232, 233);

    }
  }
  .tree-content{
    // flex: 1;//父元素采用flex布局
    
    width: calc(100% - 200px);
  }
  .table_box{
    width: 91%;
    height: 100%;
    overflow: hidden;
    //background-color: rgb(51, 153, 153);



}}

</style>