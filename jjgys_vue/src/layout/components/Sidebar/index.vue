<template>
  <div :class="{'has-logo':showLogo}">
    <logo v-if="showLogo" :collapse="isCollapse" />
    <el-scrollbar wrap-class="scrollbar-wrapper">
      <el-menu
        :default-active="activeMenu"
        :collapse="isCollapse"
        :background-color="variables.menuBg"
        :text-color="variables.menuText"
        :unique-opened="false"
        :active-text-color="variables.menuActiveText"
        :collapse-transition="false"
        mode="vertical"
        
        
      >
        <sidebar-item v-for="route in routes" :key="route.path" :item="route" :base-path="route.path" />
      </el-menu>
    </el-scrollbar>
  </div>
</template>

<script>
import { mapGetters } from 'vuex'
import Logo from './Logo'
import SidebarItem from './SidebarItem'
import variables from '@/styles/variables.scss'

export default {
  components: { SidebarItem, Logo },
  computed: {
    ...mapGetters([
      'sidebar'
    ]),
    routes() {
      let menus1=JSON.parse(JSON.stringify(global.antRouter))
      if(menus1.map(v=>v.meta.title).includes("交工管理")){
            
            menus1[menus1.map(v=>v.meta.title).indexOf("交工管理")].children=menus1[menus1.map(v=>v.meta.title).indexOf("交工管理")].children.filter(item=>!item.path.includes("gaosuName"))
            
          }
      
      //console.log("route1",this.$router.options.routes.concat(global.antRouter))
       return this.$router.options.routes.concat(menus1)
    },
    activeMenu() {
      const route = this.$route
      const { meta, path } = route
      // if set path, the sidebar will highlight the path you set
      console.log(route,meta.activeMenu,path,"activeMenu")
      if(path.includes("gaosuName")){
        return route.matched[1].path
      }
      if (meta.activeMenu) {
        return meta.activeMenu
      }
      return path
    },
    showLogo() {
      return this.$store.state.settings.sidebarLogo
    },
    variables() {
      return variables
    },
    isCollapse() {
      return !this.sidebar.opened
    }
  },
  watch:{
    isCollapse(){
      
    }
  }
}
</script>
