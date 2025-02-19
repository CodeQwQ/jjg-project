import Vue from 'vue'
import Router from 'vue-router'

Vue.use(Router)

/* Layout */
import Layout from '@/layout'

/**
 * Note: sub-menu only appear when route children.length >= 1
 * Detail see: https://panjiachen.github.io/vue-element-admin-site/guide/essentials/router-and-nav.html
 *
 * hidden: true                   if set true, item will not show in the sidebar(default is false)
 * alwaysShow: true               if set true, will always show the root menu
 *                                if not set alwaysShow, when item has more than one children route,
 *                                it will becomes nested mode, otherwise not show the root menu
 * redirect: noRedirect           if set noRedirect will no redirect in the breadcrumb
 * name:'router-name'             the name is used by <keep-alive> (must set!!!)
 * meta : {
    roles: ['admin','editor']    control the page roles (you can set multiple roles)
    title: 'title'               the name show in sidebar and breadcrumb (recommend set)
    icon: 'svg-name'/'el-icon-x' the icon show in the sidebar
    breadcrumb: false            if set false, the item will hidden in breadcrumb(default is true)
    activeMenu: '/example/list'  if set path, the sidebar will highlight the path you set
  }
 */

/**
 * constantRoutes
 * a base page that does not have permission requirements
 * all roles can be accessed
 */
export const constantRoutes = [



  {
    path: '/login',
    component: () => import('@/views/login/index'),
    hidden: true
  },
  {
    path: '/register',
    name: 'register',
    component: () => import('@/views/register/index')
  },
  {
    path: '/404',
    name: 'ERROR404',
    component: () => import("@/components/404error/404"),
  },

  // {
  //   path: '/404',
  //   component: () => import('@/views/404'),
  //   hidden: true
  // },

  {
    path: '/',
    component: Layout,
    redirect: '/dashboard', // 重定向到dashborad
    children: [{
      path: 'dashboard',
      name: 'Dashboard',
      component: () => import('@/views/dashboard/index'),
      meta: { title: '首页', icon: 'dashboard' }
    }]
  },
  {
    path: '/sysproject',
    component: Layout,
    children: [{
      path: 'projectInfo/lqs/ql',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/ql'),
      meta: { title: '桥梁清单' }
    },
    {
      path: 'projectInfo/lqs/sd',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/sd'),
      meta: { title: '隧道清单' }
    },
    {
      path: 'projectInfo/lqs/fhlm',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/fhlm'),
      meta: { title: '复合路面清单' }
    },
    {
      path: 'projectInfo/lqs/hntlmjzd',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/hntlmjzd'),
      meta: { title: '混凝土路面及匝道清单' }
    },
    {
      path: 'projectInfo/lqs/sfz',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/sfz'),
      meta: { title: '收费站清单' }
    },
    {
      path: 'projectInfo/lqs/ljx',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/ljx'),
      meta: { title: '连接线清单' }
    },
    {
      path: 'projectInfo/lqs/ljxql',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/ljxql'),
      meta: { title: '连接线' }
    },
    {
      path: 'projectInfo/lqs/ljxsd',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/ljxsd'),
      meta: { title: '连接线隧道清单' }
    },
    {
      path: 'projectInfo/lqs/ljxhntlm',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/ljxhntlm'),
      meta: { title: '连接线混凝土路面清单' }
    },
    {
      path: 'projectInfo/lqs/janum',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/janum.vue'),
      meta: { title: '交安数量' }
    },
    {
      path: 'projectInfo/lqs/ljnum',
      hidden: true,
      component: () => import('@/views/sysproject/projectInfo/lqs/ljnum'),
      meta: { title: '路基数量' }
    },
    //分部工程
    {
      path: 'fbgc/ljlist',
      hidden: true,
      component: () => import('@/views/sysproject/fbgc/ljlist'),
      meta: { title: '路基工程' }
    },
    {
      path: 'fbgc/lmlist',
      hidden: true,
      component: () => import('@/views/sysproject/fbgc/lmlist'),
      meta: { title: '路面工程' }
    },
    {
      path: 'fbgc/qllist',
      hidden: true,
      component: () => import('@/views/sysproject/fbgc/qllist'),
      meta: { title: '桥梁工程' }
    },
    {
      path: 'fbgc/sdlist',
      hidden: true,
      component: () => import('@/views/sysproject/fbgc/sdlist'),
      meta: { title: '隧道工程' }
    },
    {
      path: 'fbgc/jalist',
      hidden: true,
      component: () => import('@/views/sysproject/fbgc/jalist'),
      meta: { title: '交安工程' }
    }
    ]
  },
  {
    path: '/gaosuInfo',
    hidden: true,
    component: Layout,
    redirect: 'htdInfo',
    children: [
      {
        path: 'htdInfo',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/htdInfo.vue'),
        meta: { title: '合同段信息' }
      },
      {
        path: 'lqsInfo',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/lqsInfo.vue'),
        meta: { title: '路桥隧信息' }
      },
      {
        path: 'xmjd',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/xmjd.vue'),
        meta: { title: '项目进度' }
      },
      {
        path: 'xmjdmanage',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/xmjdmanage.vue'),
        meta: { title: '项目进度管理' }
      },
      {
        path: 'dp',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/dp/dp.vue'),
        meta: { title: '大屏' }
      },

      {
        path: 'qtxx',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/qtxx/qtxx.vue'),
        meta: { title: '其他信息' },

      },
      {
        path: 'qtxx/wgjc',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/qtxx/wgjc.vue'),
        meta: { title: '外观检查' }
      },
      {
        path: 'qtxx/nyzljc',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/qtxx/nyzljc.vue'),
        meta: { title: '内页资料检查' }
      },
      {
        path: 'qtxx/ryxx',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/qtxx/ryxx.vue'),
        meta: { title: '人员信息' }
      },
      {
        path: 'qtxx/yqxx',
        hidden: true,
        component: () => import('@/views/sysproject/projectInfo/qtxx/yqxx.vue'),
        meta: { title: '仪器信息' }
      },

    ]
  },
  {
    path: '/jgproject',
    component: Layout,
    children: [{
      path: 'guifan/new',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/new.vue'),
      meta: { title: '新规范' }
    },
    {
      path: 'guifan/old',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/old.vue'),
      meta: { title: '旧规范' }
    },
    {
      path: 'guifan/htdinfo',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/htdInfo.vue'),
      meta: { title: '合同段信息' }
    },
    {
      path: 'guifan/dwgctze',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/dwgctze.vue'),
      meta: { title: '单位工程投资额' }
    },
    {
      path: 'guifan/qlqd',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/qlqd.vue'),
      meta: { title: '桥梁清单' }
    },
    {
      path: 'guifan/sdqd',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/sdqd.vue'),
      meta: { title: '隧道清单' }
    },
    {
      path: 'guifan/ljxqd',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/ljxqd.vue'),
      meta: { title: '连接线清单' }
    },
    {
      path: 'guifan/hntlmjzd',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/hntlmjzd.vue'),
      meta: { title: '混凝土路面及匝道' }
    },
    {
      path: 'guifan/sfzqd',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/sfzqd.vue'),
      meta: { title: '收费站清单' }
    },
    {
      path: 'guifan/wgkf',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/wgkf.vue'),
      meta: { title: '外观扣分' }
    },
    {
      path: 'guifan/nyzlkf',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/nyzlkf.vue'),
      meta: { title: '内页资料扣分' }
    },
    {
      path: 'guifan/ryxxjgfc',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/lqszb/ryxxjgfc.vue'),
      meta: { title: '人员信息' }
    },
    {
      path: 'guifan/cz',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/cz.vue'),
      meta: { title: '车辙' }
    },
    {
      path: 'guifan/hntlmxlbgc',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/hntlmxlbgc.vue'),
      meta: { title: '混凝土路面相邻板高差' }
    },
    {
      path: 'guifan/mcxs',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/mcxs.vue'),
      meta: { title: '摩擦系数' }
    },
    {
      path: 'guifan/gzsdsgpsf',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/gzsdsgpsf.vue'),
      meta: { title: '构造深度手工铺砂法' }
    },
    {
      path: 'guifan/gzsdjqjcf',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/gzsdjqjcf.vue'),
      meta: { title: '构造深度机器检测法' }
    },
    {
      path: 'guifan/pzd',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/pzd.vue'),
      meta: { title: '平整度' }
    },
    {
      path: 'guifan/jglmwcbkml',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/jglmwcbkml.vue'),
      meta: { title: '路面弯沉(贝克曼梁法)' }
    },
    {
      path: 'guifan/jglmwclcf',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/jglmwclcf.vue'),
      meta: { title: '路面弯沉(落锤法)' }
    },

    {
      path: 'guifan/jabx',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/jabx.vue'),
      meta: { title: '交安标线' }
    },
    {
      path: 'guifan/jafhl',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/fczb/jafhl.vue'),
      meta: { title: '交安防护栏' }
    },


    {
      path: 'guifan/jgjcsj',
      hidden: true,
      component: () => import('@/views/jgproject/guifan/jgjcsj.vue'),
      meta: { title: '交工检测数据' }
    },
    ]
  },
  {
    path: '/jjgfbgc',
    component: Layout,
    children: [


      {
        path: 'ljgc/hdgqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/hdgqd'),
        meta: { title: '涵洞砼强度' }
      },
      {
        path: 'ljgc/hdjgcc',
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
        path: 'ljgc/ljtsfysd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/ljtsfysd'),
        meta: { title: '路基土石方压实度' }
      },
      {
        path: 'ljgc/ljysdcj',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/ljysdcj'),
        meta: { title: '路基压实度沉降' }
      },
      {
        path: 'ljgc/ljwcbkmlf',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/ljwcbkmlf'),
        meta: { title: '路基弯沉贝克曼梁法' }
      },
      {
        path: 'ljgc/ljwclcf',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/ljwclcf'),
        meta: { title: '路基弯沉落锤法' }
      },
      {
        path: 'ljgc/ljbp',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/ljbp'),
        meta: { title: '路基边坡' }
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
      },
      {
        path: 'ljgc/zdbmpzd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/ljgc/zdbmpzd'),
        meta: { title: '支挡表面平整度' }
      },

      {
        path: 'lmgc/lqlmysd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/lqlmysd'),
        meta: { title: '沥青路面压实度' }
      },
      {
        path: 'lmgc/lmwcbkmlf',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/lmwcbkmlf'),
        meta: { title: '路面弯沉贝克曼梁法' }
      },
      {
        path: 'lmgc/lmwclcf',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/lmwclcf'),
        meta: { title: '路面弯沉落锤法' }
      },
      {
        path: 'lmgc/lqlmssxs',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/lqlmssxs'),
        meta: { title: '沥青路面渗水系数' }
      },
      {
        path: 'lmgc/hntlmqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/hntlmqd'),
        meta: { title: '混凝土路面强度' }
      },
      {
        path: 'lmgc/tlmxlbgc',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/tlmxlbgc'),
        meta: { title: '砼路面相邻板高差' }
      },
      {
        path: 'lmgc/lmgzsd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/lmgzsd'),
        meta: { title: '路面构造深度手工铺砂法' }
      },
      {
        path: 'lmgc/gslqlmhdzx',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/gslqlmhdzx'),
        meta: { title: '高速沥青路面厚度钻芯法' }
      },
      {
        path: 'lmgc/hntlmhdzx',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/hntlmhdzx'),
        meta: { title: '混凝土路面厚度钻芯法' }
      },
      {
        path: 'lmgc/lmhp',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/lmhp'),
        meta: { title: '路面横坡' }
      },
      {
        path: 'lmgc/mcxs',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/mcxs'),
        meta: { title: '摩擦系数' }
      },
      {
        path: 'lmgc/gzsd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/gzsd'),
        meta: { title: '构造深度' }
      },
      {
        path: 'lmgc/pzd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/pzd'),
        meta: { title: '平整度' }
      },
      {
        path: 'lmgc/cz',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/cz'),
        meta: { title: '车辙' }
      },
      {
        path: 'lmgc/ldhd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/lmgc/ldhd'),
        meta: { title: '雷达厚度' }
      },


      {
        path: 'qlgc/qmpzd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/qmpzd'),
        meta: { title: '桥面平整度三米直尺法' }
      },
      {
        path: 'qlgc/qmhp',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/qmhp'),
        meta: { title: '桥面横坡' }
      },
      {
        path: 'qlgc/qmgzsd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/qmgzsd'),
        meta: { title: '桥面构造深度手工铺砂法' }
      },
      {
        path: 'qlgc/sbtqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/sbtqd'),
        meta: { title: '桥梁上部砼强度' }
      },
      {
        path: 'qlgc/sbjgcc',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/sbjgcc'),
        meta: { title: '桥梁上部结构尺寸' }
      },
      {
        path: 'qlgc/sbbhchd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/sbbhchd'),
        meta: { title: '桥梁上部保护层厚度' }
      },
      {
        path: 'qlgc/xbdttqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/xbdttqd'),
        meta: { title: '桥梁下部墩台砼强度' }
      },
      {
        path: 'qlgc/xbjgcc',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/xbjgcc'),
        meta: { title: '桥梁下部结构尺寸' }
      },
      {
        path: 'qlgc/xbbhchd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/xbbhchd'),
        meta: { title: '桥梁下部保护层厚度' }
      },
      {
        path: 'qlgc/xbszd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/xbszd'),
        meta: { title: '桥梁下部竖直度' }
      },
      {
        path: 'qlgc/pzd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/pzd'),
        meta: { title: '平整度' }
      },
      {
        path: 'qlgc/mcxs',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/mcxs'),
        meta: { title: '摩擦系数' }
      },
      {
        path: 'qlgc/gzsd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/qlgc/gzsd'),
        meta: { title: '构造深度' }
      },

      {
        path: 'sdgc/cqtqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/cqtqd.vue'),
        meta: { title: '隧道衬砌砼强度' }
      },
      {
        path: 'sdgc/cqhd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/cqhd.vue'),
        meta: { title: '隧道衬砌厚度' }
      },
      {
        path: 'sdgc/dmpzd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/dmpzd.vue'),
        meta: { title: '隧道大面平整度' }
      },
      {
        path: 'sdgc/lqlmysd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/lqlmysd.vue'),
        meta: { title: '隧道沥青路面压实度' }
      },
      {
        path: 'sdgc/lqlmssxs',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/lqlmssxs.vue'),
        meta: { title: '隧道沥青路面渗水系数' }
      },
      {
        path: 'sdgc/hntlmqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/hntlmqd.vue'),
        meta: { title: '隧道混凝土路面强度' }
      },
      {
        path: 'sdgc/tlmxlbgc',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/tlmxlbgc.vue'),
        meta: { title: '隧道砼路面相邻板高差' }
      },
      {
        path: 'sdgc/lmgzsd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/lmgzsd.vue'),
        meta: { title: '隧道路面构造深度手工铺砂法' }
      },
      {
        path: 'sdgc/lqlmhdzx',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/lqlmhdzx.vue'),
        meta: { title: '隧道沥青路面厚度钻芯法' }
      },
      {
        path: 'sdgc/hntlmhdzx',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/hntlmhdzx.vue'),
        meta: { title: '隧道混凝土路面厚度钻芯法' }
      },
      {
        path: 'sdgc/sdhp',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/sdhp.vue'),
        meta: { title: '隧道横坡' }
      },
      {
        path: 'sdgc/ztkd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/sdztkd.vue'),
        meta: { title: '隧道总体宽度' }
      },
      {
        path: 'sdgc/jk',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/jk.vue'),
        meta: { title: '净空' }
      },
      {
        path: 'sdgc/pzd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/pzd'),
        meta: { title: '平整度' }
      },
      {
        path: 'sdgc/mcxs',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/mcxs'),
        meta: { title: '摩擦系数' }
      },
      {
        path: 'sdgc/gzsd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/gzsd'),
        meta: { title: '构造深度' }
      },
      {
        path: 'sdgc/cz',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/cz'),
        meta: { title: '车辙' }
      },
      {
        path: 'sdgc/ldhd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/sdgc/ldhd'),
        meta: { title: '雷达厚度' }
      },



      {
        path: 'jagc/jabx',
        hidden: true,
        component: () => import('@/views/jjgfbgc/jagc/jabx.vue'),
        meta: { title: '交安标线' }
      },
      {
        path: 'jagc/jabz',
        hidden: true,
        component: () => import('@/views/jjgfbgc/jagc/jabz.vue'),
        meta: { title: '交安标志' }
      },
      {
        path: 'jagc/jabxfhl',
        hidden: true,
        component: () => import('@/views/jjgfbgc/jagc/jabxfhl.vue'),
        meta: { title: '交安波形防护栏' }
      },
      {
        path: 'jagc/jathlqd',
        hidden: true,
        component: () => import('@/views/jjgfbgc/jagc/jathlqd.vue'),
        meta: { title: '交安砼护栏强度' }
      },
      {
        path: 'jagc/jathldmcc',
        hidden: true,
        component: () => import('@/views/jjgfbgc/jagc/jathldmcc.vue'),
        meta: { title: '交安砼护栏断面尺寸' }
      },


    ],
    /*children: [
      {
      path: 'ljgc/hdgqd',
      hidden: true,
      component: () => import('@/views/jjgfbgc/ljgc/hdgqd'),
      meta: { title: '涵洞砼强度' }
      },
      {
        path: 'ljgc/hdjgcc',
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
    ]*/
  },
  /* {     
        path: '*',
        redirect: '/404', 
        hidden: true 
  } */

  // {
  //   path: '/system',
  //   component: Layout,
  //   redirect: '/system/sysUser',
  //   name: 'System',
  //   meta: { title: '系统管理', icon: 'el-icon-s-help' },
  //   children: [
  //     {
  //       path: 'sysRole',
  //       name: 'SysRole',
  //       component: () => import('@/views/system/sysRole/list'),
  //       meta: { title: '角色管理', icon: 'table' }
  //     },
  //     {
  //       path: 'sysUser',
  //       name: 'SysUser',
  //       component: () => import('@/views/system/sysUser/list'),
  //       meta: { title: '用户管理', icon: 'tree' }
  //     },
  //     {
  //       name: 'sysMenu',
  //       path: 'sysMenu',
  //       component: () => import('@/views/system/sysMenu/list'),
  //       meta: {
  //         title: '菜单管理',
  //         icon: 'el-icon-s-unfold'
  //       },
  //     },
  //     {
  //       path: 'assignAuth',
  //       component: () => import('@/views/system/sysRole/assignAuth'),
  //       meta: {
  //         activeMenu: '/system/sysRole',
  //         title: '角色授权'
  //       },
  //       hidden: true,
  //     }
  //   ]
  // },

  // // 404 page must be placed at the end !!!
  // { path: '*', redirect: '/404', hidden: true }
]

const createRouter = () => new Router({
  // mode: 'history', // require service support
  scrollBehavior: () => ({ y: 0 }),
  routes: constantRoutes
})

const router = createRouter()

// Detail see: https://github.com/vuejs/vue-router/issues/1234#issuecomment-357941465
export function resetRouter() {
  const newRouter = createRouter()
  router.matcher = newRouter.matcher // reset router
}

export default router
