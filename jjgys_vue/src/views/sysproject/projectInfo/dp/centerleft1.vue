<template>
  <div id="centerLeft1" v-loading="loading" element-loading-text="拼命加载中"
                            element-loading-spinner="el-icon-loading"
                            element-loading-background="rgba(0, 0, 0, 0.8)">
    <div>
        <span style="background: rgb(75, 123, 33);font-size: 25px;margin-left: 15px;border-radius: 5%;">
            合同段信息
        </span>
        
    </div>
    <div id="myChart"  style="margin-left: 15px; margin-top: 5px;">
        <el-table  
        :data="list"
       
        
        :header-cell-style="{ 'text-align': 'center' ,'color':'white'}"
        :cell-style="{ 'text-align': 'center' ,'color':'white'}"
        
        :max-height="tableHeight"
        :height="tableHeight"
        ref="wt_table"
        @mouseenter.native="mouseEnter"
        @mouseleave.native="mouseLeave"
        :fit="true"
        >
        
        
        <el-table-column prop="name" label="合同段名称"  />
        
        <el-table-column prop="sgdw" label="施工单位"  />
        <el-table-column prop="jldw" label="监理单位"  />
        <el-table-column prop="zhq" label="桩号起"  />
        <el-table-column prop="zhz" label="桩号止"  />

        
        </el-table>
    </div>
  </div>  
    
    
  </template>
<script>
import * as echarts from 'echarts'
import api from '@/api/project/dp.js'
let rolltimer = ''
let changetimer = ''
    export default{
        data() {
            // 初始值
            return {
                list: [], // 项目列表
                height1:"355px",
                width1:"650px",
                tableHeight:"355px",
                loading:true,
                rollPx:1,
                rollTime: 5,
                refreshTime:5,
                proname:this.$route.query.projecttitle
                
                
            }
        },
        mounted(){
            this.$nextTick(() => {
            // 初始化echarts实例
                this.init()
            })
            this.autoRoll()
            //this.autoChange()
           

        },
        destroyed(){
            this.autoRoll(true)
        },
        methods:{
            init() {
                api.gethtdinfo(this.proname).then((response) => {
                    console.log("合同段信息",response)
                    if(response.code==200){
                        this.list=response.data;
                        this.loading=false
                        
                    }
                    
                })
                

                //this.getEcharts()
            },
            
            getEcharts() {
                var element=document.getElementsByClassName("dv-border-box-12")[0];

                
                //this.width1=element.clientWidth;
                //this.height1=element.clientHeight;
                
                console.log("cews",this.width1)
                setTimeout(_ => {
                    this.list.forEach(e => {
                    let myChart = echarts.init(this.$refs['echarts']);
                    myChart.setOption({
                        grid: {
                        left: "0",
                        top: "0",
                        right: "0",
                        bottom: "0",
                        containLabel: true,
                        },
                        xAxis: {
                        type: 'category',
                        //不显示x轴线
                        show: false,

                        },
                        yAxis: {
                        type: 'value',
                        show: false,
                        },
                        series: [{
                        data: e.num,
                        //单独修改当前线条的颜色
                        lineStyle: {
                            normal: {
                            color: "#f00",
                            width: 1,
                            },
                        },
                        type: 'line',
                        smooth: true,
                        symbol: 'none',
                        }]

                    });
                    window.addEventListener("resize", () => {
                        myChart.resize();
                    });
                    })
                }, 1000)
            },
            // 设置自动滚动
            autoRoll(stop) {
            if (stop) {
                
                clearInterval(rolltimer)
                return
            }

            // 拿到表格挂载后的真实DOM
            const table = this.$refs['wt_table']
            // 拿到表格中承载数据的div元素
            
                const divData = table.bodyWrapper
            
            
            
            // 拿到元素后，对元素进行定时增加距离顶部距离，实现滚动效果
            rolltimer = setInterval(() => {
                // 元素自增距离顶部像素
                divData.scrollTop += this.rollPx
               
                // 判断元素是否滚动到底部(可视高度+距离顶部=整个高度)
                if ( parseInt(divData.clientHeight + divData.scrollTop) == divData.scrollHeight) {
                // 重置table距离顶部距离
                divData.scrollTop = 0
                }
            }, this.rollTime * 10)
            },
            
            // 鼠标进入
            mouseEnter(time) {
            // 鼠标进入停止滚动和切换的定时任务
            this.autoRoll(true)
            
            },
            // 鼠标离开
            mouseLeave() {
            // 开启
            this.autoRoll()
            
            },





        }
     
    }
    
  
</script>
<style  scoped>
#centerLeft1{
    padding: 16px;
    height: 100%;
    width: 97%;
    
    
} 
#myChart{
    background-color: transparent;
}

/deep/ .el-table, /deep/ .el-table__expanded-cell{
  background-color: transparent;
}
/* 表格内背景颜色 */
/deep/ .el-table th,
/deep/ .el-table tr,
/deep/ .el-table td {
  background-color: transparent;
}
/deep/ .el-table tbody tr { 
    pointer-events:none; 
}



</style>  
