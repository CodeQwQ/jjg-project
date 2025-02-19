<template>
    <div id="centerleft2" v-loading="loading" element-loading-text="拼命加载中"
                            element-loading-spinner="el-icon-loading"
                            element-loading-background="rgba(0, 0, 0, 0.8)">
        <el-tabs  :key="key" v-model="activeName" class="mytabs" @tab-click="handleClick" style="margin-left: 15px;margin-top: 5px;color: white; height: 50px; width: 650px;caret-color: transparent;">
            
            <el-tab-pane 
                v-for="(item,index) in list"
                :key="index"
            ><span slot="label" @mouseover="mouseover" @mouseout="mouseout" style="color:rgb(104, 186, 32);" >{{ item }}</span></el-tab-pane>
        </el-tabs>
        

        <div class="charts1">
            <div id="mychart1" style="width: 300px;height:190px; margin-left: 15px; ">

            </div>
            <div id="mychart2" style="width: 300px;height:190px; margin-left: 15px;">

            </div>
        </div>
        <div class="charts2">
            <div id="mychart3" style="width: 300px;height:190px; margin-left: 15px;">

            </div>
            <div id="mychart4" style="width: 300px;height:190px; margin-left: 15px;">

            </div>
        </div>
        
    </div>
    
    
  </template>
<script>
import * as echarts from 'echarts'
import api from '@/api/project/dp.js'
import { withConverter } from 'js-cookie'
let rolltimer=''
    export default{
        data(){
            // 初始值
            return {
                list:[
                    
                    
                ],
                qsnumdata:[],
                activeName:"0",
                rollPx:1,
                rollTime: 150,
                refreshTime:5,
                loading:true,
                key:0,
                proname:this.$route.query.projecttitle,
                option :{
                    animation: false,
                    title:{
                        text: '主标题',
                        textStyle: {
                            color: 'red'
                        },
                        x:"center",
                        y:"top"
                    },
                    tooltip : {
                        trigger: 'item',
                        show:true,
                        
                    },
                    xAxis: {
                        data: ["1"],
                        axisLabel: {
                            //x轴文字的配置
                            show: true,
                            interval: 0,//使x轴文字显示全
                            textStyle:{
                                color:'#ffffff'
                            }
                        }

                    },
                    yAxis: { minInterval: 1,
                        axisLabel: {
                            
                            show: true,
                            
                            textStyle:{
                                color:'#ffffff'
                            }
                        }
                    },
                    series: [
                        {
                        type: 'bar',
                        data: [1],
                        itemStyle: {
                                    normal: {
                        　　　　　　　　//这里是重点
                                        color: function(params) {
                                            //注意，如果颜色太少的话，后面颜色不会自动循环，最好多定义几个颜色
                                            var colorList = ['#2f4554',   '#749f83', '#ca8622','#c23531','#61a0a8'];
                                            //给大于颜色数量的柱体添加循环颜色的判断
                                            var index;
                                            if (params.dataIndex >= colorList.length) {
                                                index = params.dataIndex - colorList.length;
                                                return colorList[index];
                                            }

                                            return colorList[params.dataIndex]
                                        }
                                    }
                                }
                        }
                    ]
                }
                
                
                
            }
        },
        mounted(){
            this.$nextTick(() => {
            
                this.init()
                
            })
            
           

        },
        destroyed(){
            this.autoRoll(true)
        },
        
        methods:{
            init() {
                api.getqsnum(this.proname).then((response) => {
                    
                    if(response.code==200){
                        this.list=Object.keys(response.data).sort()
                        for(let i=0;i<this.list.length;i++){
                            if(this.list[i]=="建设项目"){
                                let temp = this.list[i];
                                this.list[i] = this.list[0];
                                this.list[0] = temp;
                                break;
                            }
                        }
                        
                        
                        this.loading=false
                        this.qsnumdata=response.data
                        console.log("桥梁隧道",response.data)
                        this.autoRoll()
                        this.drawEcharts()


                    }
                    

                })
                

                //this.getEcharts()
            },
            // 点击tab
            handleClick(){
                this.drawEcharts()
            },
            // 鼠标移入
            mouseover(){
                
                this.autoRoll(true)
            },
            // 鼠标移出
            mouseout(){
                this.autoRoll()
            },

            // 设置自动滚动
            autoRoll(stop) {
            
            if(this.list.length!=1){
                if (stop) {
                
                clearInterval(rolltimer)
                return
                }

            
            
                
                // 拿到元素后，对元素进行定时增加距离顶部距离，实现滚动效果
                rolltimer = setInterval(() => {
                    // 元素自增距离顶部像素
                    
                    //console.log("sss123",this.activeName)
                    // 判断元素是否滚动到底部(可视高度+距离顶部=整个高度)
                    if ( parseInt(this.activeName) == this.list.length-1) {
                    // 重置table距离顶部距离
                        this.activeName="0";
                        this.key=(this.key+1)%2

                    }
                    else{
                        this.activeName=(parseInt(this.activeName)+1).toString()
                    }
                    
                    
                    
                    this.handleClick()
                    
                    
                }, this.rollTime * 10)
            }    
                
                
            
            
            
            
            },
            
            
            drawEcharts(){
                var se=[]
                for(let i=0;i<this.list.length;i++){
                    se.push(this.qsnumdata[this.list[i]])
                }
                      
                let ctab=parseInt(this.activeName)
                var se1=[]
                if(se[ctab]){
                    se1=Object.keys(se[ctab]).sort()
                    for(let i=1;i<=4;i++){
                        let myChart = echarts.getInstanceByDom(document.getElementById("mychart"+i));
                        if(i>se1.length){
                            if(myChart != null){
                                myChart.clear()
                            }
                            
                        }
                        else{
                            
                            
                            
                            
                            this.option.title.text=se1[i-1]+"数量"
                            this.option.xAxis.data=Object.keys( se[ctab][se1[i-1]])
                            this.option.series[0].data=Object.values(se[ctab][se1[i-1]])
                            
                            if (myChart == null) { // 如果不存在，就进行初始化。
                                myChart = echarts.init(document.getElementById("mychart"+i));

                                myChart.setOption(this.option, true);
                            }
                            else{

                                myChart.setOption(this.option,true)
                            }
                        }
                        
                    }
                }
                
                
                
                
            }
        }
     
    }
    
  
</script> 
<style scoped>
#centerLeft1{
    padding: 16px;
    height: 100%;
    width: 97%;
    
    
}
.charts1{
    display: grid;
    
    grid-template-columns: 50% 50%;
}
.charts2{
    display: grid;
    
    grid-template-columns: 50% 50%;
}
</style> 