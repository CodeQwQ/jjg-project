<template>
    <div id="centerright1" v-loading="loading" element-loading-text="拼命加载中"
                            element-loading-spinner="el-icon-loading"
                            element-loading-background="rgba(0, 0, 0, 0.8)">
        <div class="fbgc">
            
            <el-button v-for="(item,index) in fbgcs" :key="index" @click="clickbutton(index)" :type="active===index?'primary':''" style="margin-top: 12px;margin-left: 10px;margin-right: 5px; width: 100px; ">
                {{ item }}
            </el-button>
            
        </div>
        <div class="right">
            <div class="charts">
                <div id="mychartr1" style="width: 360px;height:240px; margin-left: 25px; margin-top: 15px;">

                </div>
                <div id="mychartr2" style="width: 360px;height:240px; margin-left: 20px; margin-top: 15px;">

                </div>
            </div>
            <div class="tabs">
                <el-tabs  :key="key1" v-model="activeName1" class="mytabs1" @tab-click="handleClick" style="margin-left: 80px;color: white; height: 20px; width: 270px;caret-color: transparent;">
                
                    <el-tab-pane 
                        v-for="(item,index) in list1"
                        :key="index"
                        
                    ><span slot="label" @mouseover="mouseover1" @mouseout="mouseout1" style="color:rgb(104, 186, 32)" >{{ item }}</span></el-tab-pane>
                </el-tabs>
                <el-tabs  :key="key2" v-model="activeName2" class="mytabs2" @tab-click="handleClick" style="margin-left: 80px;color: white; height: 20px; width: 270px;caret-color: transparent;">
                
                <el-tab-pane 
                    v-for="(item,index) in list2"
                    :key="index"
                ><span slot="label" @mouseover="mouseover2" @mouseout="mouseout2" style="color:rgb(104, 186, 32)" >{{ item }}</span></el-tab-pane>
            </el-tabs>
            </div>
            
        </div>
        
    </div>
    
    
  </template>
<script>
import * as echarts from 'echarts'
import api from '@/api/project/dp.js'
let rolltimer1=''
let rolltimer2=''
let sum=0

    export default{
        data(){
            return{
                activeName1:"0",
                activeName2:"0",
                active:0,
                key1:0,
                key2:3,
                rollPx:1,
                proname:this.$route.query.projecttitle,
                fbgcs:[],
                gchgl:[],
                gcwcl:[],
                loading:true,
                dwgcdata:{},
                currentfbgc:0,
                rollTime: 150,
                refreshTime:5,
                list1:[
                    
                    
                    
                ],
                list2:[

                ],
                
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
                        data: ['1'],
                        axisLabel: {
                            //x轴文字的配置
                            show: true,
                            interval: 0,//使x轴文字显示全
                            textStyle:{
                                color:'#ffffff'
                            }
                        }
                    },
                    yAxis: {
                        axisLabel: {
                            //x轴文字的配置
                            show: true,
                            interval: 0,//使x轴文字显示全
                            textStyle:{
                                color:'#ffffff'
                            }
                        }
                    },
                    grid:{
                        left:50
                    },
                    series: [
                        {
                        type: 'bar',
                        data: [23],
                        barWidth: '25%',
                        
                        label: {
                            show: true, //开启显示
                            position: 'top', //数值展示的位置
                            textStyle: {
                                color: '#00ffff',
                                fontSize: 13
                            },
                            //echartjs 2.0 设置显示的数据 echartjs 3.0更简易formatter: '{c},({d}%)'
                            formatter: function(params) {
                                if(sum==0){
                                    return 0+'%'
                                }
                                return  ((params.data/sum)*100).toFixed(2)+'%'
                            }
                        },

                        itemStyle: {
                                    normal: {
                        　　　　　　　　//这里是重点
                                        color: function(params) {
                                            //注意，如果颜色太少的话，后面颜色不会自动循环，最好多定义几个颜色
                                            var colorList = ['#c23531','#2f4554', '#61a0a8', '#d48265', '#91c7ae','#749f83', '#ca8622'];
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
            // 初始化echarts实例
                this.init()
                
            })
            
           

        },
        destroyed(){
            this.autoRoll(1)
            this.autoRoll(2)
        },
        methods:{
            init() {
                api.getdwgc(this.proname).then((response) => {
                    
                    if(response.code==200){
                        console.log("单位工程",response.data)
                        this.fbgcs=Object.keys(response.data).sort()
                        this.dwgcdata=response.data
                        
                        this.loading=false
                        this.gchgl=response.data[this.fbgcs[this.currentfbgc]]["单位工程合格率"]
                        this.gcwcl=response.data[this.fbgcs[this.currentfbgc]]["单位工程指标完成率"]
                        
                        var htds=Object.keys(this.gchgl).sort()
                        
                        for(let i=0;i<htds.length;i++){
                            this.list1.push(htds[i])
                            
                        }
                        if(this.gcwcl!=null){
                            htds=Object.keys(this.gcwcl).sort()
                            for(let i=0;i<htds.length;i++){
                                this.list2.push(htds[i])
                                
                            }
                        }
                         
                        
                        if(this.list1.length>1){
                            this.autoRoll(-1)
                            
                            
                        }
                        if(this.list2.length>1){
                            this.autoRoll(-2)
                            
                            
                            
                        }
                        this.drawEcharts(1)
                        this.drawEcharts(2)


                    }
                    

                })
                

                //this.getEcharts()
            },
            // 点击按钮
            clickbutton(index){
                this.currentfbgc=index
                this.gchgl=this.dwgcdata[this.fbgcs[this.currentfbgc]]["单位工程合格率"]
                this.gcwcl=this.dwgcdata[this.fbgcs[this.currentfbgc]]["单位工程指标完成率"]
                this.autoRoll(1)
                
                this.activeName1="0"
                this.active=index
                var htds=Object.keys(this.gchgl).sort()
                this.list1=[]
                this.list2=[]
                for(let i=0;i<htds.length;i++){
                    this.list1.push(htds[i])
                }
                if(this.gcwcl!=null){
                    this.autoRoll(2)
                    this.activeName2="0"
                    
                    htds=Object.keys(this.gcwcl).sort()
                    for(let i=0;i<htds.length;i++){
                        this.list2.push(htds[i])
                    }
                }
                
                
                
                if(this.list1.length>1){
                    this.autoRoll(-1)
                    
                    
                }
                if(this.list2.length>1){
                    this.autoRoll(-2)
                    
                    
                }
                else{
                    this.handleClick()
                }
                
            },
            // 点击导航栏
            handleClick(){
                this.drawEcharts(1)
                this.drawEcharts(2)
            },
            mouseover1(){
               if(this.list1.length>1){
                this.autoRoll(1)
               }
                
            },
            // 鼠标移出
            mouseout1(){
                if(this.list1.length>1){
                    this.autoRoll(-1)
                }
                
            },
            mouseover2(){
                if(this.list2.length>1){
                    this.autoRoll(2)
                }
                
            },
            // 鼠标移出
            mouseout2(){
                if(this.list2.length>1){
                    this.autoRoll(-2)
                }
                
            },
            // 设置自动滚动
            autoRoll(stop) {
                
                    if (stop==1) {
                        
                        clearInterval(rolltimer1)
                        return
                    }
                    else if(stop==2){
                        clearInterval(rolltimer2)
                        return
                    }
                    else if(stop==-1){
                        rolltimer1 = setInterval(() => {
                            // 元素自增距离顶部像素
                            
                        
                            //console.log("sss123",this.activeName)
                            // 判断元素是否滚动到底部(可视高度+距离顶部=整个高度)
                            if ( parseInt(this.activeName1) == this.list1.length-1) {
                            // 重置table距离顶部距离
                                this.activeName1="0";
                                this.key1=(this.key1+1)%2

                            }
                            else{
                                this.activeName1=(parseInt(this.activeName1)+1).toString()
                            }
                            
                            //this.option.series[0].data[0]++
                            this.handleClick()
                        }, this.rollTime * 10)
                    }
                    else{
                        // 拿到元素后，对元素进行定时增加距离顶部距离，实现滚动效果
                    
                        rolltimer2 = setInterval(() => {
                            // 元素自增距离顶部像素
                            
                        
                            //console.log("sss123",this.activeName)
                            // 判断元素是否滚动到底部(可视高度+距离顶部=整个高度)
                            if ( parseInt(this.activeName2) == this.list2.length-1) {
                            // 重置table距离顶部距离
                                this.activeName2="0";
                                if(this.key%2==0){
                                    this.key2--
                                }
                                else{
                                    this.key2++
                                }
                                

                            }
                            else{
                                this.activeName2=(parseInt(this.activeName2)+1).toString()
                            }
                            
                            //this.option.series[0].data[0]++
                            this.handleClick()
                        }, this.rollTime * 10)
                    }
                
                
                
            },
            drawEcharts(index){
                    let myChart = echarts.getInstanceByDom(document.getElementById("mychartr"+index));
                    if(myChart != null){
                            myChart.clear()
                    }
                    if(index==1){
                        
                        var data=[]
                        for(let i=0;i<this.list1.length;i++){
                            data.push(this.gchgl[this.list1[i]][0])
                        }
                        var xdata=[]
                        
                        var ctab=parseInt(this.activeName1)
                        xdata.push(data[ctab]["hgds"])
                        xdata.push(data[ctab]["zds"]-data[ctab]["hgds"])
                        sum=data[ctab]["zds"]
                        
                        var labels=["合格点数","不合格点数"]
                        this.option.title.text="单位工程合格率"
                        this.option.xAxis.data=labels
                        this.option.series[0].data=xdata
                    }
                    else if(index==2&&this.gcwcl!=null){
                        var data=[]
                        for(let i=0;i<this.list2.length;i++){
                            data.push(this.gcwcl[this.list2[i]][0])
                        }
                        var xdata=[]
                        
                        var ctab=parseInt(this.activeName2)
                        xdata.push(data[ctab]["jcs"])
                        xdata.push(data[ctab]["zs"]-data[ctab]["jcs"])
                        sum=data[ctab]["zs"]
                        
                        var labels=["检测点数","未检测点数"]
                        this.option.title.text="单位工程指标完成率"
                        this.option.xAxis.data=labels
                        this.option.series[0].data=xdata 
                    }
                    else{
                        return
                    }
                    if (myChart == null) { // 如果不存在，就进行初始化。
                            myChart = echarts.init(document.getElementById("mychartr"+index));
                            
                            myChart.setOption(this.option, true);
                        
                    }
                    else{
                        myChart.setOption(this.option,true)
                    }
                     /*for(let i=1;i<=2;i++){
                        let myChart = echarts.getInstanceByDom(document.getElementById("mychartr"+i));
                        if(myChart != null){
                                myChart.clear()
                        }
                        
                        if(i==1){
                            var data=[]
                            for(let i=0;i<this.list1.length;i++){
                                data.push(this.gchgl[this.list1[i]][0])
                            }
                            var xdata=[]
                            
                            var ctab=parseInt(this.activeName1)
                            xdata.push(data[ctab]["hgds"])
                            xdata.push(data[ctab]["zds"]-data[ctab]["hgds"])
                            sum=data[ctab]["zds"]
                            
                            var labels=["合格点数","不合格点数"]
                            this.option.title.text="单位工程合格率"
                            this.option.xAxis.data=labels
                            this.option.series[0].data=xdata
                            
                        }
                        else if(this.gcwcl!=null&&i==2){
                            var data=[]
                            for(let i=0;i<this.list2.length;i++){
                                data.push(this.gcwcl[this.list2[i]][0])
                            }
                            var xdata=[]
                            
                            var ctab=parseInt(this.activeName1)
                            xdata.push(data[ctab]["jcs"])
                            xdata.push(data[ctab]["zs"]-data[ctab]["jcs"])
                            sum=data[ctab]["zs"]
                            
                            var labels=["检测点数","未检测点数"]
                            this.option.title.text="单位工程指标完成率"
                            this.option.xAxis.data=labels
                            this.option.series[0].data=xdata 
                        }
                        else{
                            return 
                        }
                        if (myChart == null) { // 如果不存在，就进行初始化。
                            myChart = echarts.init(document.getElementById("mychartr"+i));
                            
                            myChart.setOption(this.option, true);
                            
                        }
                        else{
                            myChart.setOption(this.option,true)
                        } 
                    }*/
                
                //获取已有echarts实例的DOM节点。
                
                
            }
        }
     
    }
    
  
</script>
<style scoped>
#centerright1{
    
    height: 95%;
    width: 100%;
    display: grid;
    grid-template-columns: 10% 90%;
    
    
}
.fbgc{
    margin-left: 15px;
    height: 100%;
    width: 100px;
    display: grid;
    grid-template-rows: 20% 20% 20% 20% 20%;
}
.charts{
    display: grid;
    height: 240px;
    grid-template-columns: 50% 50%;
}
.right{
    display: grid;
    grid-template-rows: 85% 15%;
}
.tabs{
    display: grid;
    grid-template-columns: 50% 50%;
}
</style>  