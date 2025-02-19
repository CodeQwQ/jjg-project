<template>
    
    <div id="centerright2" v-loading="loading" element-loading-text="拼命加载中"
                            element-loading-spinner="el-icon-loading"
                            element-loading-background="rgba(0, 0, 0, 0.8)">
        <div class="tabs" style="height: 230px; width: 30px;">
            <el-tabs  tab-position="left" :key="key" v-model="activeName" class="mytabs" @tab-click="handleClick" style="color: white; height: 250px; width: 100px;margin-top: 20px; caret-color: transparent;">
                <el-tab-pane 
                    v-for="(item,index) in list"
                    :key="index"
                ><span slot="label" @mouseover="mouseover" @mouseout="mouseout" style="color:rgb(104, 186, 32)" >{{ item }}</span></el-tab-pane>
            </el-tabs>
        </div>
        <div class="right" >
            
            <div class="charts">
                <div id="mychartr_1" style="width: 350px;height:240px; margin-left: 35px; margin-top: 15px;">

                </div>
                <div id="mychartr_2" style="width: 350px;height:240px; margin-left: 20px; margin-top: 15px;">

                </div>
            </div>
            
            
            
                
                
                
            
            <div class="buttons">
                
                <div   style="position:relative;margin-top: 10px; margin-left: 82px; height: 100px;">
                    
                    <el-button   v-show="buttonl" @mouseover.native = "mouseover" @mouseout.native="mouseout" class="button" v-for="(item,index) in fbgcs" :key="index" @click.native="clickbuttonl(index)" :type="activel===index?'primary':''" >
                        {{ item }}
                    </el-button>
                </div>
                <div  style="position:relative;margin-top: 10px;margin-left: 73px;height: 100px;">
                    <el-button   v-show="buttonr" @mouseover.native = "mouseover" @mouseout.native="mouseout" class="button" v-for="(item,index) in fbgcs" :key="index" @click.native="clickbuttonr(index)" :type="activer===index?'primary':''" >
                        {{ item }}
                    </el-button>
                </div>
                

            </div>
        </div>
        
    </div>
    
  </template>
<script>
import * as echarts from 'echarts'
import api from '@/api/project/dp.js'
let rolltimer=''
let sum1=0
let sum2=0
let flag=0
    export default{
        data(){
            return{
                activeName:"0",
                
                key:0,
                keyb:0,
                activel:0,
                activer:0,
                buttonl:true,
                buttonr:true,
                rollPx:1,
                rollTime: 150,
                refreshTime:5,
                loading:true,
                proname:this.$route.query.projecttitle,
                list:[
                    
                    
                    
                ],
                fbgcs:[],
                fbgcsall:[],
                htddata:{},
                gchgl:{},
                gcwcl:{},
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
                        data: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
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
                        data: [23, 24, 18, 25, 27, 28, 25],
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
                                var sumnow=0
                                if(flag==0){
                                    sumnow=sum1
                                }
                                else{
                                    sumnow=sum2
                                }
                                if(sumnow==0){
                                    return 0+'%'
                                }
                                return  ((params.data/sumnow)*100).toFixed(2)+'%'
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
            this.autoRoll(true)
            
        },
        methods:{
            init() {
                api.gethtddata(this.proname).then((response) => {
                    
                    if(response.code==200){
                        console.log("合同段",response.data)
                        this.htddata=response.data
                        this.list=Object.keys( response.data).sort()
                        this.loading=false
                        this.gchgl=response.data[this.list[0]]["合同段合格率"]
                        this.gcwcl=response.data[this.list[0]]["合同段指标完成率"]
                        if(this.gchgl!=null){
                            this.fbgcsall=Object.keys(this.gchgl).sort()
                        }
                        else{
                            this.fbgcsall=Object.keys(this.gcwcl).sort()
                        }
                        

                        for(let i=0;i<this.fbgcsall.length;i++){
                            this.fbgcs.push( this.fbgcsall[i].substr(0,2))
                            
                        }
                        this.autoRoll()
                        
                            this.drawEcharts(1)
                        
                        
                            this.drawEcharts(2)
                        
                        
                        


                    }

                })
            },    
            // 点击tab
            handleClick(){
                var ctab=parseInt(this.activeName)
                this.gchgl=this.htddata[this.list[ctab]]["合同段合格率"]
                this.gcwcl=this.htddata[this.list[ctab]]["合同段指标完成率"]
                
                if(this.gchgl!=null){
                    this.fbgcsall=Object.keys(this.gchgl).sort()
                }
                else{
                    
                    this.fbgcsall=Object.keys(this.gcwcl).sort()
                }
                this.fbgcs=[]
                for(let i=0;i<this.fbgcsall.length;i++){
                    this.fbgcs.push( this.fbgcsall[i].substr(0,2))
                    
                }
                
                this.drawEcharts(1)
                
                
                this.drawEcharts(2)
                
            },
            // 点击按钮
            clickbuttonl(index){
               
                this.activel=index
                this.drawEcharts(1)
            },
            clickbuttonr(index){
                this.activer=index
                this.drawEcharts(2)
            },
            mouseover(){
                console.log("进入")
                this.autoRoll(true)
            },
            // 鼠标移出
            mouseout(){
                this.autoRoll()
            },
            
            // 设置自动滚动
            autoRoll(stop) {
                if (stop) {
                    
                    clearInterval(rolltimer)
                    return
                }
                
                
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
                        //this.option.series[0].data[0]++
                        this.handleClick()
                    }, this.rollTime * 10)
            },
                
                
            
            drawEcharts(index){
                let myChart = echarts.getInstanceByDom(document.getElementById("mychartr_"+index));
                if(myChart != null){
                        myChart.clear()
                }
                if(index==1){
                    var data=[]
                    
                    if(this.gchgl==null){
                        this.buttonl=false
                        
                        
                        return ;
                    }
                    this.buttonl=true
                    data=this.gchgl[this.fbgcsall[this.activel]][0]
                    
                    
                    sum1=data["zds"]
                    flag=0
                    var xdata=[]
                    xdata.push(data["hgds"])
                    xdata.push(data["zds"]-data["hgds"])
                    //console.log("fbgc",xdata)
                    var labels=["合格点数","不合格点数"]
                    this.option.title.text="合同段合格率"
                    this.option.xAxis.data=labels
                    this.option.series[0].data=xdata
                }
                else{
                    var data=[]
                    var xdata=[]
                    
                    if(this.gcwcl==null){
                        this.buttonr=false
                        
                        return ;
                    }
                    this.buttonr=true
                    if(this.gcwcl[this.fbgcsall[this.activer]]!=null){
                        data=this.gcwcl[this.fbgcsall[this.activer]][0]
                        sum2=data["zs"]
                        flag=1
                        
                        xdata.push(data["jcs"])
                        xdata.push(data["zs"]-data["jcs"])
                        
                        
                    }
                    else{
                        data=[]
                        xdata=[0,0]
                    }
                        var labels=["检测点数","未检测点数"]
                        this.option.title.text="合同段指标完成率"
                        this.option.xAxis.data=labels
                        this.option.series[0].data=xdata
                    
                    
                }
                
                   
                    if (myChart == null) { // 如果不存在，就进行初始化。
                        myChart = echarts.init(document.getElementById("mychartr_"+index));
                        myChart.setOption(this.option, true);
                        console.log("ssss11")
                    }
                    else{
                        myChart.setOption(this.option,true)
                    }
                
                
            }
        }
     
    }
    
  
</script>
<style scoped>
#centerright2{
    height: 95%;
    width: 100%;
    display: grid;
    grid-template-columns: 10% 90%;
}
.right{
    display: grid;
    height: 240px;
    grid-template-rows: 85% 12%;
}
.charts{
    display: grid;
    height: 240px;
    grid-template-columns: 50% 50%;
    
}
.buttons{
    display: grid;
    height: 240px;
    grid-template-columns: 50% 50%;
}
.button{
        width: 50px;
        font-size: 10px;
        text-align: left;
        padding-left: 13px;
	    
}
</style>  