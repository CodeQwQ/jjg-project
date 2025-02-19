<template>
    <div id="center">
        
        <div class="charts" v-loading="loading" element-loading-text="拼命加载中"
                            element-loading-spinner="el-icon-loading"
                            element-loading-background="rgba(0, 0, 0, 0.8)">
                <div id="mychartc1" style="width: 450px;height:240px; margin-left: 100px; margin-top: 15px;">

                </div>
                <div id="mychartc2" style="width: 450px;height:240px; margin-left: 40px; margin-top: 15px;">

                </div>
        </div>
    </div>
    
    
  </template>
<script>
import * as echarts from 'echarts'
import api from '@/api/project/dp.js'
let titleTool=["合格率","完成率"]
    export default{
        data(){
            return{
                proname:this.$route.query.projecttitle,
                gchgl:[],
                gcwcl:[],
                loading:true,
                option :{
                    animation: false,
                    title:{
                        text: '主标题',
                        textStyle: {
                            color: 'red',
                            
                        },
                        x:"center",
                        y:"top"
                    },
                    tooltip:{
                        trigger:"item",
                        
                        
                            
                           
                        textStyle: {
                            
                            align:'center',
                            fontSize: 12
                        },
                        formatter: function(params){
                            
                            return params.name+"</br>"+params.seriesName+":"+params.data+"%"
                        }
                        
                    },
                    legend:{
                        data: ['△指标', '*指标',"其余指标","合计指标"],      //图例名称
                        right: "30%",                              //调整图例位置
                        bottom: 10,                                  //调整图例位置
                        itemHeight: 7,                      //修改icon图形大小
                        icon: 'circle',                         //图例前面的图标形状
                        textStyle: {                            //图例文字的样式
                            color: '#a1a1a1',               //图例文字颜色
                            fontSize: 12                      //图例文字大小
                        },


                    },
                    xAxis: {
                        data: ['交安工程', '桥梁工程', '路基工程', '路面工程', '隧道工程'],
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
                        type: 'value',  
                        axisLabel: {  
                            show: true,  
                            interval: 'auto',  
                            formatter: '{value} %' ,
                            textStyle:{
                                color:'#ffffff'
                            }
                            },  
                        show: true 
                    },
                    grid:{
                        left:50
                    },
                    series: [
                        {
                        name:"△指标",
                        type: 'bar',
                        data: [23, 24, 18, 25, 27, 28, 25],
                        

                        },
                        {
                        name:"*指标",
                        type: 'bar',
                        data: [26, 24, 18, 22, 23, 20, 27]
                        },
                        {
                        name:"其余指标",
                        type: 'bar',
                        data: [26, 24, 18, 22, 23, 20, 27]
                        },
                        {
                        name:"合计指标",
                        type: 'bar',
                        data: [26, 24, 18, 22, 23, 20, 27]
                        },
                        

                    ]
                }
            }
        },
        mounted(){
            this.$nextTick(() => {
            
                this.init()
                
            })
            
        },
        methods:{
            init() {
                api.getjsxmdata(this.proname).then((response) => {
                    
                    if(response.code==200){
                        
                        console.log("建设项目",response.data)
                        this.gchgl=response.data["建设项目合格率"]
                        this.gcwcl=response.data["建设项目实测指标完成率"]
                        console.log("建设项目",this.gchgl["交安工程"])
                        this.loading=false
                        this.drawEcharts()

                    }
                    

                })
                

                //this.getEcharts()
            },
            drawEcharts(){
                
                for(let i=1;i<=2;i++){
                    let myChart = echarts.getInstanceByDom(document.getElementById("mychartc"+i));
                    if(i==1){
                        var fbgcs=Object.keys(this.gchgl).sort()
                        var trig=[]
                        var star=[]
                        var qyzb=[]
                        var qbzb=[]
                        for(let i=0;i<fbgcs.length;i++){
                             if(i==1){
                                var tmp=" "
                                trig.push(tmp)
                                continue
                             }
                             var tmp=this.gchgl[fbgcs[i]][0]["hgl"]
                             trig.push(tmp)   
                        }
                        for(let i=0;i<fbgcs.length;i++){
                            if(i==1){
                                var tmp=this.gchgl[fbgcs[i]][0]["hgl"]
                                star.push(tmp)
                                continue
                            }
                            var tmp=this.gchgl[fbgcs[i]][1]["hgl"]
                            star.push(tmp)   
                       }
                       for(let i=0;i<fbgcs.length;i++){
                            if(i==1){
                                var tmp=this.gchgl[fbgcs[i]][1]["hgl"]
                                qyzb.push(tmp)
                                continue
                            }
                            var tmp=this.gchgl[fbgcs[i]][2]["hgl"]
                            qyzb.push(tmp)   
                       }
                       for(let i=0;i<fbgcs.length;i++){
                            if(i==1){
                                var tmp=this.gchgl[fbgcs[i]][1]["hgl"]
                                qbzb.push(tmp)
                                continue
                            }
                            var tmp=this.gchgl[fbgcs[i]][3]["hgl"]
                            qbzb.push(tmp)   
                       }
                        
                        
                        //console.log("test",star)
                        
                        this.option.title.text="建设项目合格率"
                        
                        this.option.series[0].data=trig
                        this.option.series[1].data=star
                        this.option.series[2].data=qyzb
                        this.option.series[3].data=qbzb
                    }
                    else{
                        var fbgcs=Object.keys(this.gcwcl).sort()
                        var trig=[]
                        var star=[]
                        var qyzb=[]
                        var qbzb=[]
                        for(let i=0;i<fbgcs.length;i++){
                             if(i==1){
                                var tmp=" "
                                trig.push(tmp)
                                continue
                             }
                             var tmp=this.gcwcl[fbgcs[i]][0]["wcl"]
                             trig.push(tmp)   
                        }
                        for(let i=0;i<fbgcs.length;i++){
                            if(i==1){
                                var tmp=this.gcwcl[fbgcs[i]][0]["wcl"]
                                star.push(tmp)
                                continue
                            }
                            var tmp=this.gcwcl[fbgcs[i]][1]["wcl"]
                            star.push(tmp)   
                       }
                       for(let i=0;i<fbgcs.length;i++){
                            if(i==1){
                                var tmp=this.gcwcl[fbgcs[i]][1]["wcl"]
                                qyzb.push(tmp)
                                continue
                            }
                            var tmp=this.gcwcl[fbgcs[i]][2]["wcl"]
                            qyzb.push(tmp)   
                       }
                       for(let i=0;i<fbgcs.length;i++){
                            if(i==1){
                                var tmp=this.gcwcl[fbgcs[i]][1]["wcl"]
                                qbzb.push(tmp)
                                continue
                            }
                            var tmp=this.gcwcl[fbgcs[i]][3]["wcl"]
                            qbzb.push(tmp)   
                       }
                        
                        
                        
                        //console.log("test",star)
                        
                        this.option.title.text="建设项目实测指标完成率"
                        
                        this.option.series[0].data=trig
                        this.option.series[1].data=star
                        this.option.series[2].data=qyzb
                        this.option.series[3].data=qbzb
                    }

                    if (myChart == null) { // 如果不存在，就进行初始化。
                        myChart = echarts.init(document.getElementById("mychartc"+i));
                        myChart.setOption(this.option, true);
                        console.log("ssss11")
                    }
                    else{
                        myChart.setOption(this.option,true)
                    }
                }
                
            }
        }
        
    }
</script> 
<style scoped>
#center{
    width:100%
}
.charts{
    display: grid;
    height: 100%;
    grid-template-columns: 50% 50%;
}
</style>
 