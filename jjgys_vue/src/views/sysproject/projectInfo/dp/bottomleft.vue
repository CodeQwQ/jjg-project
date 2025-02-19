<template>
    
    <div id="myChart" :style="{}"></div>
    
  </template>
<script>
    export default{
        
        methods:{
            drawEcharts(){
                let myChart = this.$echarts.init(document.getElementById('myChart'))
                
            }
        }
     
    }
    
  
</script>  