<template>
    
    <div id="myChart" :style="{width: '300px', height: '300px'}"></div>
    
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