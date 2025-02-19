<template>
  <div class="dashboard-container">
    
  <el-row :gutter="20">
    <el-col :span="8">
      <div class="grid-content bg-purple">
        
        
        <!-- 首页表格 -->
        <el-card shadow= 'hover' class="tableInfo">
          <div slot="header">
            <span class="important-font" style="white-space: pre-wrap;">{{ card1 }}</span>
            <span  style="margin-left: 160px; font-size: 50px;color: rgb(134, 184, 213);">{{ pronum }}</span>
          </div>          
        </el-card>
      </div>
    </el-col>
    <el-col :span="8">
      <div class="grid-content bg-purple">
        
        
        <!-- 首页表格 -->
        <el-card shadow= 'hover' class="tableInfo">
          <div slot="header">
            <span class="important-font" style="white-space: pre-wrap;">{{ card2 }}</span>
            <span  style="margin-left: 120px;font-size: 50px;color: rgb(134, 184, 213);">{{ tze }}</span>
          </div>          
        </el-card>
      </div>
    </el-col>
    <el-col :span="8">
      <div class="grid-content bg-purple">
        
        
        <!-- 首页表格 -->
        <el-card shadow= 'hover' class="tableInfo">
          <div slot="header">
            <span class="important-font" style="white-space: pre-wrap;">{{ card3 }}</span>
            <span  style="margin-left: 120px;font-size: 50px;color: rgb(134, 184, 213);">{{ lxcd }}</span>
          </div>          
        </el-card>
      </div>
    </el-col>   

      
  </el-row>
  <el-row>
    <el-col :span="16">
      <div class= "graph">
        <el-card shadow= 'hover' style="width:100%;height: 300px">
          <div id="lineEcharts" style="width: 100%;height: 270px;" >{{initLineEcharts()}}</div>
        </el-card>
      </div>
      

    </el-col>
    <el-col :span="8">
      <div class= "graph">
        <el-card shadow= 'hover' style="width:100%;height: 300px">
          <div id="pieEcharts" style="width: 100%;height: 300px;" >{{initPieEcharts()}}</div>
        </el-card>
      </div>
    </el-col>
  </el-row>
  <el-row>
    <el-table
          
          :data="tableData"
          border
          stripe
          style="width: 100%;max-height: 250px; margin-top: 20px;"
          height="250"
          :header-cell-style="{ 'text-align': 'center','color':'black' }"
          :cell-style="{ 'text-align': 'center','color':'black' }"
          
        >
          
          <el-table-column label="序号" width="50px">
            <template slot-scope="scope">
              {{ (page - 1) * limit+scope.$index + 1 }}
            </template>
          </el-table-column>
          
          <el-table-column prop="proName" label="项目名称" width="320px" />
          <el-table-column prop="gldj" label="公路等级" width="100px" />
          <el-table-column prop="tze" label="投资额(万元)" />
          <el-table-column prop="lxcd" label="路线长度(KM)" />
          <el-table-column prop="jgnf" label="交工年份" width="200px"/>
          <el-table-column prop="type" label="状态" :formatter="formatType"/>
          
          
          
        </el-table>
        <el-pagination
          :current-page="page"
          :total="total"
          :page-size="limit"
          :page-sizes="[4, 8]"
          style="padding: 30px 0; text-align: center"
          layout=" ->,total, sizes, prev, pager, next, jumper"
          @size-change="changePageSize"
          @current-change="changeCurrentPage"
        />
  </el-row>

      
      
  

  </div>
</template>

<script>
import { mapGetters } from 'vuex'
import Cookies from 'js-cookie'
import api from '@/api/xmtj.js'
import * as echarts from 'echarts';


export default {
  data(){
     return {
        
        
        username:Cookies.get("username"),
        type:Cookies.get("type"),

        tableData:[
          
        ],
        xmtj:[],
        pronum:0,
        tze:0,
        lxcd:0,
        total: 0, // 总记录数
        page: 1, // 当前页码
        limit: 4, // 每页记录数
        card1:"项目总数\n",
        card2:"总投资额(万元)\n",
        card3:"总路线长度(KM)\n",
        value:new Date()
      }
  },
  name: 'Dashboard',
  mounted(){
    this.getXmtj()
    
  },
  created(){
    this.fetchData()
  },
  computed: {
    ...mapGetters([
      'name'
    ])
  },
  methods:{
    formatType: function (row, column, cellValue) {
                var ret = ''  //你想在页面展示的值
                if (cellValue=='1') {
                    ret = "进行中"  //根据自己的需求设定
                } else if(cellValue=='2') {
                    ret = "已交工"
                }
                return ret;
            },

    getXmtj(){
      api.list().then((res) => {
            if(res.message=='成功'){
              
              this.xmtj=res.data
              this.pronum=this.xmtj.length
              for(let i=0;i<this.xmtj.length;i++){
                this.tze=this.tze+Number(this.xmtj[i].tze)
                this.lxcd=this.lxcd+Number(this.xmtj[i].lxcd)
              }
            }
            else{
              
            }
          });
    },
    fetchData(){
      api.pageList(this.page,this.limit).then((res) => {
            if(res.message=='成功'){
              this.tableData=res.data.records
              this.total = res.data.total;
            }
            else{
              
            }
          });
    },
    initLineEcharts(){
      let p = new Promise((resolve) => {
        resolve()
      })
      //然后异步执行echarts的初始化函数
      p.then(() => {
        
        
        let myChart = echarts.init(document.getElementById("lineEcharts"));
        
        let option= {
          title: {
            text: '折线图堆叠'
          },
          tooltip: {
              trigger: 'axis'
          },
          legend: {
            top: '0%',
            right: 'right'
          },
          grid: {
              left: '3%',
              right: '4%',
              bottom: '3%',
              containLabel: true
          },
          toolbox: {
              
          },
          xAxis: [
              {
                  type: "category",
                  data: [
                      
                  ],
                  axisTick: {
                      alignWithLabel: true
                  }
              }
          ],


          yAxis: [
              {
                  type: "value",
                  name: "路线长度(KM)",
                  nameLocation: "middle",
                  
                  nameTextStyle: {
                      padding: [
                          0,
                          0,
                          25,
                          30
                      ],
                      fontWeight: "bold",
                      fontFamily: "Segoe-UI-Bold",
                      fontSize: "13",
                      color: "#232253"
                  }
              },
              {
                  type: "value",
                  name: "投资额(万元)",
                  nameLocation: "middle",
                  
                  nameTextStyle: {
                      padding: [
                          60,
                          0,
                          0,
                          0
                      ],
                      fontWeight: "bold",
                      fontFamily: "Segoe-UI-Bold",
                      fontSize: "13",
                      color: "#232253"
                  },
                  position: "right"
              }
          ],


          series: [
              {   
                  
                  yAxisIndex: 1,
                  name: '投资额(万元)',
                  type: 'line',
                  
                  data: []
              },
              {   
                 
                  
                  name: '路线长度(KM)',
                  type: 'line',
                  
                  data: []
              },
              
              
              ]
            }

        this.xmtj.sort((a, b) => {
          return a.jgnf - b.jgnf;
        })
        for(let i=0;i<this.xmtj.length;i++){
          
          option.xAxis[0].data.push(this.xmtj[i].jgnf)
          
          option.series[0].data.push(this.xmtj[i].tze)
          option.series[1].data.push(this.xmtj[i].lxcd)
        }
        console.log("option",option)
        myChart.setOption(option);
      })
      
      
      },
    initPieEcharts(){
      let p = new Promise((resolve) => {
        resolve()
      })
      //然后异步执行echarts的初始化函数
      p.then(() => {
        let gs=0,l1=0,l2=0,l3 = 0,l4=0;
        for(let i=0;i<this.xmtj.length;i++){
          if(this.xmtj[i].gldj=="高速公路") gs++;
          else if(this.xmtj[i].gldj=="一级公路") l1++;
          else if(this.xmtj[i].gldj=="二级公路") l2++;
          else if(this.xmtj[i].gldj=="三级公路") l3++;
          else if(this.xmtj[i].gldj=="四级公路") l4++;
        }
        let myChart = echarts.init(document.getElementById("pieEcharts"));
        
        let option= {
          tooltip: {
            trigger: 'item'
          },
          legend: {
            top: '0%',
            left: 'left'
          },
          series: [
            {
              name: '公路数量',
              type: 'pie',
              radius: ['20%', '65%'],
              avoidLabelOverlap: false,
              label: {
                show: false,
                position: 'left'
              },
              labelLine: {
                show: false,
              },
              data: [
                {name:"高速公路",value:gs},
                {name:"一级公路",value:l1},
                {name:"二级公路",value:l2},
                {name:"三级公路",value:l3},
                {name:"四级公路",value:l4},
              ]
            }
          ]
        }
        
        myChart.setOption(option);
      })
      
      
      },
      // 改变每页显示的记录数
      changePageSize(size) {
          this.limit = size;
          this.fetchData();
        },
        // 改变页码数
        changeCurrentPage(page) {
          this.page = page;
          this.fetchData();
        },

    }
  
}
</script>

<style  scoped>
.dashboard-container {
  
    margin: 30px;
  
  
}
.el-carousel__item h3 {
    color: #475669;
    font-size: 18px;
    opacity: 0.75;
    line-height: 300px;
    margin: 0;
  }
  
  .el-carousel__item:nth-child(2n) {
    background-color: #99a9bf;
  }
  
  .el-carousel__item:nth-child(2n+1) {
    background-color: #d3dce6;
  }
  .el-card__body {
    padding: 10px;
}


.important-font{
    font-weight: 900;
    font-size: 25px;
}
.tableInfo{
  margin: 20px 0 0 0;
}

.el-card{
  border: none;
}
.graph{
  display: flex;
  flex-wrap: wrap;

  margin: 15px 0 0 0;
}

/deep/ .el-table, /deep/ .el-table__expanded-cell{
  background-color: transparent;
}
/* 表格内背景颜色 */
/deep/ .el-table th,
/deep/ .el-table tr{
  background-color: rgb(45, 202, 168);
}
/deep/ .el-table td {
  background-color: rgb(126, 139, 211);
}
/deep/ .el-table tbody tr { 
    pointer-events:none; 
}
</style>

