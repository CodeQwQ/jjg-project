<template>
    <div class="app-container">
        <!--查询表单-->
        <div class="search-div">
          <el-form :inline="true" class="demo-form-inline">
            <el-form-item label="试验时间:">
              <!-- <el-input v-model="searchObj.jcsj" /> -->
              <el-date-picker
                v-model="searchObj.sysj"
                type="date"
                value-format="yyyy-MM-dd"
              />
            </el-form-item>
            <el-form-item label="桩号:">
              <el-input v-model="searchObj.zh" />
            </el-form-item>
            <el-button
              type="primary"
              icon="el-icon-search"
              size="mini"
              @click="fetchData()"
              >搜索</el-button
            >
            <el-button icon="el-icon-refresh" size="mini" @click="resetData"
              >重置</el-button
            >
          </el-form>
        </div>
        <!-- 工具按钮 -->
        <div >
            <el-switch
            v-model="flag"
            @change="fetchData(),page=1,limit=10"
            active-text="灰土数据"
            inactive-text="沙砾数据"
            active-color="#13ce66"
            inactive-color="#13ce66"
            >

            </el-switch>
        </div>
        <div class="tools-div" v-show="!(username=='admin'&&type==1  )">
          <div class="anniu">
            
            <el-button
              class="btn-add"
              type="success"
              @click="exportysdht()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              v-if="flag==true"
              >导出压实度灰土模板文件</el-button
            >
            <el-button
              class="btn-add"
              type="success"
              @click="exportysdsl()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              v-if="flag==false"
              >导出压实度沙砾模板文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importysdht()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              v-if="flag==true"
              >导入压实度灰土数据文件</el-button
            >
            <el-button
              class="btn-add"
              type="success"
              @click="importysdsl()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              v-if="flag==false"
              >导入压实度沙砾数据文件</el-button
            >
            <el-button
              class="btn-add"
              type="success"
              size="mini"
              @click="scjdb()"
              style="margin-left: 1px"
              icon="el-icon-top"
              >生成鉴定表</el-button
            >
            <el-button
              class="btn-add"
              type="success"
              size="mini"
              @click="ckjdb()"
              style="margin-left: 1px"
              icon="el-icon-top"
              >查看检测结果</el-button
            >
            <el-button
            class="btn-add"
            type="primary"
            size="mini"
            @click="modify()"
            style="margin-left: 1px"
            icon="el-icon-edit"
            >修改</el-button>
            <el-button
              class="btn-add"
              type="danger"
              size="mini"
              @click="batchRemove()"
              style="margin-left: 1px"
              icon="el-icon-delete"
              >删除</el-button
            >
            <el-button
            class="btn-add"
            type="danger"
            size="mini"
            @click="removeAll()"
            style="margin-left: 1px"
            icon="el-icon-delete"
            >全部删除</el-button>
            <el-button class="shuoming" style="float：right" type="text" @click="dialogVisible = true"
              >填写说明</el-button
            >
          </div>
        </div>
        <!-- 表格 -->
        <el-table
          v-if="flag==true"
          :data="list"
          border
          stripe
          style="width: 100%;"
          class="table"
          :header-cell-style="{ 'text-align': 'center' }"
          :cell-style="{ 'text-align': 'center' }"
          @selection-change="handleSelectionChange"
        >
          <el-table-column type="selection" />
          <el-table-column label="序号" width="50px">
            <template slot-scope="scope">
              {{ (page - 1) * limit + scope.$index + 1 }}
            </template>
          </el-table-column>
          
          <el-table-column prop="sysj" label="试验时间" width="100px" />
          <el-table-column prop="htgdz" label="路基压实度_灰土规定值" width="220px" />
          <el-table-column prop="zh" label="桩号" width="250px"/>
          <el-table-column prop="jgcc" label="结构层次" />
          <el-table-column prop="jglx" label="结构类型" />
          <el-table-column prop="zdgmd" label="最大干密度" width="100px"/>
          <el-table-column prop="bzsmd" label="标准砂密度" width="100px" />
          <el-table-column prop="qyzhjwz" label="取样桩号及位置" width="200px" />
          <el-table-column prop="sksd" label="试坑深度(cm)" width="130px"/>
          <el-table-column prop="ztjjbhbmjszl" label="锥体及基板和表面间砂质量(g)" width="280px"/>
          <el-table-column prop="gsqtszl" label="灌砂前筒+砂质量(g)" width="150px"/>
          <el-table-column prop="gshtszl" label="灌砂后筒+砂质量(g)" width="150px"/>
          <el-table-column prop="syzl" label="试样质量(g)" width="100px"/>
          <el-table-column prop="hh" label="盒号" />
          <el-table-column prop="hzl" label="盒质量(g)" width="100px"/>
          <el-table-column prop="hsshzl" label="盒+湿试样质量(g)" width="160px" />
          <el-table-column prop="hgsyzl" label="盒+干试样质量(g)" width="160px"/>
          <el-table-column prop="xh" label="序号" />
          
          <el-table-column prop="createtime" label="创建时间" width="160px"/>
          
        </el-table>
        
        <el-table
          v-if="flag==false"
          :data="list"
          border
          stripe
          style="width: 100%;"
          class="table"
          :header-cell-style="{ 'text-align': 'center' }"
          :cell-style="{ 'text-align': 'center' }"
          @selection-change="handleSelectionChange"
        >
          <el-table-column type="selection" />
          <el-table-column label="序号" width="50px">
            <template slot-scope="scope">
              {{ (page - 1) * limit + scope.$index + 1 }}
            </template>
          </el-table-column>
          
          <el-table-column prop="sysj" label="试验时间" width="100px" />
          <el-table-column prop="slgdz" label="路基压实度_沙砾规定值" width="220px" />
          <el-table-column prop="zh" label="桩号" width="250px"/>
          <el-table-column prop="a" label="a" width="200px"/>
          <el-table-column prop="b" label="b" width="200px"/>
          <el-table-column prop="c" label="c" width="200px"/>
          
          <el-table-column prop="qyzhjwz" label="取样桩号及位置" width="200px" />
          <el-table-column prop="sksd" label="试坑深度(cm)" width="130px"/>
          
          <el-table-column prop="gsqtszl" label="灌砂前筒+砂质量(g)" width="150px"/>
          <el-table-column prop="gshtszl" label="灌砂后筒+砂质量(g)" width="150px"/>
          <el-table-column prop="ztsz" label="锥体砂重" width="100px"/>
          <el-table-column prop="lsdmd" label="量砂的密度" width="100px"/>
          <el-table-column prop="hhldszl" label="混合料的湿重量" width="200px"/>
          <el-table-column prop="hh" label="盒号" />
          <el-table-column prop="hgzl" label="盒+干质量" width="100px"/>
          <el-table-column prop="hzl" label="盒质量" width="100px"/>
          <el-table-column prop="klzl" label="5-38mm颗粒质量(g)" width="160px" />
          
          <el-table-column prop="xh" label="序号" />
          
          <el-table-column prop="createtime" label="创建时间" width="160px"/>
        </el-table>
        <!-- 分页组件 -->
        <el-pagination
          :current-page="page"
          :total="total"
          :page-size="limit"
          :page-sizes="[5, 10, 20, 30, 40, 50, 100]"
          style="padding: 30px 0; text-align: center"
          layout=" ->,total, sizes, prev, pager, next, jumper"
          @size-change="changePageSize"
          @current-change="changeCurrentPage"
        />
    
        <el-dialog title="导入" :visible.sync="dialogImportVisible" width="480px">
          <el-form label-position="right" label-width="170px">
            <el-form-item label="文件">
              <el-upload
                :multiple="false"
                :on-change="fileChange"
                :auto-upload="false"
                :file-list="fileList"
                :limit="1"
                action="/"
                class="upload-demo"
              >
                <el-button size="small" type="primary">点击上传</el-button>
                <div slot="tip" class="el-upload__tip">只能上传xls文件</div>
              </el-upload>
            </el-form-item>
          </el-form>
          <div slot="footer" class="dialog-footer">
            <el-button @click="onUploadSuccess">提交</el-button>
            <el-button @click="dialogImportVisible = false,fileList=[]">取消</el-button>
          </div>
        </el-dialog>
    
        <el-dialog
          title="填写说明"
          :visible.sync="dialogVisible"
          width="30%"
          
          
        >
        <span style="white-space: pre-wrap;">{{ instruction }}</span>
      
        <img v-viewer v-if="flag" style="height: 100px;width: 450px;" src="@/assets/填写说明/路基/路基土石方压实度灰土.png" alt="" >
        
        <img v-viewer v-if="!flag" style="height: 100px;width: 450px;" src="@/assets/填写说明/路基/路基土石方压实度沙砾.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="路基土石方检测结果"
         :visible.sync="dialogVisible1">
        
        <el-table
          
          :data="descriptions"
          border
          stripe
          style="width: 100%;"
          class="table"
          :header-cell-style="{ 'text-align': 'center' }"
          :cell-style="{ 'text-align': 'center' }"
          
        >
        <el-table-column prop="压实度项目" label="压实度项目" />
          <el-table-column prop="代表值" label="代表值"  />
          <el-table-column prop="检测点数" label="检测点数" />
          <el-table-column prop="合格点数" label="合格点数" />
          <el-table-column prop="合格率" label="合格率(%)" />
          <el-table-column prop="规定值" label="规定值" />
          <el-table-column prop="平均值" label="平均值" />
          <el-table-column prop="标准差" label="标准差" />
          
          <el-table-column prop="结果" label="结果"  />
         
          
          
        </el-table>
         <span slot="footer" class="dialog-footer">
          <el-button type="primary"  @click="dialogVisible1 = false">关闭</el-button>
         </span>
         
        </el-dialog>
        <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm1" v-if="flag==true" :model="modifyData"  label-width="220px" size="small" style="padding-right: 40px;" :rules="rules1">
            
            <el-form-item label="试验时间" prop="sysj">
              <el-input v-model="modifyData.sysj"></el-input>
            </el-form-item>
            <el-form-item label="灰土规定值" prop="htgdz">
              <el-input v-model="modifyData.htgdz"></el-input>
            </el-form-item>
            <el-form-item label="桩号" prop="zh">
              <el-input v-model="modifyData.zh"></el-input>
            </el-form-item>
            <el-form-item label="结构层次" prop="jgcc">
              <el-input v-model="modifyData.jgcc"></el-input>
            </el-form-item>
            <el-form-item label="结构类型" prop="jglx">
              <el-input v-model="modifyData.jglx"></el-input>
            </el-form-item>
            <el-form-item label="最大干密度" prop="zdgmd">
              <el-input v-model="modifyData.zdgmd"></el-input>
            </el-form-item>
            <el-form-item label="标准砂密度" prop="bzsmd">
              <el-input v-model="modifyData.bzsmd"></el-input>
            </el-form-item>
            <el-form-item label="取样桩号及位置" prop="qyzhjwz">
              <el-input v-model="modifyData.qyzhjwz"></el-input>
            </el-form-item>
            <el-form-item label="试坑深度(cm)" prop="sksd">
              <el-input v-model="modifyData.sksd"></el-input>
            </el-form-item>
            <el-form-item label="锥体及基板和表面间砂质量(g)" prop="ztjjbhbmjszl">
              <el-input v-model="modifyData.ztjjbhbmjszl"></el-input>
            </el-form-item>
            <el-form-item label="灌砂前筒+砂质量(g)" prop="gsqtszl">
              <el-input v-model="modifyData.gsqtszl"></el-input>
            </el-form-item>
            <el-form-item label="灌砂后筒+砂质量(g)" prop="gshtszl">
              <el-input v-model="modifyData.gshtszl"></el-input>
            </el-form-item>
            <el-form-item label="试样质量(g)" prop="syzl">
              <el-input v-model="modifyData.syzl"></el-input>
            </el-form-item>
            <el-form-item label="盒号" prop="hh">
              <el-input v-model="modifyData.hh"></el-input>
            </el-form-item>
            <el-form-item label="盒质量(g)" prop="hzl">
              <el-input v-model="modifyData.hzl"></el-input>
            </el-form-item>
            <el-form-item label="盒+湿试样质量(g)" prop="hsshzl">
              <el-input v-model="modifyData.hsshzl"></el-input>
            </el-form-item>
            <el-form-item label="盒+干试样质量(g)" prop="hgsyzl">
              <el-input v-model="modifyData.hgsyzl"></el-input>
            </el-form-item>
            <el-form-item label="序号" prop="xh">
              <el-input v-model="modifyData.xh"></el-input>
            </el-form-item>
            
            <el-form-item label="创建时间" prop="createtime">
              <el-input v-model="modifyData.createtime"></el-input>
            </el-form-item>
            
          </el-form>
          <el-form ref="dataForm2" v-if="flag==false" :model="modifyData"  label-width="150px" size="small" style="padding-right: 40px;" :rules="rules2">
            
            <el-form-item label="试验时间" prop="sysj">
              <el-input v-model="modifyData.sysj"></el-input>
            </el-form-item>
            <el-form-item label="沙砾规定值" prop="slgdz">
              <el-input v-model="modifyData.slgdz"></el-input>
            </el-form-item>
            <el-form-item label="桩号" prop="zh">
              <el-input v-model="modifyData.zh"></el-input>
            </el-form-item>
            <el-form-item label="a" prop="a">
              <el-input v-model="modifyData.a"></el-input>
            </el-form-item>
            <el-form-item label="b" prop="b">
              <el-input v-model="modifyData.b"></el-input>
            </el-form-item>
            <el-form-item label="c" prop="c">
              <el-input v-model="modifyData.c"></el-input>
            </el-form-item>
            <el-form-item label="取样桩号及位置" prop="qyzhjwz">
              <el-input v-model="modifyData.qyzhjwz"></el-input>
            </el-form-item>
            <el-form-item label="试坑深度(cm)" prop="sksd">
              <el-input v-model="modifyData.sksd"></el-input>
            </el-form-item>
            <el-form-item label="灌砂前筒+砂质量(g)" prop="gsqtszl">
              <el-input v-model="modifyData.gsqtszl"></el-input>
            </el-form-item>
            <el-form-item label="灌砂后筒+砂质量(g)" prop="gshtszl">
              <el-input v-model="modifyData.gshtszl"></el-input>
            </el-form-item>
            <el-form-item label="锥体砂重" prop="ztsz">
              <el-input v-model="modifyData.ztsz"></el-input>
            </el-form-item>
            <el-form-item label="量砂的密度" prop="lsdmd">
              <el-input v-model="modifyData.lsdmd"></el-input>
            </el-form-item>
            <el-form-item label="混合料的湿重量" prop="hhldszl">
              <el-input v-model="modifyData.hhldszl"></el-input>
            </el-form-item>
            <el-form-item label="盒号" prop="hh">
              <el-input v-model="modifyData.hh"></el-input>
            </el-form-item>
            <el-form-item label="盒+干质量" prop="hgzl">
              <el-input v-model="modifyData.hgzl"></el-input>
            </el-form-item>
            <el-form-item label="盒质量" prop="hzl">
              <el-input v-model="modifyData.hzl"></el-input>
            </el-form-item>
            <el-form-item label="5-38mm颗粒质量(g)" prop="klzl">
              <el-input v-model="modifyData.klzl"></el-input>
            </el-form-item>
            <el-form-item label="序号" prop="xh">
              <el-input v-model="modifyData.xh"></el-input>
            </el-form-item>
            
            <el-form-item label="创建时间" prop="createtime">
              <el-input v-model="modifyData.createtime"></el-input>
            </el-form-item>
            
          </el-form>
          <span slot="footer" class="dialog-footer">
            <el-button @click="dialogVisible2 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
            <el-button type="primary" icon="el-icon-check" @click="update()" size="small">确 定</el-button>
          </span>
      </el-dialog>
      </div>
    </template>
    <script>
    // 引入定义接口js文件
    import api from "@/api/project/fbgc/ljgc/ljtsfysd.js";
    import { getSystemErrorMap } from "util";
    import FileSaver from "file-saver"
    import { Loading } from 'element-ui';
    import Cookies from 'js-cookie'
    export default {
      
      data() {
        // 初始值
        let value = this.$route.query.projecttitle
        return {
          data: {
            proname: value,
          },
          flag:true,
          descriptions: [],
          dialogImportVisible: false,
          list: [], // 项目列表
          total: 0, // 总记录数
          page: 1, // 当前页码
          limit: 10, // 每页记录数
          searchObj: {}, // 查询条件
          multipleSelection: [], // 批量删除选中的记录列表
          dialogVisible: false,
          dialogVisible1:false,
          dialogVisible2:false,
          proname:this.$route.query.projecttitle,
          htd:this.$route.query.htdname,
          fbgcName:this.$route.query.fbgcName,
          username:Cookies.get('username'),
          userid:Cookies.get('userid'),
          type:Cookies.get('type'),
          hdgqd: {},
          file:'',// 待上传文件
          fileList:[],
          modifyData:{},
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；\n2，相同桩号会有两个不同的盒号，此时分两条数据填写；\n3，填写模板如下：",
          rules1:{
            jcsj:[
              {
                required:true,
                message:"检测时间不能为空！"
              }
            ],
            htgdz:[
              {
                required:true,
                message:"灰土规定值不能为空！"
              }
            ],
            zh:[
              {
                required:true,
                message:"桩号不能为空！"
              }
            ],
            jglx:[
              {
                required:true,
                message:"结构类型不能为空！"
              }
            ],
            jgcc:[
              {
                required:true,
                message:"结构层次不能为空！"
              }
            ],
            zdgmd:[
              {
                required:true,
                message:"最大干密度不能为空！"
              }
            ],
            bzsmd:[
              {
                required:true,
                message:"标准砂密度不能为空！"
              }
            ],
            qyzhjwz:[
              {
                required:true,
                message:"取样装号及位置不能为空！"
              }
            ],
            sksd:[
              {
                required:true,
                message:"试坑深度不能为空！"
              }
            ],
            ztjjbhbmjszl:[
              {
                required:true,
                message:"锥体及基板和表面间砂质量不能为空！"
              }
            ],
            gsqtszl:[
              {
                required:true,
                message:"灌砂前筒+砂质量(g)不能为空！"
              }
            ],
            gshtszl:[
              {
                required:true,
                message:"灌砂后筒+砂质量(g)试坑深度不能为空！"
              }
            ],
            syzl:[
              {
                required:true,
                message:"试样质量不能为空！"
              }
            ],
            hh:[
              {
                required:true,
                message:"盒号不能为空！"
              }
            ],
            hzl:[
              {
                required:true,
                message:"盒质量不能为空！"
              }
            ],
            hsshzl:[
              {
                required:true,
                message:"盒+湿试样质量(g)不能为空！"
              }
            ],
            hgsyzl:[
              {
                required:true,
                message:"盒+干试样质量(g)不能为空！"
              }
            ],
            xh:[
              {
                required:true,
                message:"序号不能为空！"
              }
            ],
            
            
            createtime:[
              {
                required:true,
                message:"创建时间不能为空！"
              }
            ]

          },
          rules2:{
            jcsj:[
              {
                required:true,
                message:"检测时间不能为空！"
              }
            ],
            slgdz:[
              {
                required:true,
                message:"沙砾规定值不能为空！"
              }
            ],
            zh:[
              {
                required:true,
                message:"桩号不能为空！"
              }
            ],
            a:[
              {
                required:true,
                message:"a不能为空！"
              }
            ],
            b:[
              {
                required:true,
                message:"b不能为空！"
              }
            ],
            c:[
              {
                required:true,
                message:"c不能为空！"
              }
            ],
            
            qyzhjwz:[
              {
                required:true,
                message:"取样装号及位置不能为空！"
              }
            ],
            sksd:[
              {
                required:true,
                message:"试坑深度不能为空！"
              }
            ],
            
            gsqtszl:[
              {
                required:true,
                message:"灌砂前筒+砂质量(g)不能为空！"
              }
            ],
            gshtszl:[
              {
                required:true,
                message:"灌砂后筒+砂质量(g)试坑深度不能为空！"
              }
            ],
            ztsz:[
              {
                required:true,
                message:"锥体砂重不能为空！"
              }
            ],
            lsdmd:[
              {
                required:true,
                message:"量砂的密度不能为空！"
              }
            ],
            hhldszl:[
              {
                required:true,
                message:"混合料的湿质量不能为空！"
              }
            ],
           hh:[
              {
                required:true,
                message:"盒号不能为空！"
              }
            ],
            hgzl:[
              {
                required:true,
                message:"盒+干质量不能为空！"
              }
            ],
            hzl:[
              {
                required:true,
                message:"盒质量不能为空！"
              }
            ],
            
            klzl:[
              {
                required:true,
                message:"5-38mm颗粒质量不能为空！"
              }
            ],
            xh:[
              {
                required:true,
                message:"序号不能为空！"
              }
            ],
            
            
            createtime:[
              {
                required:true,
                message:"创建时间不能为空！"
              }
            ]

          },
        };
      },
      created() {
        // 页面渲染之前
        //console.log(this.$route.query.projecttitle);
        //alert(this.$route.matched[1].meta.title)
        this.fetchData();
      },
      // 'http://localhost:8800/jjg/fbgc/ljgc/hdgqd/importhdgqd'
      methods: {
        handleClose(){
    
        },
        scjdb(){
          let commonInfoVo={}
          commonInfoVo.proname=this.proname
          commonInfoVo.htd=this.htd
          commonInfoVo.fbgc='路基土石方'
          commonInfoVo.userid=this.userid
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.scjdb(commonInfoVo).then((response) => {
                  return response.data.text()       
                })
                .then(res=>{
                  
                  let code = JSON.parse(res).code
                  if(code == 200){
                    api.download(commonInfoVo.proname,commonInfoVo.htd).then((response)=>{
                    
                    var temp = response.headers["content-disposition"]
                    var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    var matches = filenameRegex.exec(temp);
                    var filename=''
                    if (matches != null && matches[1]) { 
                      filename = matches[1].replace(/['"]/g, '');
                    }
                    console.log('鉴定表',response)
                    FileSaver.saveAs(response.data,decodeURI(filename))
                    load.close()
                  }).catch((err)=>{
                        
                        load.close()
                      })
                  
                  }
                  else{
                    load.close()
                    let s=JSON.parse(res)
                    
                    this.$message.warning(s.message)
                    
                  }
                })
                .catch((err)=>{
                        console.log(err)
                        load.close()
                      });
          
    
        },
        ckjdb(){
            this.dialogVisible1=true
            let commonInfoVo={}
            commonInfoVo.proname=this.proname
            commonInfoVo.htd=this.htd
            commonInfoVo.fbgc='路基土石方'
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('sdsds',res)
              this.descriptions=res.data
            })
        },
        importysdht() {
          this.dialogImportVisible = true;
        },
        importysdsl() {
          this.dialogImportVisible = true;
        },
        // 上传文件触发
        fileChange(file,fileList){
          this.fileList=fileList
          
        },
        onUploadSuccess() {
          
          //debugger
          let fd=new FormData()
          
          
    
          fd.append('file',this.fileList[0].raw)
          fd.append('proname',this.proname)
          fd.append('htd',this.htd)
          fd.append('fbgc','路基土石方')
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          if(this.flag==true){
            api.importysdht(fd).then((res)=>{
                console.log('res',res)
                if(res.message=='成功'){
                this.fetchData();
                this.$message({
                    type: "success",
                    message: "上传成功!",
                });
                this.dialogImportVisible = false;
                this.fileList=[]
                }
                else{
                this.$alert('上传失败！')
                }
                load.close()
            }).catch((err)=>{
              
              load.close()
            })
          }
          else{
            api.importysdsl(fd).then((res)=>{
                console.log('res',res)
                if(res.message=='成功'){
                this.fetchData();
                this.$message({
                    type: "success",
                    message: "上传成功!",
                });
                this.dialogImportVisible = false;
                this.fileList=[]
                }
                else{
                this.$alert('上传失败！')
                }
                load.close()
            }).catch((err)=>{
              
              load.close()
            })
          }
          
          
          
        },
        exportysdht() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          
          api.exportysdht().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `01路基土石方压实度_灰土实测数据` + ".xlsx";
            link.style.display = "none";
            link.href = objectUrl;
            link.click();
            document.body.appendChild(link);
          });
        },
        exportysdsl() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          
          api.exportysdsl().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `01路基土石方压实度_砂砾实测数据` + ".xlsx";
            link.style.display = "none";
            link.href = objectUrl;
            link.click();
            document.body.appendChild(link);
          });
        },
        modify(){
          if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要修改的记录！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能修改一条记录！");
            return;
          }
          this.modifyData={}
          if(this.flag==true){
            api.getById(this.multipleSelection[0].id).then((res)=>{
              if(res.message=='成功'){
                this.dialogVisible2=true
                this.modifyData=res.data
                console.log('res',this.modifyData)
              }
              else{
                this.$alert('出错')
              }
            })
          }
          else{
            api.getById1(this.multipleSelection[0].id).then((res)=>{
              if(res.message=='成功'){
                this.dialogVisible2=true
                this.modifyData=res.data
              }
              else{
                this.$alert('出错')
              }
            })
          }
          
        },
        update(){
          if(this.flag==true){
            this.$refs["dataForm1"].validate((valid)=>{
              if(valid){
                api.modify(this.modifyData).then((res)=>{
                  if(res.message=='成功'){
                    this.$message({
                      type: "success",
                      message: "修改成功!",
                    });
                    // 刷新页面
                    this.dialogVisible2=false
                    
                    this.fetchData();
                  }
                  else{
                    this.$alert('出错')
                  }
                })
              }
            })
            
          }
          else{
            this.$refs["dataForm2"].validate((valid)=>{
              if(valid){
                api.modify1(this.modifyData).then((res)=>{
                  if(res.message=='成功'){
                    this.$message({
                      type: "success",
                      message: "修改成功!",
                    });
                    // 刷新页面
                    this.dialogVisible2=false
                    
                    this.fetchData();
                  }
                  else{
                    this.$alert('出错')
                  }
                })
              }
            })
            
          }
          
        },
        // 批量删除
        batchRemove() {
          // 判断非空
          if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要删除的记录！");
            return;
          }
          this.$confirm("此操作将删除该项目的所有信息, 是否继续?", "提示", {
            confirmButtonText: "确定",
            cancelButtonText: "取消",
            type: "warning",
          }).then(() => {
            var idList = []; // [1,2,3 ]
            // 遍历数组
            for (var i = 0; i < this.multipleSelection.length; i++) {
              var obj = this.multipleSelection[i];
              var id = obj.id;
              idList.push(id);
            }
            if(this.flag==true){
                // 调用接口删除
                api.batchRemove(idList).then((response) => {
                // 提示信息
                this.$message({
                    type: "success",
                    message: "删除成功!",
                });
                // 刷新页面
                this.fetchData();
                });
            }
            else{
                // 调用接口删除
                api.batchRemove1(idList).then((response) => {
                // 提示信息
                this.$message({
                    type: "success",
                    message: "删除成功!",
                });
                // 刷新页面
                this.fetchData();
                });
            }
          });
        },
         //全部删除
        removeAll() {
          if(this.flag==true){
            this.$confirm("此操作将删除该项目的所有信息, 是否继续?", "提示", {
            confirmButtonText: "确定",
            cancelButtonText: "取消",
            type: "warning",
          }).then(() => {
            let commonInfoVo={}
            commonInfoVo.proname=this.proname
            commonInfoVo.htd=this.htd
            commonInfoVo.fbgc='路基土石方'
            commonInfoVo.userid=this.userid
            // 调用接口删除
            api.removeAllHt(commonInfoVo).then((response) => {
              // 提示信息
              this.$message({
                type: "success",
                message: "删除成功!",
              });
              // 刷新页面
              this.fetchData();
            });
          });

          }else{
            this.$confirm("此操作将删除该项目的所有信息, 是否继续?", "提示", {
            confirmButtonText: "确定",
            cancelButtonText: "取消",
            type: "warning",
          }).then(() => {
            let commonInfoVo={}
            commonInfoVo.proname=this.proname
            commonInfoVo.htd=this.htd
            commonInfoVo.fbgc='路基土石方'
            commonInfoVo.userid=this.userid
            // 调用接口删除
            api.removeAllSl(commonInfoVo).then((response) => {
              // 提示信息
              this.$message({
                type: "success",
                message: "删除成功!",
              });
              // 刷新页面
              this.fetchData();
            });
          });

          }
          
          
        },
        // 复选框发生变化，调用方法，选中复选框行内容传递
        handleSelectionChange(selection) {
          this.multipleSelection = selection;
        },
        fetchData() {
          this.searchObj.proname = this.proname;
          this.searchObj.htd=this.htd
          this.searchObj.fbgc='路基土石方'
          this.searchObj.userid=this.userid
          console.log('eeee',this.htd)
          // 调用api
          if(this.flag==true){
            api.pageList(this.page, this.limit, this.searchObj).then((response) => {
                this.list = response.data.records;
                this.total = response.data.total;
                console.log('ssss',this.list)
            });
          }
          else{
            api.pageList1(this.page, this.limit, this.searchObj).then((response) => {
                this.list = response.data.records;
                this.total = response.data.total;
                console.log('ssss',this.list)
            });
          }
          
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
        // 清空
        resetData() {
          this.searchObj = {};
          this.fetchData();
        },
      },
    };
    </script>
    <style lang="scss" scoped>
    .tools-div{
      .anniu{
        .shuoming{
          float: right;
        }
      }
    }
    .app-container{
      height: 100%;
      overflow: hidden;
    }
    </style>