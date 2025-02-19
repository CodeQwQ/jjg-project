<template>
    <div class="app-container">
        <!--查询表单-->
        <div class="search-div">
          <el-form :inline="true" class="demo-form-inline">
            <el-form-item label="检测时间:">
              <!-- <el-input v-model="searchObj.jcsj" /> -->
              <el-date-picker
                v-model="searchObj.jcsj"
                type="date"
                value-format="yyyy-MM-dd"
              />
            </el-form-item>
            <el-form-item label="位置:">
              <el-input v-model="searchObj.wz" />
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
        <div class="tools-div" v-show="!(username=='admin'&&type==1  )">
          <div class="anniu">
            <el-button
              class="btn-add"
              type="success"
              @click="exportjabz()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出交安标志文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importjabz()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入交安标志数据文件</el-button
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
          <el-table-column prop="jcsj" label="检测时间" width="200px" />
          
          <el-table-column prop="wz" label="位置" width="160px" />
          <el-table-column prop="lzlx" label="立柱类型" width="100px" />
          <el-table-column prop="szdyxps" label="竖直度允许偏差(mm/m)" width="120px"/>
          
          
          
          <el-table-column prop="fx1scz1" label="方向1实测值1(mm/m)" width="110px"/>
          <el-table-column prop="fx1scz2" label="方向1实测值2(mm/m)" width="110px"/>
          <el-table-column prop="fx2scz1" label="方向2实测值1(mm/m)" width="110px"/>
          <el-table-column prop="fx2scz2" label="方向2实测值2(mm/m)" width="110px"/>
          <el-table-column prop="jkgdz" label="净空规定值(mm)" width="100px"/>
          <el-table-column prop="jkscz" label="净空实测值(mm)" width="100px"/>

          <el-table-column prop="hdyxps" label="厚度允许偏差(mm)" width="120px"/>
          <el-table-column prop="hdcz1" label="厚度测值1" width="100px"/>
          <el-table-column prop="hdcz2" label="厚度测值2" width="100px"/>
          <el-table-column prop="bsvlyxps" label="白色V类允许偏差(cd*lx-1*m-2)" width="130px"/>
          <el-table-column prop="bsvlscz1" label="白色V类实测值1" width="140px"/>
          <el-table-column prop="bsvlscz2" label="白色V类实测值2" width="140px"/>
          <el-table-column prop="bswlyxps" label="白色IV类允许偏差" width="140px"/>
          <el-table-column prop="bswlscz1" label="白色IV类实测值1" width="140px"/>
          <el-table-column prop="bswlscz2" label="白色IV类实测值2" width="140px"/>
          <el-table-column prop="lsvlyxps" label="绿色V类允许偏差(cd*lx-1*m-2)" width="130px"/>
          <el-table-column prop="lsvlscz1" label="绿色V类实测值1" width="140px"/>
          <el-table-column prop="lsvlscz2" label="绿色V类实测值2" width="140px"/>
          <el-table-column prop="lswlyxps" label="绿色IV类允许偏差" width="140px"/>
          <el-table-column prop="lswlscz1" label="绿色IV类实测值1" width="140px"/>
          <el-table-column prop="lswlscz2" label="绿色IV类实测值2" width="140px"/>
          <el-table-column prop="hsvlyxps" label="黄色V类允许偏差(cd*lx-1*m-2)" width="130px"/>
          <el-table-column prop="hsvlscz1" label="黄色V类实测值1" width="140px"/>
          <el-table-column prop="hsvlscz2" label="黄色V类实测值2" width="140px"/>
          <el-table-column prop="hswlyxps" label="黄色IV类允许偏差" width="140px"/>
          <el-table-column prop="hswlscz1" label="黄色IV类实测值1" width="140px"/>
          <el-table-column prop="hswlscz2" label="黄色IV类实测值2" width="140px"/>
          <el-table-column prop="lasvlyxps" label="蓝色V类允许偏差(cd*lx-1*m-2)" width="130px"/>
          <el-table-column prop="lasvlscz1" label="蓝色V类实测值1" width="140px"/>
          <el-table-column prop="lasvlscz2" label="蓝色V类实测值2" width="140px"/>
          <el-table-column prop="laswlyxps" label="蓝色IV类允许偏差" width="140px"/>
          <el-table-column prop="laswlscz1" label="蓝色IV类实测值1" width="140px"/>
          <el-table-column prop="laswlscz2" label="蓝色IV类实测值2" width="140px"/>
          <el-table-column prop="rsvlyxps" label="红色V类允许偏差(cd*lx-1*m-2)" width="130px"/>
          <el-table-column prop="rsvlscz1" label="红色V类实测值1" width="140px"/>
          <el-table-column prop="rsvlscz2" label="红色V类实测值2" width="140px"/>
          <el-table-column prop="rswlyxps" label="红色IV类允许偏差" width="140px"/>
          <el-table-column prop="rswlscz1" label="红色IV类实测值1" width="140px"/>
          <el-table-column prop="rswlscz2" label="红色IV类实测值2" width="140px"/>

          <el-table-column prop="createtime" label="创建时间" width="200px"/>
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
      
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/交安/交安标志.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="交安标志检测结果"
         :visible.sync="dialogVisible1">
         <el-table
          :data="jdblist"
          border
          stripe
          style="width: 100%;"
          class="table"
          :header-cell-style="{ 'text-align': 'center' }"
          :cell-style="{ 'text-align': 'center' }"
          
        >
        <el-table-column prop="项目" label="项目" width="180px"/>
         <el-table-column prop="ys" label="颜色" />
         <el-table-column prop="规定值或允许偏差" label="规定值或允许偏差" />
         <el-table-column prop="总点数" label="总点数" />
         <el-table-column prop="合格点数" label="合格点数" />
         <el-table-column prop="合格率" label="合格率(%)" />
         
         
         
        </el-table>
         <span slot="footer" class="dialog-footer">
          <el-button type="primary"  @click="dialogVisible1 = false">关闭</el-button>
         </span>
         
        </el-dialog>
        <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm" :model="modifyData"  label-width="160px" size="small" style="padding-right: 40px;" :rules="rules">
            <el-form-item label="检测时间" prop="jcsj">
            <el-input v-model="modifyData.jcsj"></el-input>
            </el-form-item>
            <el-form-item label="位置" prop="wz">
            <el-input v-model="modifyData.wz"></el-input>
            </el-form-item>
            <el-form-item label="立柱类型" prop="lzlx">
            <el-input v-model="modifyData.lzlx"></el-input>
            </el-form-item>
            
            <el-form-item label="竖直度允许偏差" prop="szdyxps">
            <el-input v-model="modifyData.szdyxps"></el-input>
            </el-form-item>
            <el-form-item label="方向1实测值1" prop="fx1scz1">
            <el-input v-model="modifyData.fx1scz1"></el-input>
            </el-form-item>
            <el-form-item label="方向1实测值2" prop="fx1scz2">
            <el-input v-model="modifyData.fx1scz2"></el-input>
            </el-form-item>
            <el-form-item label="方向2实测值1" prop="fx2scz1">
            <el-input v-model="modifyData.fx2scz1"></el-input>
            </el-form-item>
            <el-form-item label="方向2实测值2" prop="fx2scz2">
            <el-input v-model="modifyData.fx2scz2"></el-input>
            </el-form-item>
            <el-form-item label="净空规定值" prop="jkgdz">
            <el-input v-model="modifyData.jkgdz"></el-input>
            </el-form-item>
            <el-form-item label="净空实测值" prop="jkscz">
            <el-input v-model="modifyData.jkscz"></el-input>
            </el-form-item>
            <el-form-item label="厚度允许偏差" prop="hdyxps">
            <el-input v-model="modifyData.hdyxps"></el-input>
            </el-form-item>
            <el-form-item label="厚度测值1" prop="hdcz1">
            <el-input v-model="modifyData.hdcz1"></el-input>
            </el-form-item>
            <el-form-item label="厚度测值2" prop="hdcz2">
            <el-input v-model="modifyData.hdcz2"></el-input>
            </el-form-item>
            <el-form-item label="白色V类允许偏差" prop="bsvlyxps">
            <el-input v-model="modifyData.bsvlyxps"></el-input>
            </el-form-item>
            <el-form-item label="白色V类实测值1" prop="bsvlscz1">
            <el-input v-model="modifyData.bsvlscz1"></el-input>
            </el-form-item>
            <el-form-item label="白色V类实测值2" prop="bsvlscz2">
            <el-input v-model="modifyData.bsvlscz2"></el-input>
            </el-form-item>
            <el-form-item label="白色IV类允许偏差" prop="bswlyxps">
            <el-input v-model="modifyData.bswlyxps"></el-input>
            </el-form-item>
            <el-form-item label="白色IV类实测值1" prop="bswlscz1">
            <el-input v-model="modifyData.bswlscz1"></el-input>
            </el-form-item>
            <el-form-item label="白色IV类实测值2" prop="bswlscz2">
            <el-input v-model="modifyData.bswlscz2"></el-input>
            </el-form-item>
            
            <el-form-item label="绿色V类允许偏差" prop="lsvlyxps">
            <el-input v-model="modifyData.lsvlyxps"></el-input>
            </el-form-item>
            <el-form-item label="绿色V类实测值1" prop="lsvlscz1">
            <el-input v-model="modifyData.lsvlscz1"></el-input>
            </el-form-item>
            <el-form-item label="绿色V类实测值2" prop="lsvlscz2">
            <el-input v-model="modifyData.lsvlscz2"></el-input>
            </el-form-item>
            <el-form-item label="绿色IV类允许偏差" prop="lswlyxps">
            <el-input v-model="modifyData.lswlyxps"></el-input>
            </el-form-item>
            <el-form-item label="绿色IV类实测值1" prop="lswlscz1">
            <el-input v-model="modifyData.lswlscz1"></el-input>
            </el-form-item>
            <el-form-item label="绿色IV类实测值2" prop="lswlscz2">
            <el-input v-model="modifyData.lswlscz2"></el-input>
            </el-form-item>
            
            <el-form-item label="黄色V类允许偏差" prop="hsvlyxps">
            <el-input v-model="modifyData.hsvlyxps"></el-input>
            </el-form-item>
            <el-form-item label="黄色V类实测值1" prop="hsvlscz1">
            <el-input v-model="modifyData.hsvlscz1"></el-input>
            </el-form-item>
            <el-form-item label="黄色V类实测值2" prop="hsvlscz2">
            <el-input v-model="modifyData.hsvlscz2"></el-input>
            </el-form-item>
            <el-form-item label="黄色IV类允许偏差" prop="hswlyxps">
            <el-input v-model="modifyData.hswlyxps"></el-input>
            </el-form-item>
            <el-form-item label="黄色IV类实测值1" prop="hswlscz1">
            <el-input v-model="modifyData.hswlscz1"></el-input>
            </el-form-item>
            <el-form-item label="黄色IV类实测值2" prop="hswlscz2">
            <el-input v-model="modifyData.hswlscz2"></el-input>
            </el-form-item>

            <el-form-item label="蓝色V类允许偏差" prop="lasvlyxps">
            <el-input v-model="modifyData.lasvlyxps"></el-input>
            </el-form-item>
            <el-form-item label="蓝色V类实测值1" prop="lasvlscz1">
            <el-input v-model="modifyData.lasvlscz1"></el-input>
            </el-form-item>
            <el-form-item label="蓝色V类实测值2" prop="lasvlscz2">
            <el-input v-model="modifyData.lasvlscz2"></el-input>
            </el-form-item>
            <el-form-item label="蓝色IV类允许偏差" prop="laswlyxps">
            <el-input v-model="modifyData.laswlyxps"></el-input>
            </el-form-item>
            <el-form-item label="蓝色IV类实测值1" prop="laswlscz1">
            <el-input v-model="modifyData.laswlscz1"></el-input>
            </el-form-item>
            <el-form-item label="蓝色IV类实测值2" prop="laswlscz2">
            <el-input v-model="modifyData.laswlscz2"></el-input>
            </el-form-item>

            <el-form-item label="红色V类允许偏差" prop="rsvlyxps">
            <el-input v-model="modifyData.rsvlyxps"></el-input>
            </el-form-item>
            <el-form-item label="红色V类实测值1" prop="rsvlscz1">
            <el-input v-model="modifyData.rsvlscz1"></el-input>
            </el-form-item>
            <el-form-item label="红色V类实测值2" prop="rsvlscz2">
            <el-input v-model="modifyData.rsvlscz2"></el-input>
            </el-form-item>
            <el-form-item label="红色IV类允许偏差" prop="rswlyxps">
            <el-input v-model="modifyData.rswlyxps"></el-input>
            </el-form-item>
            <el-form-item label="红色IV类实测值1" prop="rswlscz1">
            <el-input v-model="modifyData.rswlscz1"></el-input>
            </el-form-item>
            <el-form-item label="红色IV类实测值2" prop="rswlscz2">
            <el-input v-model="modifyData.rswlscz2"></el-input>
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
    import api from "@/api/project/fbgc/jagc/jabz.js";
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
          descriptions: {
             '不合格点数':0,
             '合格点数':0,
             '总点数':0,
             '合格率':0
            },
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
          jdblist:[],
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；\n2，填写模板如下：",
          rules:{
            jcsj:[
              {
                required:true,
                message:"检测时间不能为空！"
              }
            ],
            wz:[
              {
                required:true,
                message:"位置不能为空！"
              }
            ],
            lzlx:[
              {
                required:true,
                message:"立柱类型不能为空！"
              }
            ],
            szdyxps:[
              {
                required:true,
                message:"竖直度允许偏差不能为空！"
              }
            ],
            jkgdz:[
              {
                required:true,
                message:"净空规定值不能为空！"
              }
            ],
            hdyxps:[
              {
                required:true,
                message:"厚度允许偏差不能为空！"
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
          commonInfoVo.fbgc='交通安全设施'
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
            commonInfoVo.fbgc='交通安全设施'
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('tttt',res)
              this.jdblist=res.data
            })
        },
        importjabz() {
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
          fd.append('fbgc','交通安全设施')
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importjabz(fd).then((res)=>{
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
          
          
          
        },
        exportjabz() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          debugger
          api.exportjabz().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `01交安标志实测数据` + ".xlsx";
            link.style.display = "none";
            link.href = objectUrl;
            link.click();
            document.body.appendChild(link);
          });
        },
         // 修改
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
          api.getById(this.multipleSelection[0].id).then((res)=>{
            if(res.message=='成功'){
              this.dialogVisible2=true
              this.modifyData=res.data
            }
            else{
              this.$alert('出错')
            }
          })
          
        },
        update(){
          this.$refs["dataForm"].validate((valid)=>{
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
          });
        },
         //全部删除
        removeAll() {
          
          this.$confirm("此操作将删除该项目的所有信息, 是否继续?", "提示", {
            confirmButtonText: "确定",
            cancelButtonText: "取消",
            type: "warning",
          }).then(() => {
            let commonInfoVo={}
            commonInfoVo.proname=this.proname
            commonInfoVo.htd=this.htd
            commonInfoVo.fbgc='交通安全设施'
            commonInfoVo.userid=this.userid
            // 调用接口删除
            api.removeAll(commonInfoVo).then((response) => {
              // 提示信息
              this.$message({
                type: "success",
                message: "删除成功!",
              });
              // 刷新页面
              this.fetchData();
            });
          });
        },
        // 复选框发生变化，调用方法，选中复选框行内容传递
        handleSelectionChange(selection) {
          this.multipleSelection = selection;
        },
        fetchData() {
          this.searchObj.proname = this.proname;
          this.searchObj.htd=this.htd
          this.searchObj.fbgc='交通安全设施'
          this.searchObj.userid=this.userid
          // 调用api
          api.pageList(this.page, this.limit, this.searchObj).then((response) => {
            this.list = response.data.records;
            this.total = response.data.total;
            console.log('eeee',this.list)
          });
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