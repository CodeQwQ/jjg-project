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
        <div class="tools-div" v-show="!(username=='admin'&&type==1)">
          <div class="anniu">
            <el-button
              class="btn-add"
              type="success"
              @click="exportjabx()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出交安标线文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importjabx()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入交安标线数据文件</el-button
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
            >修改
          </el-button>
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
          <el-table-column prop="jcsj" label="检测时间" width="100px" />
          <el-table-column prop="bxlx" label="标线类型" width="100px" />
          <el-table-column prop="wz" label="位置(ZK0+001)" width="130px" />
          <el-table-column prop="hdgdz" label="厚度规定值(mm)" width="130px"/>
          
          <el-table-column prop="hdyxpsz" label="厚度允许偏差+" width="120px"/>
          <el-table-column prop="hdyxpsf" label="厚度允许偏差-" width="120px"/>
          
          <el-table-column prop="hdscz1" label="厚度实测值1" width="100px"/>
          <el-table-column prop="hdscz2" label="厚度实测值2" width="100px"/>
          <el-table-column prop="hdscz3" label="厚度实测值3" width="100px"/>
          <el-table-column prop="hdscz4" label="厚度实测值4" width="100px"/>
          <el-table-column prop="hdscz5" label="厚度实测值5" width="100px"/>
          <el-table-column prop="bxnfsxsgdz" label="白线逆反射系数规定值(mcd*m-2*lx-1)" width="170px"/>
          <el-table-column prop="bxscz1" label="白线实测值1" width="100px"/>
          <el-table-column prop="bxscz2" label="白线实测值2" width="100px"/>
          <el-table-column prop="bxscz3" label="白线实测值3" width="100px"/>
          <el-table-column prop="bxscz4" label="白线实测值4" width="100px"/>
          <el-table-column prop="bxscz5" label="白线实测值5" width="100px"/>
          <el-table-column prop="hxnfsxsgdz" label="黄线逆反射系数规定值" width="170px"/>
          <el-table-column prop="hxscz1" label="黄线实测值1" width="100px"/>
          <el-table-column prop="hxscz2" label="黄线实测值2" width="100px"/>
          <el-table-column prop="hxscz3" label="黄线实测值3" width="100px"/>
          <el-table-column prop="hxscz4" label="黄线实测值4" width="100px"/>
          <el-table-column prop="hxscz5" label="黄线实测值5" width="100px"/>
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
      
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/交安/交安标线.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="交安标线检测结果"
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
          
          <el-table-column prop="检测项目" label="检测项目" width="250px"/>
          <el-table-column prop="规定值或允许偏差" label="规定值或允许偏差" />
          
          <el-table-column prop="总点数" label="总点数"  />
          <el-table-column prop="合格率" label="合格率(%)" />
          <el-table-column prop="合格点数" label="合格点数" />
          
          
          
          
         
          
          
        </el-table>
         <span slot="footer" class="dialog-footer">
          <el-button type="primary"  @click="dialogVisible1 = false">关闭</el-button>
         </span>
         
        </el-dialog>
        <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm" :model="modifyData"  label-width="180px" size="small" style="padding-right: 40px;" :rules="rules">
            <el-form-item label="检测时间" prop="jcsj">
            <el-input v-model="modifyData.jcsj"></el-input>
            </el-form-item>
            <el-form-item label="标线类型" prop="bxlx">
            <el-input v-model="modifyData.bxlx"></el-input>
            </el-form-item>
            <el-form-item label="位置(ZK0+001)" prop="wz">
            <el-input v-model="modifyData.wz"></el-input>
            </el-form-item>
            <el-form-item label="厚度规定值" prop="hdgdz">
            <el-input v-model="modifyData.hdgdz"></el-input>
            </el-form-item>
            <el-form-item label="厚度允许偏差+" prop="hdyxpsz">
            <el-input v-model="modifyData.hdyxpsz"></el-input>
            </el-form-item>
            <el-form-item label="厚度允许偏差-" prop="hdyxpsf">
            <el-input v-model="modifyData.hdyxpsf"></el-input>
            </el-form-item>
            <el-form-item label="厚度实测值1" prop="hdscz1">
            <el-input v-model="modifyData.hdscz1"></el-input>
            </el-form-item>
            <el-form-item label="厚度实测值2" prop="hdscz2">
            <el-input v-model="modifyData.hdscz2"></el-input>
            </el-form-item>
            <el-form-item label="厚度实测值3" prop="hdscz3">
            <el-input v-model="modifyData.hdscz3"></el-input>
            </el-form-item>
            <el-form-item label="厚度实测值4" prop="hdscz4">
            <el-input v-model="modifyData.hdscz4"></el-input>
            </el-form-item>
            <el-form-item label="白线逆反射系数规定值" prop="bxnfsxsgdz">
            <el-input v-model="modifyData.bxnfsxsgdz"></el-input>
            </el-form-item>
            <el-form-item label="白线实测值1" prop="bxscz1">
            <el-input v-model="modifyData.bxscz1"></el-input>
            </el-form-item>
            <el-form-item label="白线实测值2" prop="bxscz2">
            <el-input v-model="modifyData.bxscz2"></el-input>
            </el-form-item>
            <el-form-item label="白线实测值2" prop="bxscz2">
            <el-input v-model="modifyData.bxscz2"></el-input>
            </el-form-item>
            <el-form-item label="白线实测值3" prop="bxscz3">
            <el-input v-model="modifyData.bxscz3"></el-input>
            </el-form-item>
            <el-form-item label="白线实测值4" prop="bxscz4">
            <el-input v-model="modifyData.bxscz4"></el-input>
            </el-form-item>
            <el-form-item label="白线实测值5" prop="bxscz5">
            <el-input v-model="modifyData.bxscz5"></el-input>
            </el-form-item>
            <el-form-item label="黄线逆反射系数规定值" prop="hxnfsxsgdz">
            <el-input v-model="modifyData.hxnfsxsgdz"></el-input>
            </el-form-item>
            <el-form-item label="黄线实测值1" prop="hxscz1">
            <el-input v-model="modifyData.hxscz1"></el-input>
            </el-form-item>
            <el-form-item label="黄线实测值2" prop="hxscz2">
            <el-input v-model="modifyData.hxscz2"></el-input>
            </el-form-item>
            <el-form-item label="黄线实测值3" prop="hxscz3">
            <el-input v-model="modifyData.hxscz3"></el-input>
            </el-form-item>
            <el-form-item label="黄线实测值4" prop="hxscz4">
            <el-input v-model="modifyData.hxscz4"></el-input>
            </el-form-item>
            <el-form-item label="黄线实测值5" prop="hxscz5">
            <el-input v-model="modifyData.hxscz5"></el-input>
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
    import api from "@/api/project/fbgc/jagc/jabx.js";
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
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；\n2，请确保每个桩号有5条数据，若不足5条，实测值可以不填写，但是其他数据则按照相同桩号的数据填写；\n3，当某条数据无逆反射系数的实测值时，不需要填写逆反射系数规定值；\n4，填写模板如下：",
          rules:{
            jcsj:[
              {
                required:true,
                message:"检测时间不能为空！"
              }
            ],
            bxlx:[
              {
                required:true,
                message:"标线类型不能为空！"
              }
            ],
            wz:[
              {
                required:true,
                message:"位置不能为空！"
              }
            ],
            hdgdz:[
              {
                required:true,
                message:"厚度规定值不能为空！"
              }
            ],
            hdyxpsz:[
              {
                required:true,
                message:"厚度允许偏差+不能为空！"
              }
            ],
            hdyxpsf:[
              {
                required:true,
                message:"厚度允许偏差-不能为空！"
              }
            ],
            bxnfsxsgdz:[
              {
                required:true,
                message:"白线逆反射系数规定值不能为空！"
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
          commonInfoVo.fbgc='标线'
          commonInfoVo.userid=this.userid
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})


          api.scjdb(commonInfoVo).then((response) => {
                  return response.data.text()       
                })
                .then(res=>{
                  
                  let code = JSON.parse(res).code
                  if(code == 200){
                    api.download(commonInfoVo.proname,commonInfoVo.htd,commonInfoVo.fbgc).then((response)=>{
                    
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
            commonInfoVo.fbgc='标线'
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('tttt',res)
              this.descriptions=res.data
            })
        },
        importjabx() {
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
          fd.append('fbgc','标线')
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importjabx(fd).then((res)=>{
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
        exportjabx() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          api.exportjabx().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `02交安标线实测数据` + ".xlsx";
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
            commonInfoVo.fbgc='边线'
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
          this.searchObj.fbgc='标线'
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