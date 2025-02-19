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
        <div class="tools-div" v-show="!(username=='admin'&&type==1  )">
          <div class="anniu">
            <el-button
              class="btn-add"
              type="success"
              @click="exporttlmxlbgc()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出砼路面相邻板高差文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importtlmxlbgc()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入砼路面相邻板高差数据文件</el-button
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
            <el-table-column prop="jcsj" label="检测时间" width="100px" />
            <el-table-column prop="qywzmc" label="取样位置名称" width="220px" />
            <el-table-column prop="zh" label="桩号" />
            <el-table-column prop="scz1" label="实测值1(mm)" />
            <el-table-column prop="scz2" label="实测值2(mm)" />
            <el-table-column prop="scz3" label="实测值3(mm)" />
            <el-table-column prop="scz4" label="实测值4(mm)" />
            <el-table-column prop="bgcgdz" label="板高差规定值" />
            <el-table-column prop="createtime" label="创建时间" width="100px" />
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
      
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/路面/砼路面相邻板高差.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="砼路面相邻板高差检测结果"
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
          
        <el-table-column prop="总点数" label="总点数"  />
          <el-table-column prop="合格点数" label="合格点数" />
          
          <el-table-column prop="合格率" label="合格率(%)" />
          <el-table-column prop="规定值" label="规定值"  />
          <el-table-column prop="max" label="最大值"  />
          <el-table-column prop="min" label="最小值"  />
          
          
          
         
          
          
        </el-table>
         <span slot="footer" class="dialog-footer">
          <el-button type="primary"  @click="dialogVisible1 = false">关闭</el-button>
         </span>
         
        </el-dialog>
        <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm" :model="modifyData"  label-width="130px" size="small" style="padding-right: 40px;" :rules="rules">
            <el-form-item label="检测时间" prop="jcsj">
            <el-input v-model="modifyData.jcsj"></el-input>
            </el-form-item>
            <el-form-item label="取样位置名称" prop="qywzmc">
            <el-input v-model="modifyData.qywzmc"></el-input>
            </el-form-item>
            <el-form-item label="桩号" prop="zh">
            <el-input v-model="modifyData.zh"></el-input>
            </el-form-item>
            <el-form-item label="实测值1" prop="scz1">
            <el-input v-model="modifyData.scz1"></el-input>
            </el-form-item>
            <el-form-item label="实测值2" prop="scz2">
            <el-input v-model="modifyData.scz2"></el-input>
            </el-form-item>
            <el-form-item label="实测值3" prop="scz3">
            <el-input v-model="modifyData.scz3"></el-input>
            </el-form-item>
            <el-form-item label="实测值4" prop="scz4">
            <el-input v-model="modifyData.scz4"></el-input>
            </el-form-item>
            <el-form-item label="板高差规定值" prop="bgcgdz">
            <el-input v-model="modifyData.bgcgdz"></el-input>
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
    import api from "@/api/project/fbgc/lmgc/tlmxlbgc";
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
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；\n2，填写模板如下：",
          rules:{
            jcsj:[
              {
                required:true,
                message:"检测时间不能为空！"
              }
            ],
            qywzmc:[
              {
                required:true,
                message:"取样位置名称不能为空！"
              }
            ],
            zh:[
              {
                required:true,
                message:"桩号不能为空！"
              }
            ],
            
            
            
            scz1:[
              {
                required:true,
                message:"实测值1不能为空！"
              }
            ],
            scz2:[
              {
                required:true,
                message:"实测值2不能为空！"
              }
            ],
            scz3:[
              {
                required:true,
                message:"实测值3不能为空！"
              }
            ],
            bgcgdz:[
              {
                required:true,
                message:"板高差规定值不能为空！"
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
          commonInfoVo.fbgc='路面工程'
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
            commonInfoVo.fbgc='路面工程'
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('tttt',res)
              this.descriptions=res.data
            })
        },
        importtlmxlbgc() {
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
          fd.append('fbgc','路面工程')
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importtlmxlbgc(fd).then((res)=>{
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
        exporttlmxlbgc() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          debugger
          api.exporttlmxlbgc().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `06砼路面相邻板高差实测数据` + ".xlsx";
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
            commonInfoVo.fbgc='路面工程'
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
          this.searchObj.fbgc='路面工程'
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