<template>
    <div class="app-container">
        <!--查询表单-->
        <div class="search-div">
          <el-form :inline="true" class="demo-form-inline">
            
            <el-form-item label="起点桩号:">
              <el-input v-model="searchObj.qdzh" />
            </el-form-item>
            <el-form-item label="终点桩号:">
              <el-input v-model="searchObj.zdzh" />
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
              @click="dialogVisible3=true"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出构造深度文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importgzsd()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入构造深度数据文件</el-button
            >
            <el-button
              class="btn-add"
              type="success"
              size="mini"
              @click="dialogVisible4=true"
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
            
            <el-table-column prop="qdzh" label="起点桩号" width="220px" />
            <el-table-column prop="zdzh" label="终点桩号" />
            <el-table-column prop="mtd" label="代表MTD" />
            <el-table-column prop="lxbs" label="路线类型" />
            <el-table-column prop="zdbs" label="匝道标识" />
            <el-table-column prop="cd" label="车道" />
            
            
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
          <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/路面/构造深度.png" alt="" >
          <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="构造深度检测结果"
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
          
          <el-table-column prop="检测项目" label="检测项目" />
          <el-table-column prop="路面类型" label="路面类型" />
          <el-table-column prop="总点数" label="总点数" width="100px" />
          <el-table-column prop="合格点数" label="合格点数" />
          <el-table-column prop="合格率" label="合格率(%)" />
          <el-table-column prop="设计值" label="设计值" />
          <el-table-column prop="最大值" label="最大值" />
          <el-table-column prop="最小值" label="最小值" />
          
          
          
         
          
          
        </el-table>
         <span slot="footer" class="dialog-footer">
          <el-button type="primary"  @click="dialogVisible1 = false">关闭</el-button>
         </span>
         
        </el-dialog>
        <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm" :model="modifyData"  label-width="120px" size="small" style="padding-right: 40px;" :rules="rules1">
            
            
            <el-form-item label="起点桩号" prop="qdzh">
            <el-input v-model="modifyData.qdzh"></el-input>
            </el-form-item>
            <el-form-item label="终点桩号" prop="zdzh">
            <el-input v-model="modifyData.zdzh"></el-input>
            </el-form-item>
            <el-form-item label="代表MTD" prop="mtd">
            <el-input v-model="modifyData.mtd"></el-input>
            </el-form-item>
            <el-form-item label="路线类型" prop="lxbs">
            <el-input v-model="modifyData.lxbs"></el-input>
            </el-form-item>
            <el-form-item label="匝道标识" prop="zdbs">
            <el-input v-model="modifyData.zdbs"></el-input>
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
        <el-dialog title="选择车道" :visible.sync="dialogVisible3" width="18%" >
          <el-select v-model="cd">
            <el-option v-for="item in options"
            :key="item.value"
            :value="item.value"
            :label="item.labek"></el-option>
          </el-select>
          <span slot="footer" class="dialog-footer">
            <el-button @click="dialogVisible3 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
            <el-button type="primary" icon="el-icon-check" @click="exportgzsd()" size="small">确 定</el-button>
          </span>
        </el-dialog>
        <el-dialog title="输入设计值" :visible.sync="dialogVisible4" width="40%"  >
          <el-form ref="sjzForm" :model="sjzdata"  label-width="120px" size="small" style="padding-right: 40px;" :rules="rules">
            
            
            <el-form-item label="设计值" prop="sjz">
            <el-input v-model="sjzdata.sjz"></el-input>
            </el-form-item>
            
          </el-form>
          <span slot="footer" class="dialog-footer">
            <el-button @click="dialogVisible4 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
            <el-button type="primary"  icon="el-icon-check" @click="scjdb()" size="small">确 定</el-button>
          </span>
        </el-dialog>
      </div>
    </template>
    <script>
    // 引入定义接口js文件
    import api from "@/api/project/fbgc/lmgc/gzsd";
    import { getSystemErrorMap } from "util";
    import FileSaver from "file-saver"
    import { Loading } from "element-ui";
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
          dialogVisible3:false,
          dialogVisible4:false,
          username:Cookies.get('username'),
          userid:Cookies.get('userid'),
          type:Cookies.get('type'),
          sjzdata:{
            sjz:""
          },
          rules:{
            sjz:[
              {
                required:true,
                message:"设计值不能为空！"
              },
              {
                type:'number',
                message:"设计值必须为数字！",
                transform: (value) => Number(value)
              }
            ]
          },
          cd:2,
          options:
                [
                    {
                      value:2,
                      label:2,
                    },
                    {
                      value:3,
                      label:3
                    },
                    {
                      value:4,
                      label:4
                    },
                    {
                      value:5,
                      label:5
                    }
                ],
          proname:this.$route.query.projecttitle,
          htd:this.$route.query.htdname,
          fbgcName:this.$route.query.fbgcName,
          hdgqd: {},
          file:'',// 待上传文件
          fileList:[],
          modifyData:{},
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；"+"\n"+"2，若是主线，则路线类型填写主线，匝道标识不填；若是互通，请填写互通的名称，并且填写匝道标识，注意互通的名称要和混凝土路面匝道清单的一致；\n"
          +"3，若匝道有收费站的情况，请不要上传收费站的数据，以起点桩号为准；\n4，填写模板如下：",
          rules1:{
            
            qdzh:[
              {
                required:true,
                message:"起点桩号不能为空！"
              }
            ],
            zdzh:[
              {
                required:true,
                message:"终点桩号不能为空！"
              }
            ],
            
            mtd:[
              {
                required:true,
                message:"代表MTD不能为空！"
              }
            ],
            lxbs:[
              {
                required:true,
                message:"路线类型不能为空！"
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
          this.$refs["sjzForm"].validate((valid)=>{
              if(valid){
                this.dialogVisible4=false
                let commonInfoVo={}
                commonInfoVo.proname=this.proname
                commonInfoVo.htd=this.htd
                commonInfoVo.sjz=this.sjzdata.sjz
                commonInfoVo.userid=this.userid
                let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
                api.scjdb(commonInfoVo).then((response) => {
                  return response.data.text()       
                })
                .then(res=>{
                  
                  let code = JSON.parse(res).code
                  if(code == 200){
                    api.download(commonInfoVo.proname,commonInfoVo.htd,commonInfoVo.userid).then((response)=>{
                    
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
                
              }
              else{
                return false
              }
            })
          
    
        },
        ckjdb(){
            this.dialogVisible1=true
            let commonInfoVo={}
            commonInfoVo.proname=this.proname
            commonInfoVo.htd=this.htd
            commonInfoVo.userid=this.userid
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('tttt',res)
              this.descriptions=res.data
            })
        },
        importgzsd() {
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
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'上传中，请不要关闭页面！'})
          
          api.importgzsd(fd).then((res)=>{
            console.log('res',res)
            if(res.message=='成功'){
              this.fetchData();
              this.$message({
                type: "success",
                message: "上传成功!",
              });
              this.dialogImportVisible = false;
              this.fileList=[]
              load.close()
            }
            else{
              load.close()
              this.$alert('上传失败！')
              
            }
          }).catch((err)=>{
            load.close()
          });
          
          
          
        },
        exportgzsd() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          this.dialogVisible3=false
          api.exportgzsd(this.cd).then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `构造深度实测数据` + ".xlsx";
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