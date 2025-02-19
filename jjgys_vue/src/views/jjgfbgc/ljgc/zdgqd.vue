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
              @click="exportzdgqd()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出支挡砼强度模板文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importzdgqd()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入支挡砼强度数据文件</el-button
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
            ><el-button
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
          <el-table-column prop="zh" label="桩号" width="260px" />
          <el-table-column prop="bw1" label="部位1" />
          <el-table-column prop="bw2" label="部位2" />
          <el-table-column prop="cdz1" label="测定值1" />
          <el-table-column prop="z2" label="值2" />
          <el-table-column prop="z3" label="值3" />
          <el-table-column prop="z4" label="值4" />
          <el-table-column prop="z5" label="值5" />
          <el-table-column prop="z6" label="值6" />
          <el-table-column prop="z7" label="值7" />
          <el-table-column prop="z8" label="值8" />
          <el-table-column prop="z9" label="值9" />
          <el-table-column prop="z10" label="值10" />
          <el-table-column prop="z11" label="值11" />
          <el-table-column prop="z12" label="值12" />
          <el-table-column prop="z13" label="值13" />
          <el-table-column prop="z14" label="值14" />
          <el-table-column prop="z15" label="值15" />
          <el-table-column prop="z16" label="值16" />
          <el-table-column prop="htjd" label="回弹角度"  />
          <el-table-column prop="jzm" label="浇筑面" />
          <el-table-column prop="thsd" label="碳化深度" />
          <el-table-column prop="sfbs" label="是否泵送" />
          <el-table-column prop="sjqd" label="设计强度" />
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
      
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/路基/支挡砼强度.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="支挡砼强度检测结果"
         :visible.sync="dialogVisible1">
         <el-descriptions :column="1"  border>
          <el-descriptions-item>
              <template slot="label" >合格点数</template>
              {{descriptions["合格点数"]}}
          </el-descriptions-item>
          <el-descriptions-item>
              <template slot="label" >总点数</template>
              {{descriptions["总点数"]}}
          </el-descriptions-item>
          <el-descriptions-item>
              <template slot="label" >合格率(%)</template>
              {{descriptions["合格率"]}}
          </el-descriptions-item>
         </el-descriptions>
         <span slot="footer" class="dialog-footer">
          <el-button type="primary"  @click="dialogVisible1 = false">关闭</el-button>
         </span>
         
        </el-dialog>
        <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
          <el-form ref="dataForm" :model="modifyData"  label-width="100px" size="small" style="padding-right: 40px;" :rules="rules">
            <el-form-item label="检测时间" prop="jcsj">
            <el-input v-model="modifyData.jcsj"></el-input>
            </el-form-item>
            <el-form-item label="桩号" prop="zh">
            <el-input v-model="modifyData.zh"></el-input>
            </el-form-item>
            <el-form-item label="部位1" prop="bw1">
            <el-input v-model="modifyData.bw1"></el-input>
            </el-form-item>
            <el-form-item label="部位2" prop="bw2">
            <el-input v-model="modifyData.bw2"></el-input>
            </el-form-item>
            <el-form-item label="测定值1" prop="cdz1">
            <el-input v-model="modifyData.cdz1"></el-input>
            </el-form-item>
            <el-form-item label="值2" prop="z2">
            <el-input v-model="modifyData.z2"></el-input>
            </el-form-item>
            <el-form-item label="值3" prop="z3">
            <el-input v-model="modifyData.z3"></el-input>
            </el-form-item>
            <el-form-item label="值4" prop="z4">
            <el-input v-model="modifyData.z4"></el-input>
            </el-form-item>
            <el-form-item label="值5" prop="z5">
            <el-input v-model="modifyData.z5"></el-input>
            </el-form-item>
            <el-form-item label="值6" prop="z6">
            <el-input v-model="modifyData.z6"></el-input>
            </el-form-item>
            <el-form-item label="值7" prop="z7">
            <el-input v-model="modifyData.z7"></el-input>
            </el-form-item>
            <el-form-item label="值8" prop="z8">
            <el-input v-model="modifyData.z8"></el-input>
            </el-form-item>
            <el-form-item label="值9" prop="z9">
            <el-input v-model="modifyData.z9"></el-input>
            </el-form-item>
            <el-form-item label="值10" prop="z10">
            <el-input v-model="modifyData.z10"></el-input>
            </el-form-item>
            <el-form-item label="值11" prop="z11">
            <el-input v-model="modifyData.z11"></el-input>
            </el-form-item>
            <el-form-item label="值12" prop="z12">
            <el-input v-model="modifyData.z12"></el-input>
            </el-form-item>
            <el-form-item label="值13" prop="z13">
            <el-input v-model="modifyData.z13"></el-input>
            </el-form-item>
            <el-form-item label="值14" prop="z14">
            <el-input v-model="modifyData.z14"></el-input>
            </el-form-item>
            <el-form-item label="值15" prop="z15">
            <el-input v-model="modifyData.z15"></el-input>
            </el-form-item>
            <el-form-item label="值16" prop="z16">
            <el-input v-model="modifyData.z16"></el-input>
            </el-form-item>
            <el-form-item label="回弹角度" prop="htjd">
            <el-input v-model="modifyData.htjd"></el-input>
            </el-form-item>
            <el-form-item label="浇筑面" prop="jzm">
            <el-input v-model="modifyData.jzm"></el-input>
            </el-form-item>
            <el-form-item label="碳化深度" prop="thsd">
            <el-input v-model="modifyData.thsd"></el-input>
            </el-form-item>
            <el-form-item label="是否泵送" prop="sfbs">
            <el-input v-model="modifyData.sfbs"></el-input>
            </el-form-item>
            <el-form-item label="设计强度" prop="sjqd">
            <el-input v-model="modifyData.sjqd"></el-input>
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
    import api from "@/api/project/fbgc/ljgc/zdgqd.js";
    import { getSystemErrorMap } from "util";
    import FileSaver from "file-saver"
    import { Loading } from 'element-ui';
    import Cookies from 'js-cookie'
    export default {
      name:'zdgqd',
      data() {
        // 初始值
        let value = this.$route.query.projecttitle
        return {
          data: {
            proname: value,
          },
          descriptions: {
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
          hdgqd: {},
          file:'',// 待上传文件
          fileList:[],
          modifyData:{},
          proname:this.$route.query.projecttitle,
         htd:this.$route.query.htdname,
          username:Cookies.get('username'),
          userid:Cookies.get('userid'),
          type:Cookies.get('type'),
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；\n2，设计强度不填写C；\n3，填写模板如下：",
          rules:{
            jcsj:[
              {
                required:true,
                message:"检测时间不能为空！"
              }
            ],
            zh:[
              {
                required:true,
                message:"桩号不能为空！"
              }
            ],
            bw1:[
              {
                required:true,
                message:"部位1不能为空！"
              }
            ],
            bw2:[
              {
                required:true,
                message:"部位2不能为空！"
              }
            ],
            cdz1:[
              {
                required:true,
                message:"测定值1不能为空！"
              }
            ],
            z2:[
              {
                required:true,
                message:"值2不能为空！"
              }
            ],
            z3:[
              {
                required:true,
                message:"值3不能为空！"
              }
            ],
            z4:[
              {
                required:true,
                message:"值4不能为空！"
              }
            ],
            z5:[
              {
                required:true,
                message:"值5不能为空！"
              }
            ],
            z6:[
              {
                required:true,
                message:"值6不能为空！"
              }
            ],
            z7:[
              {
                required:true,
                message:"值7不能为空！"
              }
            ],
            z8:[
              {
                required:true,
                message:"值8不能为空！"
              }
            ],
            z9:[
              {
                required:true,
                message:"值9不能为空！"
              }
            ],
            z10:[
              {
                required:true,
                message:"值10不能为空！"
              }
            ],
            z11:[
              {
                required:true,
                message:"值11不能为空！"
              }
            ],
            z12:[
              {
                required:true,
                message:"值12不能为空！"
              }
            ],
            z13:[
              {
                required:true,
                message:"值13不能为空！"
              }
            ],
            z14:[
              {
                required:true,
                message:"值14不能为空！"
              }
            ],
            z15:[
              {
                required:true,
                message:"值15不能为空！"
              }
            ],
            z16:[
              {
                required:true,
                message:"值16不能为空！"
              }
            ],
            htjd:[
              {
                required:true,
                message:"回弹角度不能为空！"
              }
            ],
            jzm:[
              {
                required:true,
                message:"浇筑面不能为空！"
              }
            ],
            thsd:[
              {
                required:true,
                message:"碳化深度不能为空！"
              }
            ],
            sfbs:[
              {
                required:true,
                message:"是否泵送不能为空！"
              }
            ],
            sjqd:[
              {
                required:true,
                message:"设计强度不能为空！"
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
        console.log(this.$route.query.projecttitle);
        //alert(this.$route.matched[1].meta.title)
        this.fetchData();
      },
      // 'http://localhost:8800/jjg/fbgc/ljgc/hdgqd/importhdgqd'
      methods: {
        handleClose(){
    
        },
        scjdb(){
          let commonInfoVo={}
          commonInfoVo.proname=this.$route.query.projecttitle
          commonInfoVo.htd=this.$route.query.htdname
          commonInfoVo.fbgc='支挡工程'
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
            commonInfoVo.proname=this.$route.query.projecttitle
            commonInfoVo.htd=this.$route.query.htdname
            commonInfoVo.fbgc='支挡工程'
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('ss',res.data)
              this.descriptions['合格点数']=res.data[0]["合格点数"]
              this.descriptions['总点数']=res.data[0]["总点数"]
              this.descriptions['合格率']=res.data[0]["合格率"]
            })
        },
        importzdgqd() {
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
          fd.append('proname',this.$route.query.projecttitle)
          fd.append('htd',this.$route.query.htdname)
          fd.append('fbgc','支挡工程')
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importzdgqd(fd).then((res)=>{
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
        exportzdgqd() {
          let projectname = this.$route.query.projecttitle;
          let htd =this.$route.query.htdname
          let fbgc=this.$route.query.fbgcName
          console.log(projectname)
          debugger
          api.exportzdgqd().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `10支挡砼强度实测数据` + ".xlsx";
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
            console.log('ss',idList)
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
            commonInfoVo.fbgc='支挡工程'
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
          this.searchObj.proname = this.$route.query.projecttitle;
          this.searchObj.htd=this.$route.query.htdname
          this.searchObj.fbgc='支挡工程'
          this.searchObj.userid=this.userid
          // 调用api
          api.pageList(this.page, this.limit, this.searchObj).then((response) => {
            this.list = response.data.records;
            this.total = response.data.total;
            console.log('jjj',this.list)
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
    </style>