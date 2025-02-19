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
            <el-form-item label="位置及类型:">
              <el-input v-model="searchObj.wzjlx" />
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
              @click="exportjabxfhl()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出交安波形防护栏文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importjabxfhl()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入交安波形防护栏数据文件</el-button
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
          <el-table-column prop="htd" label="合同段名称" width="200px" />
          <el-table-column prop="jcsj" label="检测时间" width="200px" />
          <el-table-column prop="wzjlx" label="位置及类型" width="160px" />
          <el-table-column prop="jdhdgdz" label="基底厚度规定值(mm)" width="120px" />
          <el-table-column prop="jdhdscz1" label="基底厚度实测值1(mm)" width="130px"/>
          <el-table-column prop="jdhdscz2" label="基底厚度实测值2(mm)" width="130px"/>
          <el-table-column prop="jdhdscz3" label="基底厚度实测值3(mm)" width="130px"/>
          <el-table-column prop="jdhdscz4" label="基底厚度实测值4(mm)" width="130px"/>
          <el-table-column prop="jdhdscz5" label="基底厚度实测值5(mm)" width="130px"/>
          <el-table-column prop="lzbhgdz" label="立柱壁厚规定值(mm)" width="120px" />
          <el-table-column prop="lzbhscz1" label="立柱壁厚实测值1(mm)" width="130px"/>
          <el-table-column prop="lzbhscz2" label="立柱壁厚实测值2(mm)" width="130px"/>
          <el-table-column prop="lzbhscz3" label="立柱壁厚实测值3(mm)" width="130px"/>
          <el-table-column prop="lzbhscz4" label="立柱壁厚实测值4(mm)" width="130px"/>
          <el-table-column prop="lzbhscz5" label="立柱壁厚实测值5(mm)" width="130px"/>
          <el-table-column prop="zxgdgdz" label="中心高度规定值(mm)" width="120px" />
          <el-table-column prop="zxgdyxpsz" label="中心高度允许偏差+(mm)" width="150px" />
          <el-table-column prop="zxgdyxpsf" label="中心高度允许偏差-(mm)" width="150px" />
          <el-table-column prop="zxgdscz1" label="中心高度实测值1(mm)" width="130px"/>
          <el-table-column prop="zxgdscz2" label="中心高度实测值2(mm)" width="130px"/>
          <el-table-column prop="zxgdscz3" label="中心高度实测值3(mm)" width="130px"/>
          <el-table-column prop="zxgdscz4" label="中心高度实测值4(mm)" width="130px"/>
          <el-table-column prop="zxgdscz5" label="中心高度实测值5(mm)" width="130px"/>
          <el-table-column prop="mrsdgdz" label="埋入深度规定值(mm)" width="120px"/>
          <el-table-column prop="mrsdscz" label="埋入深度实测值(mm)" width="120px"/>

         

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
      
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/交安/交安波形防护栏.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="交安波形防护栏检测结果"
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
        <el-table-column prop="htd" label="合同段名称" />
          <el-table-column prop="检测项目" label="检测项目" width="250px"/>
          <el-table-column prop="规定值或允许偏差" label="规定值或允许偏差"  />
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
          <el-form-item label="合同段名称" prop="htd">
            <el-input v-model="modifyData.htd"></el-input>
            </el-form-item>
            <el-form-item label="检测时间" prop="jcsj">
            <el-input v-model="modifyData.jcsj"></el-input>
            </el-form-item>
            <el-form-item label="位置及类型" prop="wzjlx">
            <el-input v-model="modifyData.wzjlx"></el-input>
            </el-form-item>
            <el-form-item label="基底厚度规定值" prop="jdhdgdz">
            <el-input v-model="modifyData.jdhdgdz"></el-input>
            </el-form-item>
            <el-form-item label="基底厚度实测值1" prop="jdhdscz1">
            <el-input v-model="modifyData.jdhdscz1"></el-input>
            </el-form-item>
            <el-form-item label="基底厚度实测值2" prop="jdhdscz2">
            <el-input v-model="modifyData.jdhdscz2"></el-input>
            </el-form-item>
            <el-form-item label="基底厚度实测值3" prop="jdhdscz3">
            <el-input v-model="modifyData.jdhdscz3"></el-input>
            </el-form-item>
            <el-form-item label="基底厚度实测值4" prop="jdhdscz4">
            <el-input v-model="modifyData.jdhdscz4"></el-input>
            </el-form-item>
            <el-form-item label="基底厚度实测值5" prop="jdhdscz5">
            <el-input v-model="modifyData.jdhdscz5"></el-input>
            </el-form-item>
            <el-form-item label="立柱壁厚规定值" prop="lzbhgdz">
            <el-input v-model="modifyData.lzbhgdz"></el-input>
            </el-form-item>
            <el-form-item label="立柱壁厚实测值1" prop="lzbhscz1">
            <el-input v-model="modifyData.lzbhscz1"></el-input>
            </el-form-item>
            <el-form-item label="立柱壁厚实测值2" prop="lzbhscz2">
            <el-input v-model="modifyData.lzbhscz2"></el-input>
            </el-form-item>
            <el-form-item label="立柱壁厚实测值3" prop="lzbhscz3">
            <el-input v-model="modifyData.lzbhscz3"></el-input>
            </el-form-item>
            <el-form-item label="立柱壁厚实测值4" prop="lzbhscz4">
            <el-input v-model="modifyData.lzbhscz4"></el-input>
            </el-form-item>
            <el-form-item label="立柱壁厚实测值5" prop="lzbhscz5">
            <el-input v-model="modifyData.lzbhscz5"></el-input>
            </el-form-item>
            <el-form-item label="中心高度规定值" prop="zxgdgdz">
            <el-input v-model="modifyData.zxgdgdz"></el-input>
            </el-form-item>
            <el-form-item label="中心高度允许偏差+" prop="zxgdyxpsz">
            <el-input v-model="modifyData.zxgdyxpsz"></el-input>
            </el-form-item>
            <el-form-item label="中心高度允许偏差-" prop="zxgdyxpsf">
            <el-input v-model="modifyData.zxgdyxpsf"></el-input>
            </el-form-item>
            <el-form-item label="中心高度实测值1" prop="zxgdscz1">
            <el-input v-model="modifyData.zxgdscz1"></el-input>
            </el-form-item>
            <el-form-item label="中心高度实测值2" prop="zxgdscz2">
            <el-input v-model="modifyData.zxgdscz2"></el-input>
            </el-form-item>
            <el-form-item label="中心高度实测值3" prop="zxgdscz3">
            <el-input v-model="modifyData.zxgdscz3"></el-input>
            </el-form-item>
            <el-form-item label="中心高度实测值4" prop="zxgdscz4">
            <el-input v-model="modifyData.zxgdscz4"></el-input>
            </el-form-item>
            <el-form-item label="中心高度实测值5" prop="zxgdscz5">
            <el-input v-model="modifyData.zxgdscz5"></el-input>
            </el-form-item>

            
            <el-form-item label="埋入深度规定值" prop="mrsdgdz">
            <el-input v-model="modifyData.mrsdgdz"></el-input>
            </el-form-item>
            <el-form-item label="埋入深度实测值" prop="mrsdscz">
            <el-input v-model="modifyData.mrsdscz"></el-input>
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
    import api from "@/api/jgproject/fczb/jabxfhl.js";
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
          username:Cookies.get('username'),
          userid:Cookies.get('userid'),
          type:Cookies.get('type'),
          proname:this.$route.query.projecttitle,
          htd:this.$route.query.htdname,
          fbgcName:this.$route.query.fbgcName,
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
            wzjlx:[
              {
                required:true,
                message:"位置及类型不能为空！"
              }
            ],
            jdhdgdz:[
              {
                required:true,
                message:"基底厚度规定值不能为空！"
              }
            ],
            jdhdscz1:[
              {
                required:true,
                message:"基底厚度实测值1不能为空！"
              }
            ],
            jdhdscz2:[
              {
                required:true,
                message:"基底厚度实测值2不能为空！"
              }
            ],
            jdhdscz3:[
              {
                required:true,
                message:"基底厚度实测值3不能为空！"
              }
            ],
            lzbhgdz:[
              {
                required:true,
                message:"立柱壁厚规定值不能为空！"
              }
            ],
            lzbhscz1:[
              {
                required:true,
                message:"立柱壁厚实测值1不能为空！"
              }
            ],
            lzbhscz2:[
              {
                required:true,
                message:"立柱壁厚实测值2不能为空！"
              }
            ],
            lzbhscz3:[
              {
                required:true,
                message:"立柱壁厚实测值3不能为空！"
              }
            ],
            zxgdgdz:[
              {
                required:true,
                message:"中心高度规定值不能为空！"
              }
            ],
            zxgdyxpsz:[
              {
                required:true,
                message:"中心高度允许偏差+不能为空！"
              }
            ],
            zxgdyxpsf:[
              {
                required:true,
                message:"中心高度允许偏差-不能为空！"
              }
            ],
            zxgdscz1:[
              {
                required:true,
                message:"中心高度实测值1不能为空！"
              }
            ],
            zxgdscz2:[
              {
                required:true,
                message:"中心高度实测值2不能为空！"
              }
            ],
            zxgdscz3:[
              {
                required:true,
                message:"中心高度实测值3不能为空！"
              }
            ],
            mrsdgdz:[
              {
                required:true,
                message:"埋入深度规定值不能为空！"
              }
            ],
            mrsdscz:[
              {
                required:true,
                message:"埋入深度实测值不能为空！"
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
          commonInfoVo.fbgc='防护栏'
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
            commonInfoVo.fbgc='防护栏'
            api.lookjdb(commonInfoVo).then((res)=>{
              console.log('tttt',res)
              this.descriptions=res.data
            })
        },
        importjabxfhl() {
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
          fd.append('fbgc','防护栏')
          fd.append('userid',this.userid)
          let load=Loading.service({fullscreen:true,text:'加载中,请稍后!'})
          api.importjabxfhl(fd).then((res)=>{
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
        exportjabxfhl() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          debugger
          api.exportjabxfhl().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `03交安波形防护栏实测数据` + ".xlsx";
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
            commonInfoVo.fbgc='防护栏'
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
          this.searchObj.fbgc='防护栏'
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