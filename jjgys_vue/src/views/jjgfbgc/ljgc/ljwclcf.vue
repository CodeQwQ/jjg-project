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
              @click="exportljwclcf()"
              style="margin-left: 0px"
              icon="el-icon-bottom"
              size="mini"
              >导出路基弯沉落锤法模板文件</el-button
            >
            <!-- :disabled="$hasBP('bnt.ql.export')  === true" -->
            <el-button
              class="btn-add"
              type="success"
              @click="importljwclcf()"
              style="margin-left: 0px"
              icon="el-icon-top"
              size="mini"
              >导入路基弯沉落锤法数据文件</el-button
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
          <el-table-column prop="zh" label="桩号" width="220px" />
          <el-table-column prop="sjwcz" label="设计弯沉值(0.01mm)" width="180px"/>
          <el-table-column prop="jgcc" label="结构层次" />
          <el-table-column prop="wdyxxs" label="温度影响系数" width="120px"/>
          <el-table-column prop="jjyxxs" label="季节影响系数" width="120px"/>
          

          <el-table-column prop="jglx" label="结构类型" />
          <el-table-column prop="mbkkzb" label="目标可靠指标" width="120px"/>
          <el-table-column prop="sdyxxs" label="湿度影响系数" width="120px"/>
          <el-table-column prop="lcz" label="落锤重(T)" width="120px"/>
          
          <el-table-column prop="yqmc" label="仪器名称" />
          <el-table-column prop="cjzh" label="抽检桩号" />
          <el-table-column prop="cd" label="车道" />
          <el-table-column prop="scwcz" label="实测弯沉值(0.01㎜)" width="180px"/>
          <el-table-column prop="lbwd" label="路表温度" />
          <el-table-column prop="xh" label="序号" width="50px"/>  
          <el-table-column prop="bz" label="备注"  />
          <el-table-column prop="createtime" label="创建时间" />
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

        <el-alert
        :title="alertTitle"
        :type="alertType">
        </el-alert>
    
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
          
          :before-close="handleClose"
        >
        <span style="white-space: pre-wrap;">{{ instruction }}</span>
      
        <img v-viewer style="height: 100px;width: 450px;" src="@/assets/填写说明/路基/路基弯沉落锤法.png" alt="" >
        <div style="color: rgb(84, 151, 198);" align="middle">双击查看大图</div>
          <span slot="footer" class="dialog-footer">
            <el-button type="primary" @click="dialogVisible = false"
              >确 定</el-button
            >
          </span>
        </el-dialog>
        <el-dialog
         title="路基弯沉落锤法检测结果"
         :visible.sync="dialogVisible1">
         <el-descriptions :column="1"  border>
          <el-descriptions-item>
              <template slot="label" >评定单元合格数</template>
              {{descriptions["合格单元数"]}}
          </el-descriptions-item>
          <el-descriptions-item>
              <template slot="label" >评定单元总数</template>
              {{descriptions["检测单元数"]}}
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
          <el-form ref="dataForm" :model="modifyData"  label-width="170px" size="small" style="padding-right: 40px;" :rules="rules">
            
            <el-form-item label="检测时间" prop="jcsj">
              <el-input v-model="modifyData.jcsj"></el-input>
            </el-form-item>
            <el-form-item label="桩号" prop="zh">
              <el-input v-model="modifyData.zh"></el-input>
            </el-form-item>
            <el-form-item label="设计弯沉值(0.01mm)" prop="scwcz">
              <el-input v-model="modifyData.sjwcz"></el-input>
            </el-form-item>
            <el-form-item label="结构层次" prop="jgcc">
              <el-input v-model="modifyData.jgcc"></el-input>
            </el-form-item>
            <el-form-item label="温度影响系数" prop="wdyxxs">
              <el-input v-model="modifyData.wdyxxs"></el-input>
            </el-form-item>
            <el-form-item label="季节影响系数" prop="jjyxxs">
              <el-input v-model="modifyData.jjyxxs"></el-input>
            </el-form-item>
            <el-form-item label="结构类型" prop="jglx">
              <el-input v-model="modifyData.jglx"></el-input>
            </el-form-item>
            <el-form-item label="目标可靠指标" prop="mbkkzb">
              <el-input v-model="modifyData.mbkkzb"></el-input>
            </el-form-item>
            <el-form-item label="湿度影响系数" prop="sdyxxs">
              <el-input v-model="modifyData.sdyxxs"></el-input>
            </el-form-item>
            <el-form-item label="落锤重(T)" prop="lcz">
              <el-input v-model="modifyData.lcz"></el-input>
            </el-form-item>
            <el-form-item label="仪器名称" prop="yqmc">
              <el-input v-model="modifyData.yqmc"></el-input>
            </el-form-item>
            <el-form-item label="抽检桩号" prop="cjzh">
              <el-input v-model="modifyData.cjzh"></el-input>
            </el-form-item>
            <el-form-item label="车道" prop="cd">
              <el-input v-model="modifyData.mbkkzb"></el-input>
            </el-form-item>
            <el-form-item label="实测弯沉值(0.01㎜)" prop="scwcz">
              <el-input v-model="modifyData.scwcz"></el-input>
            </el-form-item>
            <el-form-item label="路表温度" prop="lbwd">
              <el-input v-model="modifyData.lbwd"></el-input>
            </el-form-item>
            <el-form-item label="序号" prop="xh">
              <el-input v-model="modifyData.xh"></el-input>
            </el-form-item>
            <el-form-item label="备注" prop="bz">
              <el-input v-model="modifyData.bz"></el-input>
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
    import api from "@/api/project/fbgc/ljgc/ljwclcf.js";
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
             '合格单元数':0,
             '检测单元数':0,
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
          hdgqd: {},
          file:'',// 待上传文件
          fileList:[],
          modifyData:{},
          alertTitle:'成功获取数据',
          alertType:'success',
          username:Cookies.get('username'),
          userid:Cookies.get('userid'),
          type:Cookies.get('type'),
          instruction:"1，日期格式：yyyy/MM/dd，例如：2021/05/23；\n2，序号的含义为：相同序号数字的数据是在一起计算的；\n3，填写模板如下：",
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
            sjwcz:[
              {
                required:true,
                message:"设计弯沉值不能为空！"
              }
            ],
            jgcc:[
              {
                required:true,
                message:"结构层次不能为空！"
              }
            ],
            wdyxxs:[
              {
                required:true,
                message:"温度影响系数不能为空！"
              }
            ],
            jjyxxs:[
              {
                required:true,
                message:"季节影响系数不能为空！"
              }
            ],
            jglx:[
              {
                required:true,
                message:"结构类型不能为空！"
              }
            ],
            sdyxxs:[
              {
                required:true,
                message:"湿度影响系数不能为空！"
              }
            ],
            lcz:[
              {
                required:true,
                message:"落锤重不能为空！"
              }
            ],
            yqmc:[
              {
                required:true,
                message:"仪器名称不能为空！"
              }
            ],
            cjzh:[
              {
                required:true,
                message:"抽检桩号不能为空！"
              }
            ],
            cd:[
              {
                required:true,
                message:"车道！"
              }
            ],
            scwcz:[
              {
                required:true,
                message:"实测弯沉值不能为空！"
              }
            ],
            lbwd:[
              {
                required:true,
                message:"路表温度不能为空！"
              }
            ],
            xh:[
              {
                required:true,
                message:"序号"
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
              this.descriptions['合格单元数']=res.data[0]["合格单元数"]
              this.descriptions['检测单元数']=res.data[0]["检测单元数"]
              this.descriptions['合格率']=res.data[0]["合格率"]
            })
        },
        importljwclcf() {
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
          api.importljwclcf(fd).then((res)=>{
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
        exportljwclcf() {
          let projectname = this.proname;
          let htd =this.htd
          let fbgc=this.fbgcName
          
          
          api.exportljwclcf().then((res) => {
            const objectUrl = URL.createObjectURL(
              new Blob([res.data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              })
            );
            const link = document.createElement("a");
            // 设置导出的文件名称
            link.download = `02路基弯沉(落锤法)实测数据` + ".xlsx";
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
            commonInfoVo.fbgc='路基土石方'
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
          this.searchObj.fbgc='路基土石方'
          this.searchObj.userid=this.userid
          console.log('eeee',this.htd)
          // 调用api
          api.pageList(this.page, this.limit, this.searchObj).then((response) => {
            this.list = response.data.records;
            this.total = response.data.total;
            console.log(this.total)
            // 如果记录数为-1（没有记录但是有鉴定表），显示警告信息
            if(this.total == -1){
            this.alertTitle = '没有查询到相关信息，但是该模块在后台有鉴定表，若想弃用该模块请点击 全部删除 按钮删除鉴定表！';
            this.alertType = 'warning';
            }else if(this.total == 0){
            this.alertTitle = '没有查询到相关信息，且该模块在后台没有鉴定表。';
            this.alertType = 'success';
            }else{
            this.alertTitle = '成功查询到相关信息。';
            this.alertType = 'success';
            }
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