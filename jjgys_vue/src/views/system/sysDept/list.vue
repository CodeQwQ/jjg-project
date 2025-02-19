<template>
  <div class="app-container">
    <!--查询表单-->
    <div class="search-div">
      <el-form label-width="70px" size="small">
        <el-row>
          <el-col :span="24">
            <el-input style="width: 20%" v-model="searchObj.name" placeholder="岗位名称"></el-input>
            <el-button type="primary" icon="el-icon-search" size="mini"  @click="fetchData()">搜索</el-button>
            <!--<el-button icon="el-icon-refresh" size="mini" @click="resetData">重置</el-button>-->
          </el-col>
        </el-row>
        <el-row style="display:flex">
          <!-- <el-button type="primary" icon="el-icon-search" size="mini"  @click="fetchData()">搜索</el-button>
          <el-button icon="el-icon-refresh" size="mini" @click="resetData">重置</el-button> -->
        </el-row>
      </el-form>
    </div>
    <!-- 工具条 -->
  <div class="tools-div" v-show="username=='admin'&&type==1">
    <el-button type="success"  icon="el-icon-plus" size="mini" @click="addCompany">添加公司</el-button>
    <el-button type="success"  icon="el-icon-plus" size="mini" @click="addDept">添加部门</el-button>
    <el-button type="success"  icon="el-icon-edit" size="mini" @click="edit">修改</el-button>
    <el-button type="danger"  icon="el-icon-delete" size="mini" @click="remove">删除</el-button>
  </div>
    <!-- 表格 -->
    <el-table
      
      :data="list"
      
      border
      style="width: 100%;margin-top: 10px;"
      row-key="id"
      default-expand-all
      :tree-props="{ children: 'children', hasChildren: 'hasChildren' }"
      @selection-change="handleSelectionChange"
      >

      

      
      <el-table-column type="selection" />
      <el-table-column prop="name" label="部门名称" />
      <el-table-column prop="leader" label="负责人" />
      <el-table-column prop="phone" label="电话" />
      <el-table-column prop="createTime" label="创建时间" width="160"/>
      <!--<el-table-column label="操作"  align="center" fixed="right">
        <template slot-scope="scope">
          <el-button type="primary" icon="el-icon-edit" size="mini" @click="edit(scope.row.id)" title="修改"/>
          <el-button type="danger" icon="el-icon-delete" size="mini" @click="removeDataById(scope.row.id)" title="删除" />
        </template>
      </el-table-column>-->
      
    </el-table>

  
    <el-dialog title="添加公司" :visible.sync="dialogVisible" width="40%" >
      <el-form ref="dataForm" :model="sysCompany" :rules="addrules" label-width="100px" size="small" style="padding-right: 40px;">
        <el-form-item label="公司名称" prop="name">
          <el-input v-model="sysCompany.name"/>
        </el-form-item>
        <el-form-item label="负责人" prop="leader">
          <el-input v-model="sysCompany.leader"/>
        </el-form-item>
        <el-form-item label="电话" prop="phone">
          <el-input v-model="sysCompany.phone"/>
        </el-form-item>
        <el-form-item label="创建时间" prop="createTime">
          <el-date-picker
          v-model="sysCompany.createTime"
          value-format="yyyy-MM-dd HH:mm:ss"
        />
        </el-form-item>
      </el-form>
      <span slot="footer" class="dialog-footer">
        <el-button @click="dialogVisible = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
        <el-button type="primary" icon="el-icon-check" @click="saveCompany()" size="small">确 定</el-button>
      </span>
    </el-dialog>

    <el-dialog title="添加部门" :visible.sync="dialogVisible1" width="40%" >
      <el-form ref="dataForm1" :model="sysDept" :rules="addrules" label-width="100px" size="small" style="padding-right: 40px;">
        <el-form-item label="公司名称" prop="companyId" >
             
             <el-select
               style="width: 100%"
               v-model="sysDept.companyId"
               
               
               
               
               placeholder="请选择公司"
             >
               <el-option
                 v-for="(item,index) in options"
                 :key="index"
                 :label="item.name"
                 :value="item.id"
               >
               </el-option>
             </el-select>
           </el-form-item>
        <el-form-item label="部门名称" prop="name">
          <el-input v-model="sysDept.name"/>
        </el-form-item>
        <el-form-item label="负责人" prop="leader">
          <el-input v-model="sysDept.leader"/>
        </el-form-item>
        <el-form-item label="电话" prop="phone">
          <el-input v-model="sysDept.phone"/>
        </el-form-item>
        <el-form-item label="创建时间" prop="createTime">
          <el-date-picker
          v-model="sysDept.createTime"
          value-format="yyyy-MM-dd HH:mm:ss"
        />
        </el-form-item>
      </el-form>
      <span slot="footer" class="dialog-footer">
        <el-button @click="dialogVisible1 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
        <el-button type="primary" icon="el-icon-check" @click="saveDept()" size="small">确 定</el-button>
      </span>
    </el-dialog>

    <el-dialog title="修改" :visible.sync="dialogVisible2" width="40%" >
      <el-form ref="dataForm2" :model="modifyData" :rules="addrules" label-width="100px" size="small" style="padding-right: 40px;">
        
        <el-form-item label="公司/部门名称" prop="name">
          <el-input v-model="modifyData.name"/>
        </el-form-item>
        <el-form-item label="负责人" prop="leader">
          <el-input v-model="modifyData.leader"/>
        </el-form-item>
        <el-form-item label="电话" prop="phone">
          <el-input v-model="modifyData.phone"/>
        </el-form-item>
        <el-form-item label="创建时间" prop="createTime">
          <el-date-picker
          v-model="modifyData.createTime"
          value-format="yyyy-MM-dd HH:mm:ss"
        />
        </el-form-item>
      </el-form>
      <span slot="footer" class="dialog-footer">
        <el-button @click="dialogVisible2 = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
        <el-button type="primary" icon="el-icon-check" @click="saveUpdate()" size="small">确 定</el-button>
      </span>
    </el-dialog>

  </div>
</template>
<script>
//引入定义接口的js文件
import api from '@/api/system/dept'
import Cookies from 'js-cookie'
export default {
  //定义初始值
  data() {
    return {
      listLoading:false,//是否显示加载
      list:[],//操作日志列表
     
      searchObj:{},//条件查询封装对象
      addrules:{},
      dialogVisible: false,//弹出框
      dialogVisible1:false,
      sysDept:{},
      multipleSelection: [],
      username:Cookies.get("username"),
      type:Cookies.get("type"),
      sysCompany:{}, //封装添加表单数据
      options:[],
      companyList:[],
      modifyData:{},
      dialogVisible2:false,
      addrules:{
            name:[
              {
                required:true,
                message:"名称不能为空！"
              }
            ],
            companyId:[
              {
                required:true,
                message:"公司不能为空！"
              }
            ],
            leader:[
              {
                required:true,
                message:"负责人不能为空！"
              }
            ],
            createTime:[
              {
                required:true,
                message:"创建时间不能为空！"
              }
            ],
            phone:[
              {
                required:true,
                message:"电话不能为空！"
              },
              {
              pattern: /^[1][3,4,5,6,7,8,9][0-9]{9}$/,
              message: '请输入正确的手机号码'
            }
          ]
            

          },
    }
  },
  //页面渲染之前执行
  created() {
    //调用列表方法
    this.fetchData()
  },
  methods:{//具体方法
    
    // 复选框发生变化，调用方法，选中复选框行内容传递
    handleSelectionChange(selection) {
          this.multipleSelection = selection;
        },
    
   //添加弹框的方法
    addCompany() {
      this.dialogVisible = true
      this.sysCompany = {}
    },
    addDept(){
      this.dialogVisible1 = true

      api.selectCompany().then((res) => {
        if(res.message=='成功'){
          console.log("com",res)
          this.options=res.data
        }
        
      });
      this.sysDept = {}
    },
    saveCompany(){
      this.$refs["dataForm"].validate((valid)=>{
        if(valid){
          api.addCompany(this.sysCompany).then((res) => {
            if(res.message=='成功'){
              this.fetchData();
              this.$message({
                type: "success",
                message: "添加成功!",
              });
              this.dialogVisible = false;
              
            }
            else{
              this.$alert('添加失败！')
            }
          });
        }
      })
      
    },
    saveDept(){
      this.$refs["dataForm1"].validate((valid)=>{
        if(valid){
            api.checkDept(this.sysDept).then((res)=>{
              if(res.message=='成功'){
                api.addDept(this.sysDept).then((res) => {
                  if(res.message=='成功'){
                    this.fetchData();
                    this.$message({
                      type: "success",
                      message: "添加成功!",
                    });
                    this.dialogVisible1 = false;
                    
                  }
                  else{
                    this.$alert('添加失败！')
                  }
                });
              }
              else{
                this.$message(res.message)
              }
            })
            
        }
      })
      
    },
    edit(){
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要修改的记录！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能修改一条记录！");
            return;
          }
          console.log("选中",this.multipleSelection)
          this.dialogVisible2=true;
          this.modifyData=this.multipleSelection[0]
    },
    saveUpdate(){
      this.$refs["dataForm2"].validate((valid)=>{
        if(valid){
          api.update(this.modifyData).then((res)=>{
            if(res.message=='成功'){
              this.dialogVisible2=false
              this.$message({
                type: "success",
                message: "修改成功!",
              });
            }
            else{
              this.$alert('出错')
            }
          })
        }
      })
      
    },
    fetchData() {
      
      //ajax
      api.selectdeptInfo()
        .then(response => {
          //this.listLoading = false
          //console.log(response)
          //每页数据列表
          
          this.list=response.data
          console.log("sade",this.list)
          //总记录数
          //this.total = response.data.total
        })
    },
    //删除
    remove(id) {
      if (this.multipleSelection.length === 0) {
            this.$message.warning("请选择要删除的记录！");
            return;
          }
          if (this.multipleSelection.length > 1) {
            this.$message.warning("一次只能删除一条记录！");
            return;
          }
        this.$confirm('此操作将永久删除, 是否继续?', '提示', {
          confirmButtonText: '确定',
          cancelButtonText: '取消',
          type: 'warning'
        }).then(() => {
          //调用方法删除
          let id=this.multipleSelection[0].id
          api.remove(id)
            .then(response => {
              //提示
              this.$message({
                type: 'success',
                message: '删除成功!'
              });
              //刷新
              this.fetchData()
          })
        })
    },
  }
}
</script>