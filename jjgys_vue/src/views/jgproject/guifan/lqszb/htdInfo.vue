<template>
    <div class="app-container">
      <!--查询表单-->
      <div class="search-div">
        <el-form :inline="true" class="demo-form-inline">
          <el-form-item label="合同段名称:">
            <el-input v-model="searchObj.name" />
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
            @click="add()"
            style="margin-left: 0px"
            icon="el-icon-plus"
            size="mini"
            
            >添加</el-button
          >
          <!--:disabled="$hasBP('bnt.htd.add') === true"-->
          <el-button
            class="btn-add"
            type="danger"
            size="mini"
            @click="batchRemove()"
            style="margin-left: 1px"
            icon="el-icon-delete"
            
            >删除</el-button
          >
          <!-- <el-button
            class="btn-add"
            type="danger"
            size="mini"
            @click="removeAll()"
            style="margin-left: 1px"
            icon="el-icon-delete"
            >全部删除</el-button> -->
          <!--:disabled="$hasBP('bnt.htd.remove') === true"-->
  
          
  
          
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
        <el-table-column label="序号">
          <template slot-scope="scope">
            {{ (page - 1) * limit + scope.$index + 1 }}
          </template>
        </el-table-column>
        <el-table-column prop="name" label="合同段名称" />
        <el-table-column prop="sgdw" label="施工单位" />
        <el-table-column prop="jldw" label="监理单位" />
        <el-table-column prop="zhq" label="桩号起" />
        <el-table-column prop="zhz" label="桩号止" />
        <el-table-column prop="zy" label="ZY" />
        <el-table-column prop="lx" label="合同段类型" />
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
  
      <el-dialog title="添加" :visible.sync="dialogVisible" width="40%">
        <el-form
          ref="dataForm"
          :model="htd"
          label-width="100px"
          size="small"
          :rules="htdrules"
          style="padding-right: 40px"
        >
          <el-form-item label="合同段名称" prop="name">
            <el-input v-model="htd.name" />
          </el-form-item>
          <el-form-item label="施工单位" prop="sgdw">
            <el-input v-model="htd.sgdw" />
          </el-form-item>
          <el-form-item label="监理单位" prop="jldw">
            <el-input v-model="htd.jldw" />
          </el-form-item>
          <el-form-item label="桩号起" prop="zhq">
            <el-input v-model="htd.zhq" />
          </el-form-item>
          <el-form-item label="桩号止" prop="zhz">
            <el-input v-model="htd.zhz" />
          </el-form-item>
          <el-form-item label="ZY" prop="zy">
            <el-select v-model="zy">
              <el-option v-for="item in zys"
              :key="item.value"
              :label="item.label"
              :value="item.value"
              >
  
              </el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="合同段类型" prop="lx" >
            <!-- <el-input v-model="htd.lx"/> -->
            <el-select
              style="width: 100%"
              v-model="htd.lx"
              multiple
              
              allow-create
              default-first-option
              placeholder="请选择合同段类型"
            >
              <el-option
                v-for="item in options"
                :key="item.value"
                :label="item.label"
                :value="item.value"
              >
              </el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="创建时间" prop="createTime">
            <el-date-picker
              v-model="htd.createTime"
              value-format="yyyy-MM-dd HH:mm:ss"
            />
          </el-form-item>
        </el-form>
        <span slot="footer" class="dialog-footer">
          <el-button
            @click="dialogVisible = false"
            size="small"
            icon="el-icon-refresh-right"
            >取 消</el-button
          >
          <el-button
            type="primary"
            icon="el-icon-check"
            @click="save()"
            size="small"
            >确 定</el-button
          >
        </span>
      </el-dialog>
      
    </div>
  </template>
  <script>
  // 引入定义接口js文件
  import api from "@/api/jgproject/lqszb/htdInfo.js";
  import Cookies from 'js-cookie'
  export default {
    data() {
      // 初始值
      return {
        list: [], // 项目列表
        total: 0, // 总记录数
        page: 1, // 当前页码
        limit: 10, // 每页记录数
        searchObj: {}, // 查询条件
        multipleSelection: [], // 批量删除选中的记录列表
        dialogVisible: false,
        dialogImportVisible:false,
        username:Cookies.get('username'),
        userid:Cookies.get('userid'),
        type:Cookies.get('type'),
        htdrules:{
           name: [
            { required: true, message: '请输入合同段名称'}
          ],
          sgdw: [
            { required: true, message: '请输入施工单位' }
          ],
          jldw: [
            { required: true, message: '请输入监理单位' }
          ],
          zh: [
            { required: true, message: '请输入桩号' }
          ],
          lx: [
            { required: true, message: '请选择合同段类型',trigger: 'blur' }
          ],
          createTime: [
            { required: true, message: '请选择创建时间' }
          ]
        },
        htd: {
          name: '',
          sgdw: '',
          jldw: '',
          zh: '',
          createTime: '',
          lx: '',
          proname:'',
          zy:""
        },
        zy:"",
        zys:[
          {
            value: 'Z',
            label: "Z"
          },
          {
            value: "Y",
            label: "Y"
          },
          {
            value: "ZY",
            label: "ZY"
          }
        ],
        options: [
          {
            value: "路基工程",
            label: "路基工程",
          },
          {
            value: "路面工程",
            label: "路面工程",
          },
          {
            value: "交安工程",
            label: "交安工程",
          },
          {
            value: "桥梁工程",
            label: "桥梁工程",
          },
          {
            value: "隧道工程",
            label: "隧道工程",
          }
        ]
      }
    },
    created() {
      
      this.fetchData();
    },
    methods: {
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
            location.reload();
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
      
      add() {
        this.dialogVisible = true;
        this.htd = {};
      },
      save() {
        this.htd.lx = JSON.stringify(this.htd.lx)
        this.htd.zy=this.zy
        this.htd.proname = this.$route.query.projecttitle
        api.save(this.htd).then((response) => {
          this.list = response.data;
          // 提示
          this.$message({
            type: "success",
            message: "添加成功!",
          });
          // 关闭弹框
          this.dialogVisible = false;
          // 刷新
          this.fetchData();
          location.reload();
        });
      },
      
      // 检查复选框
      checkSelect(){
        if (this.multipleSelection.length === 0) {
              this.$message.warning("请选择要导入数据的合同段！");
              return;
            }
            this.dialogImportVisible=true
      },
  
       
      fetchData() {
        
        // 调用api
        this.searchObj.proname = this.$route.query.projecttitle
        this.searchObj.userid=this.userid
        api.pageList(this.page, this.limit, this.searchObj).then((response) => {
          this.list = response.data.records;
          this.total = response.data.total;
          console.log("ssq",this.list)
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
  