<template>
  <div class="app-container">
    <div class="search-div">
      <el-form label-width="70px" size="small">
        <el-row>
          <el-col>
            <el-form-item label="关 键 字">
              <el-input
                style="width: 20%"
                v-model="searchObj.keyword"
                placeholder="用户名/姓名/手机号码"
              ></el-input>
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
            </el-form-item>
          </el-col>
        </el-row>
        <el-row style="display: flex">
          <!-- <el-button type="primary" icon="el-icon-search" size="mini"  @click="fetchData()">搜索</el-button>
          <el-button icon="el-icon-refresh" size="mini" @click="resetData">重置</el-button> -->
        </el-row>
      </el-form>
    </div>

    <!-- 工具条 -->
    <div class="tools-div">
      <el-button
        type="success"
        v-if="username=='admin'&&type==1"
        icon="el-icon-plus"
        size="mini"
        @click="add"
        >添 加</el-button
      >
    </div>

    <el-container>
      <el-aside :style="{ width: '20%' }">
        <!-- 公司部门树-->
        <el-tree
          style="width: 100%; margin-top: 10px"
          :data="treeData"
          default-expand-all
          :props="treeProps"
          @node-click="handleNodeClick"
        ></el-tree>
      </el-aside>
      <el-main>
        <!-- 列表 -->
        <el-table
          v-loading="listLoading"
          :data="list"
          stripe
          border
          style="width: 100%; margin-top: 10px"
        >
          <el-table-column label="序号" width="70" align="center">
            <template slot-scope="scope">
              {{ (page - 1) * limit + scope.$index + 1 }}
            </template>
          </el-table-column>

          <el-table-column prop="username" label="用户名" width="110" />
          <el-table-column prop="name" label="姓名" width="110" />
          <el-table-column prop="phone" label="手机" />
          <el-table-column prop="type" label="用户类型" />
          <el-table-column prop="deptName" label="部门名称" />
          <el-table-column prop="description" label="有效期" />
          <el-table-column prop="createTime" label="创建时间" />

          <el-table-column label="操作" align="center" fixed="right" width="200">
            <template slot-scope="scope">
              <el-button
                type="primary"
                icon="el-icon-edit"
                size="mini"
                v-if="username=='admin'&&type==1"
                @click="edit(scope.row.id)"
                title="修改"
              />
              <el-button
                type="danger"
                icon="el-icon-delete"
                size="mini"
                v-if="username=='admin'&&type==1"
                @click="removeDataById(scope.row.id)"
                title="删除"
              />
              <!--<el-button
                type="warning"
                icon="el-icon-baseball"
                size="mini"
                @click="showAssignRole(scope.row)"
                title="分配角色"
              />-->
              <el-button
                type="warning"
                icon="el-icon-baseball"
                size="mini"
                @click="showAssignAuth(scope.row)"
                title="分配权限"
                :style="{ display: ((scope.row.username==username) || ( username !== 'admin' && scope.row.type == '公司管理员')) ? 'none':''}"
              
              />
            </template>
          </el-table-column>
        </el-table>

        <!-- 分页组件 -->
        <el-pagination
          :current-page="page"
          :total="total"
          :page-size="limit"
          style="padding: 30px 0; text-align: center"
          layout="total, prev, pager, next, jumper"
          @current-change="fetchData"
        />
      </el-main>
    </el-container>

    <el-dialog title="分配角色" :visible.sync="dialogRoleVisible">
      <el-form label-width="80px">
        <el-form-item label="用户名">
          <el-input disabled :value="sysUser.username"></el-input>
        </el-form-item>

        <el-form-item label="角色列表">
          <el-checkbox
            :indeterminate="isIndeterminate"
            v-model="checkAll"
            @change="handleCheckAllChange"
            >全选</el-checkbox
          >
          <div style="margin: 15px 0"></div>
          <el-checkbox-group
            v-model="userRoleIds"
            @change="handleCheckedChange"
          >
            <el-checkbox
              v-for="role in allRoles"
              :key="role.id"
              :label="role.id"
              >{{ role.roleName }}</el-checkbox
            >
          </el-checkbox-group>
        </el-form-item>
      </el-form>
      <div slot="footer">
        <el-button type="primary" @click="assignRole" size="small"
          >保存</el-button
        >
        <el-button @click="dialogRoleVisible = false" size="small"
          >取消</el-button
        >
      </div>
    </el-dialog>
    <el-dialog title="分配权限" :visible.sync="dialogAuthorityVisible">
      <el-tree
        ref="menuTree"
        :data="menuTreeData"
        show-checkbox
        default-expand-all
        node-key="id"
        highlight-current
        :props="defaultPropsTree"
      />

      <div slot="footer">
        <el-button type="primary" @click="assignAuthority" size="small"
          >保存</el-button
        >
        <el-button @click="dialogAuthorityVisible = false" size="small"
          >取消</el-button
        >
      </div>
    </el-dialog>

    <el-dialog title="选择部门" :visible.sync="dialogVisibleDept" width="40%">
      <el-tree
        ref="selcetDept"
        :data="selectDeptArr"
        default-expand-all
        node-key="id"
        highlight-current
        :props="defaultPropsTree"
        @node-click="handleNodeClickSelectDept"
      />

      <div slot="footer">
        <el-button
          type="primary"
          @click="dialogVisibleDept = false"
          size="small"
          >确认</el-button
        >
      </div>
    </el-dialog>
    <el-dialog title="添加" :visible.sync="dialogVisible" width="40%">
      
      <el-form
        ref="dataForm"
        :model="sysUser"
        :rules="addrulse"
        label-width="100px"
        size="small"
        style="padding-right: 40px"
      >
        <el-form-item label="公司及部门" prop="companyName" >
          <el-input
            placeholder="请选择部门"
            v-model="sysUser.companyName"
            class="input-with-select"
            readonly
          >
            <el-button
              slot="append"
              icon="el-icon-search"
              @click="selectDept"
            >请选择所属公司及部门</el-button>

          </el-input>
        </el-form-item>

        <el-form-item label="用户名" prop="username">
          <el-input v-model="sysUser.username" />
        </el-form-item>
        <el-form-item v-if="!sysUser.id" label="密码" prop="password">
          <el-input v-model="sysUser.password" type="password" />
        </el-form-item>
        <el-form-item label="姓名" prop="name">
          <el-input v-model="sysUser.name" />
        </el-form-item>
        <el-form-item label="手机" prop="phone">
          <el-input v-model="sysUser.phone" />
        </el-form-item>
        <el-form-item label="人员类型" prop="type">
          <el-select
            style="width: 100%"
            v-model="sysUser.type"
            default-first-option
            placeholder="请选择人员类型"
          >
            <el-option
              v-for="item in types"
              :key="item.value"
              :label="item.label"
              :value="item.value"
            >
            </el-option>
          </el-select>
        </el-form-item>
        <el-form-item label="有效期类型" prop="check">
          <el-select
            style="width: 100%"
            v-model="sysUser.check"
            default-first-option
            placeholder="请选择有效期类型"
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
        <el-form-item v-if="sysUser.check == 1" label="选择时间" prop="expire">
          <el-date-picker
            v-model="sysUser.expire"
            type="datetime"
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
          @click="saveOrUpdate()"
          size="small"
          >确 定</el-button
        >
      </span>
    </el-dialog>
  </div>
</template>
<script>
import api from "@/api/system/user";
import roleApi from "@/api/system/role";
import deptApi from "@/api/system/dept";
import menuApi from "@/api/system/menu";
import userMenuApi from "@/api/system/userMenu";
import Cookies from 'js-cookie'
export default {
  data() {
    var validateUsername = (rule, value, callback) => {
      if (value === "admin") {
        callback(new Error("用户名不能为admin"));
      }
    };
    return {
      defaultPropsTree: {
        children: "children",
        label: "name",
      },
      menuTreeData: [],
      treeData: [],
      treeProps: {
        children: "children",
        label: "name",
      },
      username:Cookies.get("username"),
      type:Cookies.get("type"),
      listLoading: false, // 数据是否正在加载
      list: null, // 列表
      total: 0, // 数据库中的总记录数
      page: 1, // 默认页码
      limit: 10, // 每页记录数
      searchObj: {}, // 查询表单对象

      createTimes: [],

      dialogVisible: false,
      sysUser: {},
      options: [
        { label: "期限", value: 1 },
        { label: "永久", value: 2 },
      ],
      types:[
        { label: "普通用户", value: 3 },
        { label: "项目负责人", value: 2 },
        { label: "公司管理员", value: 4 },
      ],
      dialogRoleVisible: false,
      dialogAuthorityVisible: false,
      allRoles: [], // 所有角色列表
      userRoleIds: [], // 用户的角色ID的列表
      isIndeterminate: false, // 是否是不确定的
      checkAll: false, // 是否全选
      addrulse: {
        companyName:[
          { required: true, message: "请选择所属公司及部门", trigger: 'blur' },
        ],
        username: [
          { required: true, message: "请输入用户名" },
          {
            min: 3,
            max: 10,
            message: "长度在 3 到 10 个字符",
          },
          {
            pattern: /^[a-zA-Z][a-zA-Z0-9]+$/,
            message: "只允许输入字母数字，只能以字母开头",
          },
        ],
        password: [
          { required: true, message: "请输入密码", trigger: "blur" },
          {
            min: 6,
            max: 20,
            message: "长度在 6 到 20 个字符",
          },
        ],
        name: [{ required: true, message: "姓名不能为空！" }],
        phone: [
          { required: true, message: "手机号不能为空！" },
          {
            pattern: /^[1][3,4,5,6,7,8,9][0-9]{9}$/,
            message: "请输入正确的手机号码",
          },
        ],
        type:[
          { required: true, message: "请选择人员类型" },
        ],
        check:[
          { required: true, message: "请选择有效期类型" },
        ],
        expire:[
          { required: true, message: "请选择到期时间" },
        ],
      },
      userCompanyId: "",
      dialogVisibleDept: false,
      selectDeptArr: [],
    
    };
  },
  created() {
    //调用列表方法
    this.getCompanyTree();
    this.fetchDataByPost();
    this.getMenuTree();
  },
  methods: {
    getMenuTree() {
      // 查询菜单
      let allNode = [];
      menuApi.findNodesByAuthUser().then((response) => {
        this.menuTreeData = response.data;
        if(this.type==2){
          let idx=this.menuTreeData.findIndex(item=>{
            if(item.name=="系统管理"){
              return true
            }
          })
          this.menuTreeData.splice(idx,1)
        }
        else if(this.type==1){
          let idx=this.menuTreeData.findIndex(item=>{
            if(item.name=="系统管理"){
              return true
            }
          })
          let idx2=this.menuTreeData[idx].children.findIndex(item=>{
            if(item.name=="部门管理"){
              return true
            }
          })
          
          this.menuTreeData[idx].children.splice(idx2,1)
          console.log("sdedsd",this.menuTreeData)
        }
        
        console.log("树",this.menuTreeData)
      });
    },
    // 递归查找父节点并组成数组
    findParentNodes(tree, targetNodeId, parentNodes = []) {
      console.log(tree, targetNodeId, parentNodes);
      for (const node of tree) {
        if (node.id === targetNodeId) {
          return [...parentNodes, node]; // 将目标节点的父节点和父节点的父节点组成数组
        }

        if (node.children && node.children.length > 0) {
          const foundNodes = this.findParentNodes(node.children, targetNodeId, [
            ...parentNodes,
            node,
          ]);
          if (foundNodes.length > 0) {
            return foundNodes; // 返回找到的节点数组
          }
        }
      }
      return []; // 没有找到目标节点的父节点和父节点的父节点
    },

    assignAuthority() {
      const treeInstance = this.$refs.menuTree;
      const checkedNodes = treeInstance.getCheckedNodes();
      const parentNodes = [];

      let selectDataNode = this.$refs.menuTree.getCheckedNodes();
      console.log(selectDataNode);
      // 保存新的菜单
      if (!(selectDataNode && selectDataNode.length > 0)) {
        this.$message.warning("未选择数据");
        return;
      }
      let parentIds = new Set();
      for (const node of selectDataNode) {
        parentIds.add(node.parentId);
      }
      for (const node of parentIds) {
        let parentNode = this.findParentNodes(this.menuTreeData, node);
        if (parentNode && parentNode.length > 0) {
          parentNode.forEach((el) => {
            let flag = true;
            for (let index = 0; index < parentNodes.length; index++) {
              const element = parentNodes[index];
              if (element.id == el.id) {
                flag = false;
                break;
              }
            }
            for (let index = 0; index < selectDataNode.length; index++) {
              const element = selectDataNode[index];
              if (element.id == el.id) {
                flag = false;
                break;
              }
            }

            if (flag) {
              parentNodes.push(el);
            }
          });
        }
      }
      selectDataNode.push(...parentNodes);
      let addUserMenu = [];
      selectDataNode.forEach((menu) => {
        addUserMenu.push({
          userId: this.userId,
          menuId: menu.id,
        });
      });
      userMenuApi.saveUserMenus({ data: addUserMenu }).then((response) => {
        console.log(response.data);
        this.$message.success("保存成功");
        treeInstance.setCheckedKeys([]);

        this.dialogAuthorityVisible = false;
      });
    },

    selectDept() {
      this.dialogVisibleDept = true;

      this.selectDeptArr = JSON.parse(JSON.stringify(this.treeData));
      this.selectDeptArr.forEach((ss) => {
        ss.disabled = true;
      });
    },
    handleNodeClickSelectDept(data, node) {
      if (node.level == 1) {
        this.$message.warning("请选择部门");
        return;
      }
      this.sysUser.companyId = data.companyId;
      this.sysUser.companyName = data.name;
      this.sysUser.deptId = data.id;
    },
    // 公司部门树
    getCompanyTree() {
      deptApi.selectdeptTree().then((response) => {
        this.treeData = response.data;
        if (this.treeData && this.treeData.length > 0) {
          this.searchObj.companyId = this.treeData[0].companyId;
        }
      });
    },
    // 树点击事件
    handleNodeClick(data, node, tree) {
      console.log(data, node, tree);
      // 查询用户
      // 一级 根据公司查询用户
      this.searchObj = {};
      if (node.level == 1) {
        this.searchObj.companyId = data.companyId;
      } else if (node.level == 2) {
        this.searchObj.deptId = data.id;
      }
      this.fetchDataByPost();
    },
    showAssignAuth(row) {
      console.log(row, "Sssss");
      this.dialogAuthorityVisible = true;
      this.userId = row.id;
      // 根据用户查询菜单
      let selectNode = [];

      menuApi.findNodesByUserId(row.id).then((response) => {
        selectNode = response.data;
        console.log("tree",response.data);
        let nodeKeys = [];
        selectNode.forEach((menu) => {
          if (menu.type == 2) {
            nodeKeys.push(menu.id);
          } else if (menu.type == 1) {
            let flag = true
            selectNode.forEach((menuT) => {
              if (menu.id  == menuT.parentId) {
                flag = false
              }
            });
            let node = this.$refs.menuTree.getNode(menu.id);
            if (flag && !(node.children && node.children.length > 0)) {
              nodeKeys.push(menu.id);
            }
          }
        });
        console.log(nodeKeys+"wq");

        this.$refs.menuTree.setCheckedKeys(nodeKeys);
      });
    },
    //展示分配角色
    showAssignRole(row) {
      this.sysUser = row;
      this.dialogRoleVisible = true;
      roleApi.getRolesByUserId(row.id).then((response) => {
        this.allRoles = response.data.allRoles;
        console.log(this.allRoles);
        this.userRoleIds = response.data.userRoleIds;
        console.log(this.userRoleIds);
        this.checkAll = this.userRoleIds.length === this.allRoles.length;
        this.isIndeterminate =
          this.userRoleIds.length > 0 &&
          this.userRoleIds.length < this.allRoles.length;
      });
    },

    /*
    全选勾选状态发生改变的监听
    */
    handleCheckAllChange(value) {
      // value 当前勾选状态true/false
      // 如果当前全选, userRoleIds就是所有角色id的数组, 否则是空数组
      this.userRoleIds = value ? this.allRoles.map((item) => item.id) : [];
      // 如果当前不是全选也不全不选时, 指定为false
      this.isIndeterminate = false;
    },

    /*
    角色列表选中项发生改变的监听
    */
    handleCheckedChange(value) {
      const { userRoleIds, allRoles } = this;
      this.checkAll =
        userRoleIds.length === allRoles.length && allRoles.length > 0;
      this.isIndeterminate =
        userRoleIds.length > 0 && userRoleIds.length < allRoles.length;
    },

    //分配角色
    assignRole() {
      let assginRoleVo = {
        userId: this.sysUser.id,
        roleIdList: this.userRoleIds,
      };
      roleApi.assignRoles(assginRoleVo).then((response) => {
        this.$message.success(response.message || "分配角色成功");
        this.dialogRoleVisible = false;
        this.fetchData(this.page);
      });
    },
    //更改用户状态
    switchStatus(row) {
      //判断，如果当前用户可用，修改禁用
      row.status = row.status === 1 ? 0 : 1;
      api.updateStatus(row.id, row.status).then((response) => {
        this.$message.success(response.message || "操作成功");
        this.fetchData();
      });
    },
    //删除
    removeDataById(id) {
      this.$confirm("此操作将永久删除该用户, 是否继续?", "提示", {
        confirmButtonText: "确定",
        cancelButtonText: "取消",
        type: "warning",
      }).then(() => {
        //调用方法删除
        api.removeById(id).then((response) => {
          //提示
          this.$message({
            type: "success",
            message: "删除成功!",
          });
          //刷新
          this.fetchData();
        });
      });
    },
    //根据id查询，数据回显
    edit(id) {
      //弹出框
      this.dialogVisible = true;
      //调用接口查询
      api.getUserId(id).then((response) => {
        this.sysUser = response.data;
      });
    },
    //添加或者修改方法
    saveOrUpdate() {
      this.$refs["dataForm"].validate((valid) => {
        if (valid) {
          if (this.sysUser.username == "admin") {
            this.$message.error("用户名不能为admin")
            return
          }
          if (!this.sysUser.id) {
            this.save();
          } else {
            this.update();
          }
        }
      });
    },
    //修改
    update() {
      api.update(this.sysUser).then((response) => {
        //提示
        this.$message.success("操作成功");
        //关闭弹框
        this.dialogVisible = false;
        //刷新
        this.fetchData();
      });
    },
    //添加
    save() {
      if (this.sysUser.check == 2) {
        this.sysUser.expire = "9999-12-31 23:59:59";
      }
      console.log("time", this.sysUser.expire);
      api.save(this.sysUser).then((response) => {
        //提示
        this.$message.success("操作成功");
        //关闭弹框
        this.dialogVisible = false;
        //刷新
        this.fetchData();
      });
    },
    //添加弹框的方法
    add() {
      this.dialogVisible = true;
      this.sysUser = {};
    },
    // 重置查询表单
    resetData() {
      console.log("重置查询表单");
      this.searchObj = {};
      this.createTimes = [];
      this.fetchData();
    },
    //列表
    fetchData(page = 1) {
      this.page = page;
      if (this.createTimes && this.createTimes.length == 2) {
        this.searchObj.createTimeBegin = this.createTimes[0];
        this.searchObj.createTimeEnd = this.createTimes[1];
      }
      this.searchObj.companyId = this.userCompanyId;
      api
        .getPageList(this.page, this.limit, this.searchObj)
        .then((response) => {
          this.list = response.data.records;
          console.log("ssss", this.list);
          this.total = response.data.total;
        });
    },
    fetchDataByPost(page = 1) {
      this.page = page;
      if (this.createTimes && this.createTimes.length == 2) {
        this.searchObj.createTimeBegin = this.createTimes[0];
        this.searchObj.createTimeEnd = this.createTimes[1];
      }
      api
        .getPageListByPost(this.page, this.limit, this.searchObj)
        .then((response) => {
          this.list = response.data.records;
          console.log("ssss", this.list);
          this.total = response.data.total;
        });
    },
  },
 
  
};
</script>