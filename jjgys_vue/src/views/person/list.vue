<template>
  
    <div class="app-container" >
    <el-row :gutter="20" style="margin-top: 20px">
      <el-col :span="2" style="margin-left: 20px">
        <div style="background-color:#e4e0db;width:80px;height:80px;;border-radius:50%;overflow:hidden;">
            <img src="@/assets/user/user.png" style="width:80px;height:80px;"></img>
        </div>

        <div >
              <h4 style="font-size: 25px; margin-top: 0;">{{userInfo.username}}</h4>
               
            </div>
      </el-col>

      <el-col :span="13">
        <div class="grid-content bg-purple">

          <el-card class="box-card" >
            <div slot="header" style="height:30px">
              <h4 style="font-size: 30px; margin-top: 0;">个人资料</h4>
               <el-button class="xiugai" @click="update" type="primary" round >修改</el-button>
            </div>
            <div class="personal-relation" style="background: rgb(152, 229, 89);">
              <div class="relation-item" >
                用户名:
                <div style="float: right; padding-right: 20px;">
                  {{userInfo.username}}
                </div>
              </div>
            </div>
            <div class="personal-relation" style="background: rgb(184, 210, 98);">
              <div class="relation-item" >
                姓名:
                <div style="float: right; padding-right: 20px">
                  {{userInfo.name}}
                </div>
              </div>
            </div>
          
            <div class="personal-relation" style="background: rgb(205, 217, 41);">
              <div class="relation-item">
                联系方式:
                <div style="float: right; padding-right: 20px">
                   {{userInfo.phone}}
                </div>
              </div>
            </div>
            <!--<div class="personal-relation" style="background: rgb(217, 233, 116);">
              <div class="relation-item">
                部门:
                <div style="float: right; padding-right: 20px">
                   
                </div>
              </div>
            </div>-->
          </el-card>
        </div>
      </el-col>
    </el-row>

    <el-dialog title="修改" :visible.sync="dialogVisible" width="40%" append-to-body >
      <el-form ref="dataForm" :model="userInfo" :rules="registerRules"  label-width="100px" size="small" style="padding-right: 40px;">
        <el-form-item label="用户名" prop="username">
          <el-input v-model="userInfo.username" />
        </el-form-item>
        <el-form-item label="姓名" prop="name">
          <el-input v-model="userInfo.name" />
        </el-form-item>
        <el-form-item label="联系方式" prop="phone">
          <el-input v-model="userInfo.phone"  />
        </el-form-item>
        <!--<el-form-item label="部门" prop="dept">
          ///<el-input v-model="userInfo.param.departname" /> 
          <el-select
          v-model="userInfo.deptId"
          placeholder="">
          <el-option
            v-for="dept in departList"
            :key="dept.id"
            :label="dept.name"
            :value="dept.id"/>
        </el-select>
        </el-form-item>-->
      </el-form>
      <span slot="footer" class="dialog-footer">
        <el-button @click="dialogVisible = false" size="small" icon="el-icon-refresh-right">取 消</el-button>
        <el-button type="primary" icon="el-icon-check" @click="save()" size="small">确 定</el-button>
      </span>
    </el-dialog>
  </div>
  

</template>
<script>
import useradmin from '@/api/sysuser/sysuser.js'
import Cookies from 'js-cookie'

export default {
  data() {
    return {
        userid: '',
        dialogVisible: false,
        userInfo: {},
        deptId:'',
      departList: [],
            registerRules: {
        username: [
          { required: true, message: '请输入用户名' },
          {
            min: 3,
            max: 10,
            message: '长度在 3 到 10 个字符'
          },
          {
            pattern: /^[a-zA-Z][a-zA-Z0-9]+$/,
            message: '只允许输入字母数字，只能以字母开头'
          }
        ],
        name:[
        { required: true, message: '请输入姓名', trigger: 'blur' },
        ],
        password: [
          { required: true, message: '请输入密码', trigger: 'blur' },
          {
            min: 6,
            max: 20,
            message: '长度在 6 到 20 个字符'
          }],
          phone: [
            {
              pattern: /^[1][3,4,5,6,7,8,9][0-9]{9}$/,
              message: '请输入正确的手机号码'
            }
          ],
          mail: [
            {
              pattern: /^[A-Za-z0-9\u4e00-\u9fa5]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/,
              message: '请输入正确的邮件地址'
            }
          ]
    }
    }
  },
  created() {
    // 调用列表方法
    this.fetchData()
  },
  methods: {
    initDepartList() {
      useradmin.selectdept().then(response => {
        this.departList = response.data
      })
    },
    fetchData() {
      let username = Cookies.get('username')
      // 调用api
      useradmin.getUserByName(username)
        .then((response) => {
          this.userInfo = response.data
        })
    },
    update() {
      this.dialogVisible = true
      // 获取部门列表
      this.initDepartList()
      let username = Cookies.get('username')
      useradmin.getUserByName(username)
        .then(response => {
          this.userInfo = response.data
        })
    },
    save(userInfo) {
      this.$refs["dataForm"].validate((valid)=>{
        if(valid){
          useradmin.update(this.userInfo).then(response => {
        // 提示
            this.$message.success('修改成功')
            this.dialogVisible = false
            // 刷新
            this.fetchData()
          })
        }
      })
      
    }
  }
}
</script>
<style lang="scss" scoped>
//卡片样式
h4 {
    font-weight: 600;
    padding-top: 10px;
  display:inline-block;
}

.box-card {
  margin-top: 120px;
  width: 130%;
  height: 500px;
  border-radius: 2%;
  .xiugai{
    float: right;
  }
}
//文本样式区


.personal-relation {
  font-size: 20px;
  padding: 0px 5px 15px;
  margin-right: 1px;
  width: 100%;
  margin-top: 20px;
  
  
}
.relation-item {
  padding: 12px;
}
.dialog-footer {
  padding-top: 10px;
  padding-left: 10%;
}

.bg-purple {
  //background: #0b52a3;
}
.grid-content {
  border-radius: 4px;
  min-height: 36px;
}
.app-container{
      height: 100%;
      width: 100%;
      position:fixed;
      
      
  
    }
</style>
