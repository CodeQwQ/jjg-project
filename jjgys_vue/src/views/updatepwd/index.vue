<template>
  <div class="main" style="background-color: rgb(159, 196, 236)">
    <el-form
    :model="ruleForm"
    status-icon
    :rules="rules"
    ref="ruleForm"
    label-width="100px"
    class="demo-ruleForm"
  >
    <el-form-item label="旧密码" prop="oldpass">
      <el-input
        type="password"
        v-model="ruleForm.oldpass"
        autocomplete="off"
      ></el-input>
    </el-form-item>
    <el-form-item label="新密码" prop="newpass">
      <el-input
        type="password"
        v-model="ruleForm.newpass"
        autocomplete="off"
      ></el-input>
    </el-form-item>
    <el-form-item label="确认密码" prop="checkpass">
      <el-input
        type="password"
        v-model="ruleForm.checkpass"
        autocomplete="off"
      ></el-input>
    </el-form-item>
    <el-form-item>
      <el-button type="primary" @click="submitForm('ruleForm')">提交</el-button>
      <el-button @click="resetForm('ruleForm')">重置</el-button>
    </el-form-item>
  </el-form>
  </div>

</template>
<script>
import useradmin from '@/api/sysuser/sysuser.js'
import Cookies from 'js-cookie'
export default {
  data() {
    var validateOldPass = (rule, value, callback) => {
      if (value === "") {
        callback(new Error("请输入旧密码"));
      } 
        //调用后台接口查询旧密码是否正确
        this.checkIsRepeat(
          value,
          function () {
            callback();
          },function (err) {
            callback();
          }
        );
    };
    var validatePass = (rule, value, callback) => {
      if (value === "") {
        callback(new Error("请输入密码"));
      } else {
        if (this.ruleForm.checkpass !== "") {
          this.$refs.ruleForm.validateField("checkpass");
        }
        callback();
      }
    };
    var validatePass2 = (rule, value, callback) => {
      if (value === "") {
        callback(new Error("请再次输入密码"));
      } else if (value !== this.ruleForm.newpass) {
        callback(new Error("两次输入密码不一致!"));
      } else {
        callback();
      }
    };
    return {
      ruleForm: {
        oldpass: "",
        newpass: "",
        checkpass: "",
      },
      checkpwd:false,
      rules: {
        oldpass: [{ validator: validateOldPass, trigger: "blur" }],
        newpass: [{ validator: validatePass, trigger: "blur" },
         {
            min: 6,
            max: 20,
            message: '长度在 6 到 20 个字符'
          }],
        checkpass: [{ validator: validatePass2, trigger: "blur" },
         {
            min: 6,
            max: 20,
            message: '长度在 6 到 20 个字符'
          }],
      },
    };
  },
  methods: {
    //校验旧密码
    checkIsRepeat(oldpw, success, error) {
      //checkIsRepeat(oldpw) {
      let username = Cookies.get('username')
      if (!oldpw) return;
      useradmin.checkSysUserPwd(oldpw,username).then((response) => {
          if(response.status == 2000){
            success()
            this.pwd==true
          }else{
            error()
          }
        })

    },

    //修改
    submitForm() {
      this.$refs["ruleForm"].validate((valid)=>{
        if(valid){
          let username = Cookies.get('username')
          useradmin.updatepwd(username,this.ruleForm.newpass)
            .then(response => {
              //提示
              this.$message.success('修改成功')
            setTimeout(() => {
                      this.$store.dispatch('user/logout')
            this.$router.push(`/login?redirect=${this.$route.fullPath}`)
                    }, 2000);  
            
              
            })
        }
        else{
          return false;
        }
      })
      
      // var formData = new FormData();
      // formData.append("username", username);
      // formData.append("newpass", this.ruleForm.newpass);
      
      
    },

    submitForm1() {
      let username = Cookies.get('username')
      var formData = new FormData();
      formData.append("username", username);
      formData.append("newpass", this.ruleForm.newpass);
      //this.$refs.ruleForm.validate((valid) => {
        this.$refs["ruleForm"].validate(valid => {
          if (valid) {
            useradmin.updatepwd(formData).then(response=>{
              this.$message.success('修改成功');
            });
          }
        })
    },
    resetForm(formName) {
      this.$refs[formName].resetFields();
    },
  },
};
</script>
<style lang="scss">
.main{
  height: auto;
}
.demo-ruleForm {
  width: 80%;
  padding: 100px 200px 35px 250px;
}
</style>