<template>
  <div class="login-container">


    <div class='logo' style="width:150px;height:150px;;border-radius:50%;overflow:hidden;">
      <img src="@/assets/login/logo.png" style="clear: both;display: block;margin: auto; width: 150px;height: 150px">
    </div>

    <h3 class="title-name">陕西交控工程技术有限公司</h3>

    <el-form ref="loginForm" :model="loginForm" :rules="loginRules" class="login-form" auto-complete="on"
      label-position="left">



      <div class="title-container">
        <h3 class="title">公路工程交（竣）工验收检测系统</h3>
      </div>

      <el-form-item prop="username">
        <span class="svg-container">
          <svg-icon icon-class="user" />
        </span>
        <el-input ref="username" v-model="loginForm.username" placeholder="用户名" name="username" type="text" tabindex="1"
          auto-complete="on" />
      </el-form-item>

      <el-form-item prop="password">
        <span class="svg-container">
          <svg-icon icon-class="password" />
        </span>
        <el-input :key="passwordType" ref="password" v-model="loginForm.password" :type="passwordType" placeholder="密码"
          name="password" tabindex="2" auto-complete="on" @keyup.enter.native="handleLogin" />
        <span class="show-pwd" @click="showPwd">
          <svg-icon :icon-class="passwordType === 'password' ? 'eye' : 'eye-open'" />
        </span>
      </el-form-item>

      <el-button :loading="loading" type="primary" style="width:100%;margin-bottom:30px;"
        @click.native.prevent="handleLogin">登录</el-button>
      <el-row style="text-align: right;margin-top: -10px;">
        <!--<el-link type="primary" @click="toRegister">还没有账号？立即注册</el-link>-->
      </el-row>


    </el-form>
  </div>
</template>

<script>
import { validUsername } from '@/utils/validate'
import Cookies from "js-cookie"
export default {
  name: 'Login',
  data() {
    const validateUsername = (rule, value, callback) => {
      if (value.length < 5) {
        callback(new Error('Please enter the correct user name'))
      } else {
        callback()
      }
    }
    const validatePassword = (rule, value, callback) => {
      if (value.length < 6) {
        callback(new Error('The password can not be less than 6 digits'))
      } else {
        callback()
      }
    }
    return {
      loginForm: {
        username: '',
        password: ''
      },
      loginRules: {
        userName: [
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
        pwd: [
          { required: true, message: "请输入密码", trigger: "blur" },
          {
            min: 6,
            max: 20,
            message: "长度在 8 到 20 个字符",
          },
        ],
      },
      loading: false,
      passwordType: 'password',
      redirect: undefined
    }
  },
  watch: {
    $route: {
      handler: function (route) {
        this.redirect = route.query && route.query.redirect
      },
      immediate: true
    }
  },
  methods: {
    showPwd() {
      if (this.passwordType === 'password') {
        this.passwordType = ''
      } else {
        this.passwordType = 'password'
      }
      this.$nextTick(() => {
        this.$refs.password.focus()
      })
    },
    handleLogin() {
      // this.$refs.loginForm.validate在上面的表格中已经绑定过，这是element-UI组件自带的函数，可以检验表单的有效性
      this.$refs.loginForm.validate(valid => {
        if (valid) {
          Cookies.set('username', this.loginForm.username)
          this.loading = true
          this.$store.dispatch('user/login', this.loginForm) // 进行一个动作，是ueser.js里面的login函数
            .then(() => {
              this.$router.push({ path: '/' }) //这个才是跳转页面
              this.loading = false
            })
            .catch(() => {
              this.loading = false
            })
        } else {
          console.log('error submit!!')
          return false
        }
      })
    },
    toRegister: function () {
      this.$router.push('/register')
    }
  }
}
</script>

<style lang="scss">
/* 修复input 背景不协调 和光标变色 */
/* Detail see https://github.com/PanJiaChen/vue-element-admin/pull/927 */

$bg: #283443;
$light_gray: #fff;
$cursor: #fff;
$shadow: #636e7d;

@supports (-webkit-mask: none) and (not (cater-color: $cursor)) {
  .login-container .el-input input {
    color: $cursor;
  }
}

/* reset element-ui css */
.login-container {
  input:-webkit-autofill {
    -webkit-box-shadow: 0 0 0px 1000px transparent inset;
    transition: background-color 5000s ease-in-out 0s;
    -webkit-text-fill-color: #fff !important
  }

  .el-input {
    display: inline-block;
    height: 47px;
    width: 85%;

    input {
      background: transparent;
      border: 0px;
      -webkit-appearance: none;
      border-radius: 0px;
      padding: 12px 5px 12px 15px;
      color: $light_gray;
      height: 47px;
      caret-color: $cursor;



      /*&:-webkit-autofill {
         box-shadow: 0 0 0px 1000px rgba(0, 0, 0, 0.1) inset !important;
        //-webkit-text-fill-color:  #fff !important;
      };*/

    }
  }

  .el-form-item {
    border: 1px solid rgba(255, 255, 255, 0.1);
    background: rgba(0, 0, 0, 0.1);
    border-radius: 5px;
    color: #454545;
  }
}
</style>

<style lang="scss" scoped>
$bg: #348091;
$dark_gray: #889aa4;
$light_gray: #eee;

.login-container {
  min-height: 100%;
  width: 100%;
  background-color: $bg;
  overflow: hidden;

  .login-form {
    position: relative;
    width: 520px;
    max-width: 100%;
    //padding: 160px 35px 0;
    margin: 0 auto;
    overflow: hidden;
  }

  .tips {
    font-size: 14px;
    color: #fff;
    margin-bottom: 10px;

    span {
      &:first-of-type {
        margin-right: 16px;
      }
    }
  }

  .svg-container {
    padding: 6px 5px 6px 15px;
    color: $dark_gray;
    vertical-align: middle;
    width: 30px;
    display: inline-block;
  }

  .title-container {
    position: relative;

    .title {
      font-size: 26px;
      color: $light_gray;
      margin: 0px auto 40px auto;
      text-align: center;
      font-weight: bold;
    }
  }

  .show-pwd {
    position: absolute;
    right: 10px;
    top: 7px;
    font-size: 16px;
    color: $dark_gray;
    cursor: pointer;
    user-select: none;
  }
}

.title-name {
  font-size: 26px;
  color: rgb(15, 15, 17);
  margin: 0px auto 40px auto;
  text-align: center;
  font-weight: bold;
}

.logo {
  position: relative;
  width: 520px;
  max-width: 100%;
  margin: 0 auto;
  margin-top: 50px;
  margin-bottom: 10px;
}
</style>
