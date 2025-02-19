import { login, logout, getInfo } from '@/api/user'
import { getToken, setToken, removeToken } from '@/utils/auth'
import { resetRouter } from '@/router'
import Cookies from "js-cookie"

const getDefaultState = () => {
  return {
    token: getToken(),
    name: '',
    avatar: '',
    buttons: [], // 新增
    menus: '', //新增
    tablemenus:[] // tree
  }
}

const state = getDefaultState()

const mutations = {
  RESET_STATE: (state) => {
    Object.assign(state, getDefaultState())
  },
  SET_TOKEN: (state, token) => {
    state.token = token
  },
  SET_NAME: (state, name) => {
    state.name = name
  },
  SET_AVATAR: (state, avatar) => {
    state.avatar = avatar
  },
  // 新增
  SET_BUTTONS: (state, buttons) => {
    state.buttons = buttons
  },
  // 新增
  SET_MENUS: (state, menus) => {
    state.menus = menus
  },
  SET_TABLE_MENUS: (state, menus) => {
    let hasdata =  state.tablemenus.filter(item=>item.title==menus.title)
    console.log("state",state)
    if(!hasdata.length)state.tablemenus=[...state.tablemenus,menus]
    
  }
}

const actions = {
  // user login
  login({ commit }, userInfo) {
    const { username, password } = userInfo
    return new Promise((resolve, reject) => {
      login({ username: username.trim(), password: password }).then(response => {
        const { data } = response
        commit('SET_TOKEN', data.token)
        setToken(data.token)
        resolve()
      }).catch(error => {
        reject(error)
      })
    })
  },

  settablemue({ commit },playloader){
    return new Promise((resolve)=>{
      commit('SET_TABLE_MENUS',playloader)
      debugger
      resolve()
    })
  },


  // get user info
  getInfo({ commit, state }) {
    return new Promise((resolve, reject) => {
      getInfo(state.token).then(response => {
        const { data } = response

        if (!data) {
          return reject('Verification failed, please Login again.')
        }

        const { name, avatar,routers } = data
        Cookies.set('userid',name.id)
        Cookies.set('type',name.type)
        Cookies.set('companyId',name.companyId)
        commit('SET_NAME', name)
        commit('SET_AVATAR', avatar)

        commit("SET_BUTTONS", data.buttons)
        commit("SET_MENUS", data.routers)
        resolve(data)
      }).catch(error => {
        reject(error)
      })
    })
  },

  // user logout
  logout({ commit, state }) {
    return new Promise((resolve, reject) => {
      //logout(state.token).then(() => {
        removeToken() // must remove  token  first
        resetRouter()
        commit('RESET_STATE')
        resolve()
      // }).catch(error => {
      //   reject(error)
      // })
    })
  },
  // remove token
  resetToken({ commit }) {
    return new Promise(resolve => {
      removeToken() // must remove  token  first
      commit('RESET_STATE')
      resolve()
    })
  }
}

export default {
  namespaced: true,
  state,
  mutations,
  actions
}

