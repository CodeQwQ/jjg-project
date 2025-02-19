package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.common.utils.MD5;
import glgc.jjgys.model.system.SysDept;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.system.SysUserMenu;
import glgc.jjgys.model.system.SysUserRole;
import glgc.jjgys.model.vo.SysUserQueryVo;
import glgc.jjgys.system.mapper.SysRoleMenuMapper;
import glgc.jjgys.system.mapper.SysUserRoleMapper;
import glgc.jjgys.system.service.SysDepartService;
import glgc.jjgys.system.service.SysUserMenuService;
import glgc.jjgys.system.service.SysUserService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * <p>
 * 用户表 前端控制器
 * </p>
 */
@Api(tags = "用户管理接口")
@RestController
@RequestMapping("/admin/system/sysUser")
public class SysUserController {

    @Autowired
    private SysUserService sysUserService;
    @Autowired
    private SysDepartService sysDepartService;
    @Autowired
    private SysUserMenuService sysUserMenuService;

    @ApiOperation("校验旧密码")
    @GetMapping("checkSysUserPwd/{oldpw}/{username}")
    public Result checkSysUserPwd(@PathVariable String oldpw,
                                  @PathVariable String username) {
        boolean pwd = sysUserService.checkSysUserPwd(oldpw,username);
        if (pwd){
            return Result.ok(null).message("旧密码校验成功");
        } else {
            return Result.fail(null).message("旧密码校验失败");
        }

    }

    @ApiOperation("用户修改密码")
    @PostMapping("updatepwd/{username}/{newpass}")
    public Result updatepwd(@PathVariable String username,
                            @PathVariable String newpass) {
        boolean is_Success = sysUserService.updatepwd(username,newpass);
        if(is_Success) {
            return Result.ok(null);
        } else {
            return Result.fail(null);
        }
    }

    @ApiOperation("根据用户名查询")
    @GetMapping("getUserByName/{username}")
    public Result getUserByName(@PathVariable String username) {
        SysUser user = sysUserService.getUserInfoByUserName(username);
        SysDept sysDept = sysDepartService.getById(user.getDeptId());
        if (sysDept != null){
            String name = sysDept.getName();
            Long parentId = sysDept.getParentId();
            SysDept byId = sysDepartService.getById(parentId);
            String name1 = byId.getName();
            Map<String,Object> map=new HashMap<>();
            map.put("departname",name+","+name1);
            user.setParam(map);
            return Result.ok(user);
        }
        return Result.ok(user);

    }

    @ApiOperation("用户列表")
    @GetMapping("/{page}/{limit}")
    public Result list(@PathVariable Long page,
                       @PathVariable Long limit,
                       SysUserQueryVo sysUserQueryVo) {
        //创建page对象
        Page<SysUser> pageParam = new Page<>(page,limit);
        //调用service方法

        IPage<SysUser> pageModel = sysUserService.selectPage(pageParam,sysUserQueryVo);
        // 在查询结果中过滤用户名为 "admin" 的用户
        List<SysUser> filteredUsers = pageModel.getRecords().stream()
                .filter(user -> !"admin".equals(user.getUsername()))
                .collect(Collectors.toList());
        pageModel.setRecords(filteredUsers);

        List<SysUser> userList = pageModel.getRecords();
        List<SysUser> modifiedUserList = new ArrayList<>();

        // 遍历分页中的记录
        for (SysUser user : userList) {
            // 创建一个新的用户对象，避免直接修改原始对象
            SysUser modifiedUser = new SysUser();
            BeanUtils.copyProperties(user, modifiedUser); // Assuming you have a method like copyProperties to copy the properties

            // 处理每个用户对象
            String type = user.getType();
            if ("2".equals(type) ) {
                modifiedUser.setType("项目负责人");
            } else if ("3".equals(type)) {
                modifiedUser.setType("普通用户");
            }else if ("4".equals(type)) {
                modifiedUser.setType("公司管理员");
            }

            Long deptId = user.getDeptId();
            QueryWrapper<SysDept> deptQueryWrapper = new QueryWrapper<>();
            deptQueryWrapper.eq("id", deptId);
            List<SysDept> list = sysDepartService.list(deptQueryWrapper);

            if (!list.isEmpty()) {
                modifiedUser.setDeptName(list.get(0).getName());
            }

            modifiedUserList.add(modifiedUser);
        }

        pageModel.setRecords(modifiedUserList);
        System.out.println(pageModel+"sss");
        return Result.ok(pageModel);
    }

    /**
     *
     * @param request
     * @return
     */
    private SysUser getUserInfoByRequest(HttpServletRequest request){
        //获取请求头token字符串,因为把token放到请求头不存在跨域
        String token = request.getHeader("token");
        //从token字符串获取用户名称（id）
        String userId = JwtHelper.getUserId(token);
        //根据用户名称获取用户信息
        return sysUserService.getById(userId);
    }

    @ApiOperation("用户列表")
    @PostMapping("/{page}/{limit}")
    public Result listByPost(@PathVariable Long page,
                             @PathVariable Long limit,
                             @RequestBody  SysUserQueryVo sysUserQueryVo, HttpServletRequest request) {
        //创建page对象
        Page<SysUser> pageParam = new Page<>(page,limit);
        //调用service方法

        SysUser userInfoByRequest = getUserInfoByRequest(request);
        if (!"1".equals(userInfoByRequest.getType())){
            sysUserQueryVo.setCompanyId(userInfoByRequest.getCompanyId());
        }
        IPage<SysUser> pageModel = sysUserService.selectPage(pageParam,sysUserQueryVo);
        // 在查询结果中过滤用户名为 "admin" 的用户
        List<SysUser> filteredUsers = pageModel.getRecords().stream()
                .filter(user -> !"admin".equals(user.getUsername()))
                .collect(Collectors.toList());
        pageModel.setRecords(filteredUsers);
        List<SysUser> userList = pageModel.getRecords();
        // 遍历分页中的记录
        for (SysUser user : userList) {
            // 处理每个用户对象
            String type = user.getType();
            LocalDateTime expire = user.getExpire();

            if ("2".equals(type)){
                user.setType("项目负责人");
            } else if ("3".equals(type)) {
                user.setType("普通用户");
            } else if ("4".equals(type)) {
                user.setType("公司管理员");
            }
            if (String.valueOf(expire).equals("9999-12-31T23:59:59")){
                user.setDescription("永久期限");
            }else {
                user.setDescription(expire.toString());
            }
            Long deptId = user.getDeptId();
            QueryWrapper<SysDept> deptQueryWrapper = new QueryWrapper<>();
            deptQueryWrapper.eq("id",deptId);
            List<SysDept> list = sysDepartService.list(deptQueryWrapper);
            user.setDeptName(list.get(0).getName());
        }
        pageModel.setRecords(userList);
        return Result.ok(pageModel);
    }

    @ApiOperation("添加用户")
    @PostMapping("save")
    public Result save(@RequestBody SysUser user) {
        //把输入密码进行加密 MD5
        String encrypt = MD5.encrypt(user.getPassword());
        user.setPassword(encrypt);
        boolean is_Success = sysUserService.save(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }



    @ApiOperation("根据id查询")
    @GetMapping("getUser/{id}")
    public Result getUser(@PathVariable String id) {
        SysUser user = sysUserService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改用户")
    @PostMapping("update")
    public Result update(@RequestBody SysUser user) {
        boolean is_Success = sysUserService.updateById(user);

        return is_Success ? Result.ok(): Result.fail();
    }

    @Autowired
    private SysUserRoleMapper sysUserRoleMapper;

    @Autowired
    private SysRoleMenuMapper sysRoleMenuMapper;

    @ApiOperation("删除用户")
    @DeleteMapping("remove/{id}")
    public Result remove(@PathVariable String id) {
        //删除菜单权限信息 sys_user_role sys_role_menu
        /*QueryWrapper<SysUserRole> userrolewrapper1 = new QueryWrapper<>();
        userrolewrapper1.eq("user_id",id);
        List<SysUserRole> sysUserRoles = sysUserRoleMapper.selectList(userrolewrapper1);
        if (sysUserRoles != null){
            for (SysUserRole sysUserRole : sysUserRoles) {
                String roleId = sysUserRole.getRoleId();
                QueryWrapper<SysRoleMenu> wrapper = new QueryWrapper<>();
                wrapper.eq("role_id",roleId);
                sysRoleMenuMapper.delete(wrapper);
            }
        }*/
        QueryWrapper<SysUserRole> userrolewrapper = new QueryWrapper<>();
        userrolewrapper.eq("user_id",id);
        sysUserRoleMapper.delete(userrolewrapper);

        //删除user_user_menu
        QueryWrapper<SysUserMenu> userMenuQueryWrapper = new QueryWrapper<>();
        userMenuQueryWrapper.eq("user_id",id);
        boolean remove = sysUserMenuService.remove(userMenuQueryWrapper);

        boolean b = sysUserService.removeById(id);

        return b ? Result.ok(): Result.fail();
    }
}

