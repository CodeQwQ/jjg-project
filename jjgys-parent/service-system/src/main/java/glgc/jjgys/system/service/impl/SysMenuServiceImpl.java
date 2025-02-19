package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.LambdaQueryWrapper;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.model.system.SysRoleMenu;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.system.SysUserMenu;
import glgc.jjgys.model.vo.AssginMenuVo;
import glgc.jjgys.model.vo.RouterVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.SysMenuMapper;
import glgc.jjgys.system.mapper.SysRoleMenuMapper;
import glgc.jjgys.system.service.SysMenuService;
import glgc.jjgys.system.service.SysUserMenuService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.MenuHelper;
import glgc.jjgys.system.utils.RouterHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import java.util.*;
import java.util.stream.Collectors;

/**
 * <p>
 * 菜单表 服务实现类
 * </p>
 */
@Service
public class SysMenuServiceImpl extends ServiceImpl<SysMenuMapper, SysMenu> implements SysMenuService {

    @Autowired
    private SysRoleMenuMapper sysRoleMenuMapper;

    @Autowired
    private SysMenuMapper sysMenuMapper;

    //菜单列表（树形）
    @Override
    public List<SysMenu> findNodes() {
        //获取所有菜单
        List<SysMenu> sysMenuList = baseMapper.selectList(null);
        //所有菜单数据转换要求数据格式
        List<SysMenu> resultList = MenuHelper.bulidTree(sysMenuList);
        return resultList;
    }


    public List<SysMenu> findNodes1() {
        //获取所有菜单
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.in("type", 0, 1);
        List<SysMenu> sysMenuList = baseMapper.selectList(wrapper);
        //所有菜单数据转换要求数据格式
        List<SysMenu> resultList = MenuHelper.bulidTree(sysMenuList);
        return resultList;
    }

    //菜单列表（树形）
    @Override
    public List<SysMenu> findNodesByAuthUser(HttpServletRequest request) {
        SysUser userInfo = getUserInfoByRequest(request);
        String type = userInfo.getType();
        if ("2".equals(type) || "4".equals(type)){
            // 根据公司id查询
            //获取所有菜单
            LambdaQueryWrapper<SysMenu> sysMenuLambdaQueryWrapper = new LambdaQueryWrapper<>();
            List<SysMenu> nodesByUserId = findNodesByUserId(String.valueOf(userInfo.getId()));

            List<Long> longStream = nodesByUserId.stream().map(m->m.getId()).collect(Collectors.toList());
            sysMenuLambdaQueryWrapper.eq(SysMenu::getCompanyId,userInfo.getCompanyId());
            if (longStream!= null && longStream.size() > 0) {
                sysMenuLambdaQueryWrapper.notIn(SysMenu::getId, longStream);
            }
            List<SysMenu> sysMenuList = baseMapper.selectList(sysMenuLambdaQueryWrapper);

            nodesByUserId.addAll(sysMenuList);

            //所有菜单数据转换要求数据格式
            // 使用 Stream 进行筛选
            List<SysMenu> filteredNodes = nodesByUserId.stream()
                    .filter(menu -> menu.getType() == 0 || menu.getType() == 1)
                    .collect(Collectors.toList());
            //List<SysMenu> resultList = MenuHelper.bulidTree(nodesByUserId);
            List<SysMenu> resultList = MenuHelper.bulidTree(filteredNodes);
            return resultList;
        }else if ("3".equals(type)){
            // 根据 用户id 查询
            return findNodesByUserId(String.valueOf(userInfo.getId()));
        }else if ("1".equals(type)) {
            // 超级管理员 查询全部
            //return findNodes();
            return findNodes1();
        }
        return new ArrayList<>();
    }

    @Resource
    SysUserService sysUserService;
    private SysUser getUserInfoByRequest(HttpServletRequest request){
        //获取请求头token字符串,因为把token放到请求头不存在跨域
        String token = request.getHeader("token");
        //从token字符串获取用户名称（id）
        String userId = JwtHelper.getUserId(token);
        //根据用户名称获取用户信息
        return sysUserService.getById(userId);
    }
    //菜单列表（树形）
    @Resource
    SysUserMenuService sysUserMenuService;

    @Override
    public List<SysMenu> findNodesByUserId(String userId) {
        //获取所有菜单
        LambdaQueryWrapper<SysUserMenu> sysMenuLambdaQueryWrapper = new LambdaQueryWrapper<>();
        sysMenuLambdaQueryWrapper.eq(SysUserMenu::getUserId,userId);
        List<SysUserMenu> list = sysUserMenuService.list(sysMenuLambdaQueryWrapper);

        List<Long> menuList = list.stream().map(m -> m.getMenuId()).collect(Collectors.toList());
        if (menuList == null || menuList.size()  == 0){
            return new ArrayList<>();
        }
        LambdaQueryWrapper<SysMenu> sysMenuLambdaQueryWrapper1 = new LambdaQueryWrapper<>();
        sysMenuLambdaQueryWrapper1.in(SysMenu::getId,menuList);
        List<SysMenu> sysMenuList = baseMapper.selectList(sysMenuLambdaQueryWrapper1);
        //所有菜单数据转换要求数据格式
        return sysMenuList;

    }

    //删除菜单
    @Override
    public void removeMenuById(String id) {
        //查询当前删除菜单下面是否子菜单
        //根据id = parentid
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id",id);
        Integer count = baseMapper.selectCount(wrapper);
        if(count > 0) {//有子菜单
            throw new JjgysException(201,"请先删除子菜单");
        }
        //调用删除
        baseMapper.deleteById(id);
    }

    //根据角色分配菜单
    @Override
    public List<SysMenu> findMenuByRoleId(String roleId) {
        //获取所有菜单 status=1
        /*QueryWrapper<SysMenu> wrapperMenu = new QueryWrapper<>();
        wrapperMenu.eq("status",1);*/
        List<SysMenu> menuList = baseMapper.selectList(null);

        //根据角色id查询 角色分配过的菜单列表
        QueryWrapper<SysRoleMenu> wrapperRoleMenu = new QueryWrapper<>();
        wrapperRoleMenu.eq("role_id",roleId);
        List<SysRoleMenu> roleMenus = sysRoleMenuMapper.selectList(wrapperRoleMenu);

        //从第二步查询列表中，获取角色分配所有菜单id
        List<String> roleMenuIds = new ArrayList<>();
        for (SysRoleMenu sysRoleMenu:roleMenus) {
            String menuId = sysRoleMenu.getMenuId();
            roleMenuIds.add(menuId);
        }

        //数据处理：isSelect 如果菜单选中 true，否则false
        // 拿着分配菜单id 和 所有菜单比对，有相同的，让isSelect值true
        for (SysMenu sysMenu:menuList) {
            String s = sysMenu.getId().toString();
            if(roleMenuIds.contains(s)) {
                sysMenu.setSelect(true);
            } else {
                sysMenu.setSelect(false);
            }
        }

        //转换成树形结构为了最终显示 MenuHelper方法实现
        List<SysMenu> sysMenus = MenuHelper.bulidTree(menuList);
        return sysMenus;
    }

    //给角色分配菜单权限
    @Override
    public void doAssign(AssginMenuVo assginMenuVo) {
        //根据角色id删除菜单权限
        QueryWrapper<SysRoleMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("role_id",assginMenuVo.getRoleId());
        sysRoleMenuMapper.delete(wrapper);

        //遍历菜单id列表，一个一个进行添加
        List<String> menuIdList = assginMenuVo.getMenuIdList();
        for (String menuId:menuIdList) {
            SysRoleMenu sysRoleMenu = new SysRoleMenu();
            sysRoleMenu.setMenuId(menuId);
            sysRoleMenu.setRoleId(assginMenuVo.getRoleId());
            sysRoleMenuMapper.insert(sysRoleMenu);
        }
    }

    //根据userid查询菜单权限值
    @Override
    public List<RouterVo> getUserMenuList(String userId,HttpServletRequest request) {
        SysUser userInfoByRequest = getUserInfoByRequest(request);
        String type = userInfoByRequest.getType();

        //admin是超级管理员，操作所有内容
        List<SysMenu> sysMenuList = null;
        //判断userid值是1代表超级管理员，查询所有权限数据
        if("1".equals(userId)) {
            QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
            wrapper.orderByAsc("sort_value");
            sysMenuList = baseMapper.selectList(wrapper);
        } else {
            if ("2".equals(type) || "4".equals(type) || "4".equals(type)){
                sysMenuList = findNodesByAuthUser(request);
                Collections.sort(sysMenuList, new Comparator<SysMenu>() {
                    @Override
                    public int compare(SysMenu o1, SysMenu o2) {
                        // 名字相同时按照 qdzh 排序
                        String zh = o1.getSortValue().toString();
                        String zh1 = o2.getSortValue().toString();
                        return zh.compareTo(zh1);
                    }
                });
                //转换成前端路由要求格式数据
                List<RouterVo> routerVoList = RouterHelper.buildRouters(sysMenuList);
                return routerVoList ;

            }else if ("3".equals(type)){
                //如果userid不是1，其他类型用户，查询这个用户权限
                sysMenuList = findNodesByAuthUser(request);
                Collections.sort(sysMenuList, new Comparator<SysMenu>() {
                    @Override
                    public int compare(SysMenu o1, SysMenu o2) {
                        // 名字相同时按照 qdzh 排序
                        String zh = o1.getSortValue().toString();
                        String zh1 = o2.getSortValue().toString();
                        return zh.compareTo(zh1);
                    }
                });
                List<SysMenu> menuList = MenuHelper.bulidTree(sysMenuList);
                //转换成前端路由要求格式数据
                List<RouterVo> routerVoList = RouterHelper.buildRouters(menuList);
                return routerVoList;
            }

        }

        //构建是树形结构
        List<SysMenu> sysMenuTreeList = MenuHelper.bulidTree(sysMenuList);
        Collections.sort(sysMenuTreeList, new Comparator<SysMenu>() {
            @Override
            public int compare(SysMenu o1, SysMenu o2) {
                // 名字相同时按照 qdzh 排序
                String zh = o1.getSortValue().toString();
                String zh1 = o2.getSortValue().toString();
                return zh.compareTo(zh1);
            }
        });

        //转换成前端路由要求格式数据
        List<RouterVo> routerVoList = RouterHelper.buildRouters(sysMenuTreeList);
        return routerVoList;
    }

    //根据userid查询按钮权限值
    @Override
    //public List<String> getUserButtonList(String userId) {
    public List<String> getUserButtonList(String userId,HttpServletRequest request) {
        /*List<SysMenu> sysMenuList = null;
        //判断是否管理员
        if("1".equals(userId)) {
            sysMenuList =
                    baseMapper.selectList(new QueryWrapper<SysMenu>().eq("status",1));
        } else {
            sysMenuList = baseMapper.findMenuListUserId(userId);
        }
        //sysMenuList遍历
        List<String> permissionList = new ArrayList<>();
        for (SysMenu sysMenu:sysMenuList) {
            // type=2
            if(sysMenu.getType()==2) {
                String perms = sysMenu.getPerms();
                permissionList.add(perms);
            }
        }
        return permissionList;*/
        List<SysMenu> sysMenuList = null;
        //判断是否管理员

        SysUser userInfo = new SysUser();
        if (request == null){
            userInfo = sysUserService.getById(userId);
        }else {
            userInfo = getUserInfoByRequest(request);
        }
        String type = userInfo.getType();
        if ("2".equals(type) || "4".equals(type) || "4".equals(type)){
            // 根据公司id查询
            //获取所有按钮
            LambdaQueryWrapper<SysMenu> sysMenuLambdaQueryWrapper = new LambdaQueryWrapper<>();
            List<SysMenu> nodesByUserId = findNodesByUserId(String.valueOf(userInfo.getId()));
            sysMenuLambdaQueryWrapper.eq(SysMenu::getCompanyId,userInfo.getCompanyId());
            List<SysMenu> sysMenuListT = baseMapper.selectList(sysMenuLambdaQueryWrapper);
            // 去重
            List<SysMenu> lists1 = new ArrayList<>();
            if (nodesByUserId!= null && nodesByUserId.size()>0){
                lists1.addAll(nodesByUserId);
            }
            if (sysMenuListT!= null && sysMenuListT.size()>0){
                for (SysMenu sysMenu : sysMenuListT) {
                    Optional<SysMenu> first = lists1.stream().filter(m -> m.getId().equals(sysMenu.getId())).findFirst();
                    if (first.isPresent()){
                        lists1.add(sysMenu);
                    }
                }
            }
            sysMenuList = lists1;
        }else if ("3".equals(type)){
            // 根据 用户id 查询
            //获取所有菜单
            sysMenuList = findNodesByUserId(String.valueOf(userInfo.getId()));
        }else if ("1".equals(type)) {
            // 超级管理员 查询全部
            //sysMenuList = findNodes();
            sysMenuList = baseMapper.selectList(new QueryWrapper<SysMenu>());
        }
        //sysMenuList遍历
        List<String> permissionList = new ArrayList<>();
        for (SysMenu sysMenu:sysMenuList) {
            // type=2
            if(sysMenu.getType()==2) {
                String perms = sysMenu.getPerms();
                permissionList.add(perms);
            }
        }
        return permissionList;
    }

    @Override
    public List<SysMenu> selectname(String proname) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", proname);
        List<SysMenu> htdList = sysMenuMapper.selectList(wrapper);
        return htdList;
    }

    @Override
    public List<SysMenu> selectscname(Long pronameid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", "实测数据");
        wrapper.eq("parent_id", pronameid);
        List<SysMenu> htdList = sysMenuMapper.selectList(wrapper);
        return htdList;
    }

    @Override
    public boolean delecthtd(Long scid, String htd) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", htd);
        wrapper.eq("parent_id", scid);
        sysMenuMapper.delete(wrapper);
        return true;
    }

    @Override
    public List<SysMenu> selecthtd(Long scid, String htd) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", scid);
        List<SysMenu> htdList = sysMenuMapper.selectList(wrapper);
        return htdList;

    }

    @Override
    public boolean delectfbgc(Long fbgcid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", fbgcid);
        List<SysMenu> list = sysMenuMapper.selectList(wrapper);
        for (SysMenu sysMenu : list) {
            Long id = sysMenu.getId();
            sysMenuMapper.deleteById(id);
        }
        return true;
    }

    @Override
    public SysMenu selectcdinfo(String proName) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", proName);
        wrapper.eq("parent_id", "1600779312636739586");
        SysMenu sysMenu = sysMenuMapper.selectOne(wrapper);
        return sysMenu;
    }

    @Override
    public SysMenu getscChildrenMenu(Long proid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", "实测数据");
        wrapper.eq("parent_id", proid);
        SysMenu sysMenu = sysMenuMapper.selectOne(wrapper);
        return sysMenu;

    }

    @Override
    public List<SysMenu> getAllHtd(Long scid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", scid);
        List<SysMenu> menuList = sysMenuMapper.selectList(wrapper);
        return menuList;
    }

    @Override
    public void removeFbgc(List<SysMenu> htdlist) {
        for (SysMenu sysMenu : htdlist) {
            Long htdid = sysMenu.getId();
            QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
            wrapper.eq("parent_id", htdid);
            List<SysMenu> menuList = sysMenuMapper.selectList(wrapper);//所有分部工程的信息
            removefbgcInfo(menuList);
            removeuserrolrInfo(menuList);
        }
        for (SysMenu sysMenu : htdlist) {
            sysMenuMapper.deleteById(sysMenu.getId());
            //删除SysRoleMenu表中的数据
            QueryWrapper<SysRoleMenu> wrapper = new QueryWrapper<>();
            wrapper.eq("menu_id",sysMenu.getId());
            sysRoleMenuMapper.delete(wrapper);
        }
    }

    /**
     * 删除sys_role_menu表中的数据
     * @param menuList
     */
    private void removeuserrolrInfo(List<SysMenu> menuList) {
        for (SysMenu sysMenu : menuList) {
            Long id = sysMenu.getId();
            QueryWrapper<SysRoleMenu> wrapper = new QueryWrapper<>();
            wrapper.eq("menu_id",id);
            sysRoleMenuMapper.delete(wrapper);
        }
    }

    @Override
    public void delectChildrenMenu(Long proid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", proid);
        List<SysMenu> menuList = sysMenuMapper.selectList(wrapper);
        for (SysMenu sysMenu : menuList) {
            sysMenuMapper.deleteById(sysMenu.getId());
        }
    }

    private void removefbgcInfo(List<SysMenu> menuList) {
        for (SysMenu sysMenu : menuList) {
            sysMenuMapper.deleteById(sysMenu.getId());
        }
    }
}
