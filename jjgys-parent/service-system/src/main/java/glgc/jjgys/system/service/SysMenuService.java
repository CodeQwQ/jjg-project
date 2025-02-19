package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.model.vo.AssginMenuVo;
import glgc.jjgys.model.vo.RouterVo;

import javax.servlet.http.HttpServletRequest;
import java.util.List;

/**
 * <p>
 * 菜单表 服务类
 * </p>
 */
public interface SysMenuService extends IService<SysMenu> {

    //菜单列表（树形）
    List<SysMenu> findNodes();

    List<SysMenu> findNodesByAuthUser(HttpServletRequest request);

    List<SysMenu> findNodesByUserId(String userId);

    // //删除菜单
    void removeMenuById(String id);

    //根据角色分配菜单
    List<SysMenu> findMenuByRoleId(String roleId)   ;

    //给角色分配菜单权限
    void doAssign(AssginMenuVo assginMenuVo);

    //根据userid查询菜单权限值
    List<RouterVo> getUserMenuList(String id,HttpServletRequest request);

    //根据userid查询按钮权限值
    List<String> getUserButtonList(String id,HttpServletRequest request);

    List<SysMenu> selectname(String proname);

    List<SysMenu> selectscname(Long pronameid);

    boolean delecthtd(Long scid, String htd);

    List<SysMenu> selecthtd(Long scid, String htd);

    boolean delectfbgc(Long fbgcid);

    SysMenu selectcdinfo(String proName);

    SysMenu getscChildrenMenu(Long proid);

    List<SysMenu> getAllHtd(Long scid);

    void removeFbgc(List<SysMenu> htdlist);

    void delectChildrenMenu(Long proid);

}
