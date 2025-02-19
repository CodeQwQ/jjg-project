package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.SysUserMenu;

import java.util.List;

/**
 * 用户菜单表(SysUserMenu)表服务接口
 *
 * @author makejava
 * @since 2023-12-06 12:04:59
 */
public interface SysUserMenuService extends IService<SysUserMenu> {

    Boolean saveMenus(List<SysUserMenu> data);
}

