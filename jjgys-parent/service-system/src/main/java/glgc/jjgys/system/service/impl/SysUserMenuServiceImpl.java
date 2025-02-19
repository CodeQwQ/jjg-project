package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.LambdaQueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.system.SysUserMenu;
import glgc.jjgys.system.mapper.SysUserMenuMapper;
import glgc.jjgys.system.service.SysUserMenuService;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * 用户菜单表(SysUserMenu)表服务实现类
 *
 * @author makejava
 * @since 2023-12-06 12:04:59
 */
@Service
public class SysUserMenuServiceImpl extends ServiceImpl<SysUserMenuMapper, SysUserMenu> implements SysUserMenuService {

    @Override
    public Boolean saveMenus(List<SysUserMenu> data) {
        if (data == null || data.size() == 0){
            return false;
        }
        // 删除原来的
        Long userId = data.get(0).getUserId();
        remove(new LambdaQueryWrapper<SysUserMenu>().eq(SysUserMenu::getUserId,userId));

        return saveBatch(data);
    }
}

