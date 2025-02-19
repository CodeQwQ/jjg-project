package glgc.jjgys.model.system;

import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableName;
import glgc.jjgys.model.base.BaseEntity;
import io.swagger.annotations.ApiModel;
import lombok.Data;

import java.util.Date;

/**
 * 用户菜单表(SysUserMenu)表实体类
 *
 * @author makejava
 * @since 2023-12-06 11:58:47
 */

@Data
@TableName("sys_user_menu")
@ApiModel(description = "用户菜单表")
@SuppressWarnings("serial")
public class SysUserMenu extends BaseEntity {



    @TableField("user_id")
    private Long userId;


    @TableField("menu_id")
    private Long menuId;


    @TableField("create_time")
    private Date createTime;

}

