package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.api.ApiController;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.system.SysUserMenu;
import glgc.jjgys.system.service.SysUserMenuService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.*;

import javax.annotation.Resource;
import java.io.Serializable;
import java.util.HashMap;
import java.util.List;

/**
 * 用户菜单表(SysUserMenu)表控制层
 *
 * @author makejava
 * @since 2023-12-06 12:13:15
 */
@Api(tags = "用户菜单")
@RestController
@RequestMapping("/admin/system/sysUserMenu")
public class SysUserMenuController extends ApiController {
    /**
     * 服务对象
     */
    @Resource
    private SysUserMenuService sysUserMenuService;

    /**
     * 分页查询所有数据
     *
     * @param page        分页对象
     * @param sysUserMenu 查询实体
     * @return 所有数据
     */
    @ApiOperation("分页查询所有数据")
    @GetMapping
    public Result selectAll(Page<SysUserMenu> page, SysUserMenu sysUserMenu) {
        return Result.ok(this.sysUserMenuService.page(page, new QueryWrapper<>(sysUserMenu)));
    }

    /**
     * 通过主键查询单条数据
     *
     * @param id 主键
     * @return 单条数据
     */
    @ApiOperation("通过主键查询单条数据")
    @GetMapping("{id}")
    public Result selectOne(@PathVariable Serializable id) {
        return Result.ok(this.sysUserMenuService.getById(id));
    }

    /**
     * 新增数据
     *
     * @param sysUserMenu 实体对象
     * @return 新增结果
     */
    @ApiOperation("新增数据")
    @PostMapping
    public Result insert(@RequestBody SysUserMenu sysUserMenu) {
        return Result.ok(this.sysUserMenuService.save(sysUserMenu));
    }

    /**
     * 修改数据
     *
     * @param sysUserMenu 实体对象
     * @return 修改结果
     */
    @ApiOperation("修改数据")
    @PutMapping
    public Result update(@RequestBody SysUserMenu sysUserMenu) {
        return Result.ok(this.sysUserMenuService.updateById(sysUserMenu));
    }

    /**
     * 删除数据
     *
     * @param idList 主键结合
     * @return 删除结果
     */
    @ApiOperation("删除数据")
    @DeleteMapping
    public Result delete(@RequestParam("idList") List<Long> idList) {
        return Result.ok(this.sysUserMenuService.removeByIds(idList));
    }

    //保存
    @ApiOperation("保存修改")
    @PostMapping("saveMenus")
    public Result saveMenus(@RequestBody HashMap<String,List<SysUserMenu>> map) {
        System.out.println(map);
        return Result.ok(this.sysUserMenuService.saveMenus(map.get("data")));
    }


}

