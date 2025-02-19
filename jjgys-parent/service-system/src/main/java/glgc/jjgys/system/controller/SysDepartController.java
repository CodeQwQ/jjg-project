package glgc.jjgys.system.controller;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.system.SysDept;
import glgc.jjgys.system.service.SysDepartService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import java.util.List;

@Api(tags = "部门接口")
@RestController
@RequestMapping(value = "/dept")
@CrossOrigin
public class SysDepartController {
    @Autowired
    private SysDepartService sysDepartService;

    @ApiOperation("查询部门名称")
    @GetMapping("selectdept")
    public Result selectdept(){
        List<SysDept> list = sysDepartService.list();
        return Result.ok(list);
    }

    @ApiOperation("查询部门树")
    @GetMapping("selectdeptTree")
    public Result selectdeptTree(HttpServletRequest request){
        List<SysDept> list = sysDepartService.selectdeptTree(request);
        return Result.ok(list);
    }

    /**
     * 部门管理界面的显示
     * @return
     */
    @ApiOperation("查询部门树(部门管理)")
    @GetMapping("selectdeptinfo")
    public Result selectdeptinfo(){
        List<SysDept> list = sysDepartService.selectdeptinfo();
        return Result.ok(list);
    }

    /**
     *
     * @param dept
     * @return
     */
    @PostMapping("addcompany")
    public Result addcompany(@RequestBody SysDept dept) {
        String name = dept.getName();
        QueryWrapper<SysDept> wrapper = new QueryWrapper<>();
        wrapper.eq("name",name);
        wrapper.eq("parent_id",0);
        List<SysDept> list = sysDepartService.list(wrapper);
        if (list == null && list.size() < 0){
            return Result.fail().message("新增公司'" + dept.getName() + "'失败，公司名称已存在");
        }
        SysDept sysDept = new SysDept();
        long l = generateId();
        sysDept.setId(l);
        sysDept.setName(dept.getName());
        sysDept.setParentId(Long.valueOf(0));
        sysDept.setTreePath("0");
        sysDept.setStatus(1);
        sysDept.setCompanyId(l);
        sysDepartService.save(sysDept);
        return Result.ok();
    }

    @ApiOperation("查询公司")
    @GetMapping("selectcompany")
    public Result selectcompany(){
        QueryWrapper<SysDept> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id",0);
        wrapper.eq("tree_path",0);
        List<SysDept> list = sysDepartService.list(wrapper);
        return Result.ok(list);
    }

    @PostMapping("checkbm")
    public Result checkbm(@RequestBody SysDept sysDept) {
        Long id = sysDept.getCompanyId();
        QueryWrapper<SysDept> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id",id);
        List<SysDept> list = sysDepartService.list(wrapper);
        if (list != null && list.size()>0){
            return Result.fail().message("改公司下已有部门，不允许添加");
        }else {
            return Result.ok();
        }
    }



    /**
     * 新增部门
     * @param sysDept
     * @return
     */
    @PostMapping("add")
    public Result add(@RequestBody SysDept sysDept) {
        Long id = sysDept.getCompanyId();
        QueryWrapper<SysDept> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id",id);
        List<SysDept> list = sysDepartService.list(wrapper);

        String deptnamevo = sysDept.getName();

        boolean isNamePresent = list.stream()
                .map(SysDept::getName)
                .anyMatch(name -> name.equals(deptnamevo));

        if (isNamePresent) {
            return Result.fail().message("新增部门'" + sysDept.getName() + "'失败，部门名称已存在");
        } else {
            sysDept.setId(generateId());
            sysDept.setParentId(id);
            sysDept.setTreePath("0,"+id);
            sysDepartService.save(sysDept);
            //sysDepartService.insertDept(sysDept);
            return Result.ok();
        }
    }

    /**
     * 修改部门
     */
    @PutMapping("updata")
    public Result edit(@RequestBody SysDept dept)
    {
        boolean is_Success = sysDepartService.updateById(dept);
        return is_Success ? Result.ok(): Result.fail();
    }

    /**
     * 删除部门
     */
    @DeleteMapping("/remove/{deptId}")
    public Result remove(@PathVariable Long deptId)
    {
        if (sysDepartService.hasChildByDeptId(deptId))
        {
            return Result.fail().message("存在下级部门,不允许删除");
        }

        if (sysDepartService.checkDeptExistUser(deptId))
        {
            return Result.fail().message("部门存在用户,不允许删除");
        }
        sysDepartService.deleteDeptById(deptId);
        return Result.ok();
    }

    /**
     * 生成id
     * @return
     */
    public static long generateId() {
        // 获取当前时间戳
        long timestamp = System.currentTimeMillis();

        // 生成两位随机数
        int random = (int) (Math.random() * 100);
        // 将随机数转换为两位数
        String formattedRandom = String.format("%02d", random);

        // 将时间戳和随机数拼接，确保总长度为20位
        String idString = String.valueOf(timestamp) + formattedRandom;
        // 截取前20位并转换为long类型
        long id = Long.parseLong(idString.substring(0, 15));

        return id;
    }


}
