package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgRyinfo;
import glgc.jjgys.model.project.JjgWgjc;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgRyinfoService;
import glgc.jjgys.system.service.OperLogService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Date;
import java.util.List;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-12-05
 */
@RestController
@RequestMapping("/jjg/ryinfo")
public class JjgRyinfoController {

    @Autowired
    private JjgRyinfoService jjgRyinfoService;

    @Autowired
    private OperLogService operLogService;

    @ApiOperation("人员信息文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response, @RequestParam String projectname){
        jjgRyinfoService.export(response,projectname);
    }

    @ApiOperation(value = "人员信息数据文件导入")
    @PostMapping("importryxx")
    public Result importryxx(@RequestParam("file") MultipartFile file, String proname) {
        jjgRyinfoService.importryxx(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgRyinfo jjgRyinfo) {
        //创建page对象
        Page<JjgRyinfo> pageParam = new Page<>(current, limit);
        if (jjgRyinfo != null) {
            QueryWrapper<JjgRyinfo> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgRyinfo.getProname());
            //调用方法分页查询
            IPage<JjgRyinfo> pageModel = jjgRyinfoService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除人员信息数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgRyinfoService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabx/{id}")
    public Result getJabx(@PathVariable String id) {
        JjgRyinfo user = jjgRyinfoService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改人员信息数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgRyinfo user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgRyinfoService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd("-");
            sysOperLog.setFbgc("-");
            sysOperLog.setTitle("人员信息数据");
            sysOperLog.setBusinessType("修改");
            sysOperLog.setOperName(JwtHelper.getUsername(request.getHeader("token")));
            sysOperLog.setOperIp(IpUtil.getIpAddress(request));
            sysOperLog.setOperTime(new Date());
            operLogService.saveSysLog(sysOperLog);
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

