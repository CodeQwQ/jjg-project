package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgRyinfoJgfc;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgRyinfoJgfcService;
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
 * @since 2024-05-02
 */
@RestController
@RequestMapping("/jjg/ryinfojgfc")
public class JjgRyinfoJgfcController {

    @Autowired
    private OperLogService operLogService;

    @Autowired
    private JjgRyinfoJgfcService jjgRyinfoJgfcService;

    @ApiOperation("人员信息文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response, @RequestParam String projectname){
        jjgRyinfoJgfcService.export(response,projectname);
    }

    @ApiOperation(value = "人员信息数据文件导入")
    @PostMapping("importryxx")
    public Result importryxx(@RequestParam("file") MultipartFile file, String proname) {
        jjgRyinfoJgfcService.importryxx(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgRyinfoJgfc jjgRyinfo) {
        //创建page对象
        Page<JjgRyinfoJgfc> pageParam = new Page<>(current, limit);
        if (jjgRyinfo != null) {
            QueryWrapper<JjgRyinfoJgfc> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgRyinfo.getProname());
            //调用方法分页查询
            IPage<JjgRyinfoJgfc> pageModel = jjgRyinfoJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除人员信息数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgRyinfoJgfcService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabx/{id}")
    public Result getJabx(@PathVariable String id) {
        JjgRyinfoJgfc user = jjgRyinfoJgfcService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改人员信息数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgRyinfoJgfc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgRyinfoJgfcService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd("-");
            sysOperLog.setFbgc("-");
            sysOperLog.setTitle("竣工人员信息数据");
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

