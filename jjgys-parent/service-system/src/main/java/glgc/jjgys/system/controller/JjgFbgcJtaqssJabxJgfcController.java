package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxJgfcService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2024-05-20
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jabxjgfc")
public class JjgFbgcJtaqssJabxJgfcController {

    @Autowired
    private JjgFbgcJtaqssJabxJgfcService jjgFbgcJtaqssJabxJgfcService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname) throws IOException {
        List<String> htdList = jjgFbgcJtaqssJabxJgfcService.selecthtd(proname);
        List fileName = new ArrayList();
        if (htdList != null && htdList.size()>0) {
            for (String htd : htdList) {
                String fileName1 = "57交安标线厚度";
                String fileName2 = "57交安标线逆反射系数";
                fileName.add(htd+File.separator+fileName1);
                fileName.add(htd+File.separator+fileName2);
            }
        }
        String zipname = "交安标线鉴定表";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);

    }

    @ApiOperation("生成交安标线鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcJtaqssJabxJgfcService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看交安标线鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJabxJgfcService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安标线模板文件导出")
    @GetMapping("exportjabx")
    public void exportjabx(HttpServletResponse response){
        jjgFbgcJtaqssJabxJgfcService.exportjabx(response);
    }

    @ApiOperation(value = "交安标线数据文件导入")
    @PostMapping("importjabx")
    public Result importjabx(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJabxJgfcService.importjabx(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabxJgfc jjgFbgcJtaqssJabx) throws ParseException {
        //创建page对象
        Page<JjgFbgcJtaqssJabxJgfc> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabx != null) {
            QueryWrapper<JjgFbgcJtaqssJabxJgfc> wrapper = new QueryWrapper<>();
            wrapper.eq("proname", jjgFbgcJtaqssJabx.getProname());
            Date jcsj = jjgFbgcJtaqssJabx.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcJtaqssJabx.getWz())) {
                wrapper.like("wz", jjgFbgcJtaqssJabx.getWz());
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabxJgfc> pageModel = jjgFbgcJtaqssJabxJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }


    @ApiOperation("批量删除交安标线数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabxJgfcService.removeByIds(idList);
        return hd ? Result.ok():Result.fail().message("删除失败！");
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        QueryWrapper<JjgFbgcJtaqssJabxJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);
        boolean remove = jjgFbgcJtaqssJabxJgfcService.remove(queryWrapper);
        return Result.ok();

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabx/{id}")
    public Result getJabx(@PathVariable String id) {
        JjgFbgcJtaqssJabxJgfc user = jjgFbgcJtaqssJabxJgfcService.getById(id);
        return Result.ok(user);
    }

    //@Log(title = "交安标线数据",businessType = BusinessType.UPDATE)
    @ApiOperation("修改交安标线数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabxJgfc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcJtaqssJabxJgfcService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("竣工交安标线数据");
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

