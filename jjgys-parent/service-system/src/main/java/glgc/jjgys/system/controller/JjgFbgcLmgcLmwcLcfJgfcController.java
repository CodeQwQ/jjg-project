package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcLcfJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcLmgcLmwcLcfJgfcService;
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
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2024-04-30
 */
@RestController
@RequestMapping("/jjg/jgfc/lmwclcf")
public class JjgFbgcLmgcLmwcLcfJgfcController {

    @Autowired
    private JjgFbgcLmgcLmwcLcfJgfcService jjgFbgcLmgcLmwcLcfJgfcService;

    @Autowired
    private OperLogService operLogService;


    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname) throws IOException {
        List<Map<String,Object>> htdList = jjgFbgcLmgcLmwcLcfJgfcService.selecthtd(proname);
        List<String> fileName = new ArrayList<>();
        if (htdList != null){
            for (Map<String, Object> map1 : htdList) {
                String htd = map1.get("htd").toString();
                fileName.add(htd+ File.separator+"13路面弯沉(落锤法)");
            }
        }
        String zipname = "路面弯沉(落锤法)";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);
    }

    @ApiOperation("生成路面弯沉落锤法鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        boolean is_Success = jjgFbgcLmgcLmwcLcfJgfcService.generateJdb(proname);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("路面弯沉落锤法模板文件导出")
    @GetMapping("exportlmwclcf")
    public void exportlmwclcf(HttpServletResponse response){
        jjgFbgcLmgcLmwcLcfJgfcService.exportlmwclcf(response);
    }

    @ApiOperation(value = "路面弯沉落锤法数据文件导入")
    @PostMapping("importlmwclcf")
    public Result importlmwclcf(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLmwcLcfJgfcService.importlmwclcf(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcLmwcLcfJgfc jjgFbgcLmgcLmwcLcf) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcLmwcLcfJgfc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLmwcLcf != null){
            QueryWrapper<JjgFbgcLmgcLmwcLcfJgfc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLmwcLcf.getProname());

            Date jcsj = jjgFbgcLmgcLmwcLcf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcLmwcLcf.getZh())) {
                wrapper.like("zh", jjgFbgcLmgcLmwcLcf.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLmwcLcfJgfc> pageModel = jjgFbgcLmgcLmwcLcfJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除路面弯沉落锤法数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLmwcLcfJgfcService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        QueryWrapper<JjgFbgcLmgcLmwcLcfJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);
        jjgFbgcLmgcLmwcLcfJgfcService.remove(queryWrapper);
        return Result.ok();

    }

    @ApiOperation("根据id查询")
    @GetMapping("getLmwcLcf/{id}")
    public Result getLmwcLcf(@PathVariable String id) {
        JjgFbgcLmgcLmwcLcfJgfc user = jjgFbgcLmgcLmwcLcfJgfcService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路面弯沉落锤法数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcLmwcLcfJgfc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcLmgcLmwcLcfJgfcService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("竣工路面弯沉落锤法数据");
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

