package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcLmgcLmwcJgfcService;
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
@RequestMapping("/jjg/jgfc/lmwcbkml")
public class JjgFbgcLmgcLmwcJgfcController {

    @Autowired
    private JjgFbgcLmgcLmwcJgfcService jjgFbgcLmgcLmwcJgfcService;

    @Autowired
    private OperLogService operLogService;


    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname) throws IOException {
        List<Map<String,Object>> htdList = jjgFbgcLmgcLmwcJgfcService.selecthtd(proname);
        List<String> fileName = new ArrayList<>();
        if (htdList!=null){
            for (Map<String, Object> map1 : htdList) {
                String htd = map1.get("htd").toString();
                fileName.add(htd+File.separator+"13路面弯沉(贝克曼梁法)");
            }
        }
        String zipname = "路面弯沉(贝克曼梁法)";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        QueryWrapper<JjgFbgcLmgcLmwcJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);
        jjgFbgcLmgcLmwcJgfcService.remove(queryWrapper);
        return Result.ok();

    }

    @ApiOperation("生成路面弯沉鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        boolean is_Success = jjgFbgcLmgcLmwcJgfcService.generateJdb(proname);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看路面弯沉鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLmgcLmwcJgfcService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("路面弯沉模板文件导出")
    @GetMapping("exportlmwc")
    public void exportlmwc(HttpServletResponse response){
        jjgFbgcLmgcLmwcJgfcService.exportlmwc(response);
    }

    @ApiOperation(value = "路面弯沉数据文件导入")
    @PostMapping("importlmwc")
    public Result importlmwc(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLmwcJgfcService.importlmwc(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcLmwcJgfc jjgFbgcLmgcLmwc) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcLmwcJgfc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLmwc != null){
            QueryWrapper<JjgFbgcLmgcLmwcJgfc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLmwc.getProname());

            Date jcsj = jjgFbgcLmgcLmwc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcLmwc.getZh())) {
                wrapper.like("zh", jjgFbgcLmgcLmwc.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLmwcJgfc> pageModel = jjgFbgcLmgcLmwcJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除路面弯沉数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLmwcJgfcService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLjwc/{id}")
    public Result getLjwc(@PathVariable String id) {
        JjgFbgcLmgcLmwcJgfc user = jjgFbgcLmgcLmwcJgfcService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路面弯沉数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcLmwcJgfc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcLmgcLmwcJgfcService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("竣工路面弯沉贝克曼梁法数据");
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

