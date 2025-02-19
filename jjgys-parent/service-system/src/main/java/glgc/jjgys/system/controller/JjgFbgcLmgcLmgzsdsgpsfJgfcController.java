package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcLmgzsdsgpsfJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.service.JjgFbgcLmgcLmgzsdsgpsfJgfcService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
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
 * @since 2023-09-23
 */
@RestController
@RequestMapping("/jjg/jgfc/lmgzsdsgpsf")
public class JjgFbgcLmgcLmgzsdsgpsfJgfcController {

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfJgfcService jjgFbgcLmgcLmgzsdsgpsfJgfcService;

    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname) throws IOException {
        List<Map<String,Object>> htdList = jjgFbgcLmgcLmgzsdsgpsfJgfcService.selecthtd(proname);
        List<String> fileName = new ArrayList<>();
        if (htdList!=null){
            for (Map<String, Object> map1 : htdList) {
                String htd = map1.get("htd").toString();
                fileName.add(htd+File.separator+"20构造深度手工铺沙法");
            }
        }
        String zipname = "构造深度手工铺沙法";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);
    }

    @ApiOperation("生成构造深度手工铺沙法鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        boolean is_Success =  jjgFbgcLmgcLmgzsdsgpsfJgfcService.generateJdb(proname);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("构造深度手工铺沙法模板文件导出")
    @GetMapping("exportlmgzsdsgpsf")
    public void exportlmgzsdsgpsf(HttpServletResponse response){
        jjgFbgcLmgcLmgzsdsgpsfJgfcService.exportlmgzsdsgpsf(response);
    }


    @ApiOperation(value = "构造深度手工铺沙法数据文件导入")
    @PostMapping("importlmgzsdsgpsf")
    public Result importlmgzsdsgpsf(@RequestParam("file") MultipartFile file, String proname,String username) {
        jjgFbgcLmgcLmgzsdsgpsfJgfcService.importlmgzsdsgpsf(file,proname,username);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcLmgzsdsgpsfJgfc jjgFbgcLmgcLmgzsdsgpsf) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcLmgzsdsgpsfJgfc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLmgzsdsgpsf != null){
            QueryWrapper<JjgFbgcLmgcLmgzsdsgpsfJgfc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLmgzsdsgpsf.getProname());

            Date jcsj = jjgFbgcLmgcLmgzsdsgpsf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcLmgzsdsgpsf.getZh())) {
                wrapper.like("zh", jjgFbgcLmgcLmgzsdsgpsf.getZh());
            }

            //调用方法分页查询
            IPage<JjgFbgcLmgcLmgzsdsgpsfJgfc> pageModel = jjgFbgcLmgcLmgzsdsgpsfJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        QueryWrapper<JjgFbgcLmgcLmgzsdsgpsfJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);

        boolean remove = jjgFbgcLmgcLmgzsdsgpsfJgfcService.remove(queryWrapper);
        return remove?Result.ok():Result.fail().message("删除失败！");


    }

    @ApiOperation("批量删构造深度手工铺沙法数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLmgzsdsgpsfJgfcService.removeByIds(idList);
        return hd?Result.ok():Result.fail().message("删除失败！");

    }

     @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcLmgcLmgzsdsgpsfJgfc> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcLmgcLmgzsdsgpsfJgfcService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

