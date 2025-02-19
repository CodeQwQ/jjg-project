package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.base.JgCommonEntity;
import glgc.jjgys.model.project.JjgZdhMcxsJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgZdhMcxsJgfcService;
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
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-09-23
 */
@RestController
@RequestMapping("/jjg/jgfc/zdh/mcxs")
@CrossOrigin
public class JjgZdhMcxsJgfcController {

    @Autowired
    private JjgZdhMcxsJgfcService jjgZdhMcxsJgfcService;

    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @ApiOperation("查看平均值")
    @GetMapping("lookpjz")
    public Result lookpjz(@RequestParam String proname) throws IOException {
        List<Map<String,Object>> list = jjgZdhMcxsJgfcService.lookpjz(proname);
        return Result.ok(list);
    }

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname) throws IOException {
        List<Map<String,Object>> htdList = jjgZdhMcxsJgfcService.selecthtd(proname);
        List<String> fileName = new ArrayList<>();
        if (htdList!=null){
            for (Map<String, Object> map1 : htdList) {
                String htd = map1.get("htd").toString();
                String lxbs = map1.get("lxbs").toString();
                if (lxbs.equals("主线")){
                    fileName.add(htd+File.separator+"19路面摩擦系数");
                }else {
                    fileName.add(htd+File.separator+"62互通摩擦系数-"+lxbs);
                }
                /*List<Map<String,Object>> lxlist = jjgZdhMcxsJgfcService.selectlx(proname,htd);
                for (Map<String, Object> map : lxlist) {
                    String lxbs = map.get("lxbs").toString();
                    if (lxbs.equals("主线")){
                        fileName.add(htd+File.separator+"19路面摩擦系数");
                    }else {
                        fileName.add(htd+File.separator+"62互通摩擦系数-"+lxbs);
                    }
                }*/
            }
        }
        String zipname = "摩擦系数鉴定表";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);
    }

    @ApiOperation("生成摩擦系数鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody JgCommonEntity jgCommon) throws Exception {
        boolean is_Success = jjgZdhMcxsJgfcService.generateJdb(jgCommon);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");
    }



    @ApiOperation("摩擦系数模板文件导出")
    @GetMapping("exportmcxs")
    public void exportmcxs(HttpServletResponse response, @RequestParam String cd) throws IOException {
        jjgZdhMcxsJgfcService.exportmcxs(response,cd);
    }

    @ApiOperation(value = "摩擦系数数据文件导入")
    @PostMapping("importmcxs")
    public Result importmcxs(@RequestParam("file") MultipartFile file, String proname,String username) throws IOException {
        jjgZdhMcxsJgfcService.importmcxs(file,proname,username);
        return Result.ok();
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        QueryWrapper<JjgZdhMcxsJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);
        boolean remove = jjgZdhMcxsJgfcService.remove(queryWrapper);
        return remove ? Result.ok():Result.fail().message("删除失败！");


    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgZdhMcxsJgfc jjgZdhMcxs) {
        //创建page对象
        Page<JjgZdhMcxsJgfc> pageParam = new Page<>(current, limit);
        if (jjgZdhMcxs != null) {
            QueryWrapper<JjgZdhMcxsJgfc> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgZdhMcxs.getProname());

            if (!StringUtils.isEmpty(jjgZdhMcxs.getQdzh())) {
                wrapper.eq("qdzh", jjgZdhMcxs.getQdzh());
            }
            if (!StringUtils.isEmpty(jjgZdhMcxs.getZdzh())) {
                wrapper.eq("zdzh", jjgZdhMcxs.getZdzh());
            }
            if (!StringUtils.isEmpty(jjgZdhMcxs.getLxbs())) {
                wrapper.like("lxbs", jjgZdhMcxs.getLxbs());
            }

            //调用方法分页查询
            IPage<JjgZdhMcxsJgfc> pageModel = jjgZdhMcxsJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除摩擦系数数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgZdhMcxsJgfcService.removeByIds(idList);
        return hd ? Result.ok():Result.fail().message("删除失败！");

    }

}

