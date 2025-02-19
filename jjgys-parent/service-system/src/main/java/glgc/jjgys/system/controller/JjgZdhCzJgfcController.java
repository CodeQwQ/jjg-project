package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgZdhCzJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgZdhCzJgfcService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

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
@RequestMapping("/jjg/jgfc/zdh/cz")
public class JjgZdhCzJgfcController {

    @Autowired
    private JjgZdhCzJgfcService jjgZdhCzJgfcService;


    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @ApiOperation("查看平均值")
    @GetMapping("lookpjz")
    public Result lookpjz(@RequestParam String proname) throws IOException {
        List<Map<String,Object>> list = jjgZdhCzJgfcService.lookpjz(proname);
        return Result.ok(list);
    }

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname) throws IOException {
        List<Map<String,Object>> htdList = jjgZdhCzJgfcService.selecthtd(proname);
        List<String> fileName = new ArrayList<>();
        if (htdList!=null){
            for (Map<String, Object> map1 : htdList) {
                String htd = map1.get("htd").toString();
                String lxbs = map1.get("lxbs").toString();
                if (lxbs.equals("主线")){
                    fileName.add(htd+File.separator+"14路面车辙");
                }else {
                    fileName.add(htd+File.separator+"69互通车辙-"+lxbs);
                }

                /*List<Map<String,Object>> lxlist = jjgZdhCzJgfcService.selectlx(proname,htd);
                for (Map<String, Object> map : lxlist) {
                    String lxbs = map.get("lxbs").toString();
                    if (lxbs.equals("主线")){
                        fileName.add(htd+File.separator+"14路面车辙");
                    }else {
                        fileName.add(htd+File.separator+"69互通车辙-"+lxbs);
                    }
                }*/
            }
        }
        String zipname = "车辙鉴定表";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        QueryWrapper<JjgZdhCzJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);

        boolean remove = jjgZdhCzJgfcService.remove(queryWrapper);
        if(remove){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("生成车辙鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        String sjz = commonInfoVo.getSjz();
        boolean is_Success = jjgZdhCzJgfcService.generateJdb(proname,sjz);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("车辙模板文件导出")
    @GetMapping("exportcz")
    public void exportcz(HttpServletResponse response, @RequestParam String cd) throws IOException {
        jjgZdhCzJgfcService.exportcz(response,cd);
    }


    @ApiOperation(value = "车辙数据文件导入")
    @PostMapping("importcz")
    public Result importcz(@RequestParam("file") MultipartFile file,String proname,String username) throws IOException {
        jjgZdhCzJgfcService.importcz(file,proname,username);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgZdhCzJgfc jjgZdhCz) {
        //创建page对象
        Page<JjgZdhCzJgfc> pageParam = new Page<>(current, limit);
        if (jjgZdhCz != null) {
            QueryWrapper<JjgZdhCzJgfc> wrapper = new QueryWrapper<>();
            wrapper.eq("proname", jjgZdhCz.getProname());
            if (!StringUtils.isEmpty(jjgZdhCz.getQdzh())) {
                wrapper.eq("qdzh", jjgZdhCz.getQdzh());
            }
            if (!StringUtils.isEmpty(jjgZdhCz.getZdzh())) {
                wrapper.eq("zdzh", jjgZdhCz.getZdzh());
            }

            //调用方法分页查询
            IPage<JjgZdhCzJgfc> pageModel = jjgZdhCzJgfcService.page(pageParam, wrapper);
            System.out.println(pageModel);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除竣工车辙数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgZdhCzJgfcService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

