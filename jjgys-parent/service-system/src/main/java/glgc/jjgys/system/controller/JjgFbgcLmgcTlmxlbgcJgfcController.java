package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcTlmxlbgcJgfc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.service.JjgFbgcLmgcTlmxlbgcJgfcService;
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
@RequestMapping("/jjg/jgfc/tlmxlbgc")
public class JjgFbgcLmgcTlmxlbgcJgfcController {

    @Autowired
    private JjgFbgcLmgcTlmxlbgcJgfcService jjgFbgcLmgcTlmxlbgcJgfcService;


    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname) throws IOException {
        List<Map<String,Object>> htdList = jjgFbgcLmgcTlmxlbgcJgfcService.selecthtd(proname);
        List<String> fileName = new ArrayList<>();
        if (htdList!=null){
            for (Map<String, Object> map1 : htdList) {
                String htd = map1.get("htd").toString();
                fileName.add(htd+File.separator+"17混凝土路面相邻板高差");
            }
        }
        String zipname = "混凝土路面相邻板高差";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,jgfilepath+ File.separator+proname);
    }

    @ApiOperation("生成相邻板高差鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        boolean is_Success = jjgFbgcLmgcTlmxlbgcJgfcService.generateJdb(proname);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("相邻板高差模板文件导出")
    @GetMapping("exportxlbgs")
    public void exportxlbgs(HttpServletResponse response){
        jjgFbgcLmgcTlmxlbgcJgfcService.exportxlbgs(response);
    }


    @ApiOperation(value = "相邻板高差数据文件导入")
    @PostMapping("importxlbgs")
    public Result importxlbgs(@RequestParam("file") MultipartFile file, String proname) {
        jjgFbgcLmgcTlmxlbgcJgfcService.importxlbgs(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcTlmxlbgcJgfc jjgFbgcLmgcTlmxlbgc) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcTlmxlbgcJgfc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcTlmxlbgc != null){
            QueryWrapper<JjgFbgcLmgcTlmxlbgcJgfc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcTlmxlbgc.getProname());

            /*String username = jjgFbgcLmgcTlmxlbgc.getUsername();
            QueryWrapper<SysUser> sysUserQueryWrapper = new QueryWrapper<>();
            sysUserQueryWrapper.eq("username", username);
            SysUser one = sysUserService.getOne(sysUserQueryWrapper);
            String userid = one.getId().toString();

            QueryWrapper<SysUserRole> sysUserRoleQueryWrapper = new QueryWrapper<>();
            sysUserRoleQueryWrapper.eq("user_id", userid);
            SysUserRole sysUserRole = sysUserRoleMapper.selectOne(sysUserRoleQueryWrapper);
            String roleId = sysUserRole.getRoleId();

            QueryWrapper<SysRole> sysRoleQueryWrapper = new QueryWrapper<>();
            sysRoleQueryWrapper.eq("id", roleId);
            SysRole role = sysRoleService.getOne(sysRoleQueryWrapper);
            String rolecode = role.getRoleCode();

            if (rolecode.equals("YH")){
                wrapper.eq("username", username);
            }*/
            Date jcsj = jjgFbgcLmgcTlmxlbgc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcTlmxlbgc.getZh())) {
                wrapper.like("zh", jjgFbgcLmgcTlmxlbgc.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcTlmxlbgcJgfc> pageModel = jjgFbgcLmgcTlmxlbgcJgfcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        QueryWrapper<JjgFbgcLmgcTlmxlbgcJgfc> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);
        boolean remove = jjgFbgcLmgcTlmxlbgcJgfcService.remove(queryWrapper);
        return remove?Result.ok():Result.fail().message("删除失败！");

    }


    @ApiOperation("批量删除相邻板高差数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcTlmxlbgcJgfcService.removeByIds(idList);
        return hd?Result.ok():Result.fail().message("删除失败！");
    }

     @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcLmgcTlmxlbgcJgfc> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcLmgcTlmxlbgcJgfcService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

