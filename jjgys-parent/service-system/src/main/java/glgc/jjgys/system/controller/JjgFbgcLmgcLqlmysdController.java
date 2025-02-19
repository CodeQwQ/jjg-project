package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcLmgcLqlmysdService;
import glgc.jjgys.system.service.SysUserService;
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
 * @since 2023-04-10
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/lqlmysd")
public class JjgFbgcLmgcLqlmysdController {

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Autowired
    private SysUserService sysUserService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd,String userid) throws IOException {
        String fileName1 = "12沥青路面压实度-(不分离上面层).xlsx";
        String fileName2 = "12沥青路面压实度-(分离上面层).xlsx";
        String p1 = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName1;
        String p2 = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName2;
        File file1 = new File(p1);
        File file2 = new File(p2);
        if (file1.exists() && file2.exists()){
            List listname = new ArrayList<>();
            //拿到部门下所有用户的id
            List<Long> idlist = new ArrayList<>();
            if ("1".equals(userid)){
            }else {
                //查询用户类型
                QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
                wrapperuser.eq("id",userid);
                SysUser one = sysUserService.getOne(wrapperuser);
                String type = one.getType();
                if ("2".equals(type) || "4".equals(type)){
                    //公司管理员
                    Long deptId = one.getDeptId();
                    QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
                    wrapperid.eq("dept_id",deptId);
                    List<SysUser> list = sysUserService.list(wrapperid);


                    if (list!=null){
                        for (SysUser user : list) {
                            Long id = user.getId();
                            idlist.add(id);
                        }
                    }
                }else if ("3".equals(type)){
                    //普通用户
                    idlist.add(Long.valueOf(userid));
                }
            }
            List<Map<String, Object>> a = jjgFbgcLmgcLqlmysdService.selectsfl(proname,htd,idlist);
            if (a.size() == 1){
                String sffl = a.get(0).get("sffl").toString();
                if ("否".equals(sffl)){
                    listname.add("(不分离上面层)");
                }else {
                    listname.add("(分离上面层)");
                }
            }else {
                listname.add("(不分离上面层)");
                listname.add("(分离上面层)");
            }
            String zipName = "12沥青路面压实度";
            JjgFbgcCommonUtils.batchDownloadlmhpFile(request,response,zipName,listname,filespath+File.separator+proname+File.separator+htd);
        }else if (file1.exists()){
            JjgFbgcCommonUtils.download(response,p1,fileName1);
        }else if (file2.exists()){
            JjgFbgcCommonUtils.download(response,p2,fileName2);
        }
    }

    @ApiOperation("生成沥青路面压实度鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcLmgcLqlmysdService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看沥青路面压实度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("沥青路面压实度模板文件导出")
    @GetMapping("exportlqlmysd")
    public void exportlqlmysd(HttpServletResponse response) {
        jjgFbgcLmgcLqlmysdService.exportlqlmysd(response);
    }

    @ApiOperation(value = "沥青路面压实度数据文件导入")
    @PostMapping("importlqlmysd")
    public Result importlqlmysd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLqlmysdService.importlqlmysd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPageHt/{current}/{limit}")
    public Result findQueryPageHt(@PathVariable long current,
                                  @PathVariable long limit,
                                  @RequestBody JjgFbgcLmgcLqlmysd jjgFbgcLmgcLqlmysd) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcLqlmysd> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLqlmysd != null){
            QueryWrapper<JjgFbgcLmgcLqlmysd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLqlmysd.getProname());
            wrapper.like("htd",jjgFbgcLmgcLqlmysd.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcLqlmysd.getFbgc());
            String userid = jjgFbgcLmgcLqlmysd.getUserid();
            if ("1".equals(userid)){

            }else {
                //查询用户类型
                QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
                wrapperuser.eq("id",userid);
                SysUser one = sysUserService.getOne(wrapperuser);
                String type = one.getType();
                if ("2".equals(type) || "4".equals(type) || "4".equals(type)){
                    //公司管理员
                    Long deptId = one.getDeptId();
                    QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
                    wrapperid.eq("dept_id",deptId);
                    List<SysUser> list = sysUserService.list(wrapperid);
                    //拿到部门下所有用户的id
                    List<Long> idlist = new ArrayList<>();

                    if (list!=null){
                        for (SysUser user : list) {
                            Long id = user.getId();
                            idlist.add(id);
                        }
                    }
                    wrapper.in("userid",idlist);
                }else if ("3".equals(type)){
                    //普通用户
                    //String username = project.getUsername();
                    wrapper.like("userid",userid);
                }
            }
            Date jcsj = jjgFbgcLmgcLqlmysd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcLqlmysd.getZh())) {
                wrapper.like("zh", jjgFbgcLmgcLqlmysd.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLqlmysd> pageModel = jjgFbgcLmgcLqlmysdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量沥青路面压实度数据")
    @DeleteMapping("removeBeatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLqlmysdService.removeByIds(idList);
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
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgFbgcLmgcLqlmysd> queryWrapper = new QueryWrapper<>();
        queryWrapper.eq("proname",proname);
        queryWrapper.eq("htd",htd);
        String userid = commonInfoVo.getUserid();
        if ("1".equals(userid)){
        }else {
            //查询用户类型
            QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
            wrapperuser.eq("id",userid);
            SysUser one = sysUserService.getOne(wrapperuser);
            String type = one.getType();
            if ("2".equals(type) || "4".equals(type)){
                //公司管理员
                Long deptId = one.getDeptId();
                QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
                wrapperid.eq("dept_id",deptId);
                List<SysUser> list = sysUserService.list(wrapperid);
                //拿到部门下所有用户的id
                List<Long> idlist = new ArrayList<>();

                if (list!=null){
                    for (SysUser user : list) {
                        Long id = user.getId();
                        idlist.add(id);

                    }
                }
                queryWrapper.in("userid",idlist);
            }else if ("3".equals(type)){
                //普通用户
                //String username = project.getUsername();
                queryWrapper.like("userid",userid);

            }
            boolean remove = jjgFbgcLmgcLqlmysdService.remove(queryWrapper);
            if (remove){
                File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"12沥青路面压实度-(不分离上面层).xlsx");
                File f1 = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"12沥青路面压实度-(分离上面层).xlsx");
                if (f.exists()){
                    f.delete();
                }
                if (f1.exists()){
                    f1.delete();
                }
            }
        }
        return Result.ok();

    }

    @ApiOperation("根据id查询")
    @GetMapping("getysd/{id}")
    public Result getysd(@PathVariable String id) {
        JjgFbgcLmgcLqlmysd user = jjgFbgcLmgcLqlmysdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改沥青路面压实度数据")
    @PostMapping("updateysd")
    public Result updateysd(@RequestBody JjgFbgcLmgcLqlmysd ysd) {
        boolean is_Success = jjgFbgcLmgcLqlmysdService.updateById(ysd);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcLmgcLqlmysd> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcLmgcLqlmysdService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

