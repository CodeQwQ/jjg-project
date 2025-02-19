package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcGslqlmhdzxf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcLmgcGslqlmhdzxfService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
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
 * @since 2023-04-25
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/gslqlmhdzxf")
public class JjgFbgcLmgcGslqlmhdzxfController {

    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;

    @Autowired
    private SysUserService sysUserService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "22沥青路面厚度-钻芯法.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }


    @ApiOperation("生成沥青路面厚度鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success =jjgFbgcLmgcGslqlmhdzxfService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }
    @ApiOperation("查看沥青路面厚度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("沥青路面厚度模板文件导出")
    @GetMapping("exportLqlmhd")
    public void exportlqlmhd(HttpServletResponse response){
        jjgFbgcLmgcGslqlmhdzxfService.exportlqlmhd(response);
    }

    @ApiOperation(value = "沥青路面厚度数据文件导入")
    @PostMapping("importLqlmhd")
    public Result importLmssxs(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcGslqlmhdzxfService.importLqlmhd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcGslqlmhdzxf jjgFbgcLmgcGslqlmhdzxf) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcGslqlmhdzxf> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcGslqlmhdzxf != null){
            QueryWrapper<JjgFbgcLmgcGslqlmhdzxf> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcGslqlmhdzxf.getProname());
            wrapper.like("htd",jjgFbgcLmgcGslqlmhdzxf.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcGslqlmhdzxf.getFbgc());
            String userid = jjgFbgcLmgcGslqlmhdzxf.getUserid();
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
            Date jcsj = jjgFbgcLmgcGslqlmhdzxf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcGslqlmhdzxf.getZh())) {
                wrapper.like("zh", jjgFbgcLmgcGslqlmhdzxf.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcGslqlmhdzxf> pageModel = jjgFbgcLmgcGslqlmhdzxfService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除沥青路面厚度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcGslqlmhdzxfService.removeByIds(idList);
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
        QueryWrapper<JjgFbgcLmgcGslqlmhdzxf> queryWrapper = new QueryWrapper<>();
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
            boolean remove = jjgFbgcLmgcGslqlmhdzxfService.remove(queryWrapper);
            if (remove){
                File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"22沥青路面厚度-钻芯法.xlsx");
                if (f.exists()){
                    f.delete();
                }
            }
        }
        return Result.ok();

    }

    @ApiOperation("根据id查询")
    @GetMapping("getLmssxs/{id}")
    public Result getLmssxs(@PathVariable String id) {
        JjgFbgcLmgcGslqlmhdzxf user = jjgFbgcLmgcGslqlmhdzxfService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改沥青路面厚度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcGslqlmhdzxf user) {
        boolean is_Success = jjgFbgcLmgcGslqlmhdzxfService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }
    
     @PostMapping("createManyRecords")
        public Result createManyRecords(@RequestBody List<JjgFbgcLmgcGslqlmhdzxf> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcLmgcGslqlmhdzxfService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }
}

