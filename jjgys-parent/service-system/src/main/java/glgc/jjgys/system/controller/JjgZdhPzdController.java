package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgZdhPzd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgZdhPzdService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.service.SysUserService;
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
 * @since 2023-06-17
 */
@RestController
@RequestMapping("/jjg/zdh/pzd")
public class JjgZdhPzdController {

    @Autowired
    private JjgZdhPzdService jjgZdhPzdService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd,String userid) throws IOException {
        List<Map<String,Object>> lxlist = jjgZdhPzdService.selectlx(proname,htd,userid);
        List<String> fileName = new ArrayList<>();
        for (Map<String, Object> map : lxlist) {
            String lxbs = map.get("lxbs").toString();
            if (lxbs.equals("主线")){
                fileName.add("18路面平整度");;
            }else {
                fileName.add("61互通平整度-"+lxbs);
            }
        }
        String zipname = "平整度鉴定表";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,filespath+File.separator+proname+File.separator+htd);
    }

    /**
     * 沥青隧道路面平整度规定值
     * 沥青桥面平整度规定值
     * 沥青路面平整度规定值
     *
     * 混凝土隧道路面平整度规定值
     * 混凝土桥面平整度规定值
     * 混凝土路面平整度规定值
     *
     */

    @ApiOperation("生成平整度鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgZdhPzdService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgZdhPzd> queryWrapper = new QueryWrapper<>();
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
                queryWrapper.in("userid",idlist);
            }else if ("3".equals(type)){
                //普通用户
                //String username = project.getUsername();
                queryWrapper.like("userid",userid);

            }
            List<Map<String,Object>> lxlist = jjgZdhPzdService.selectlx(proname,htd,userid);
            List<String> fileName = new ArrayList<>();
            for (Map<String, Object> map : lxlist) {
                String lxbs = map.get("lxbs").toString();
                if (lxbs.equals("主线")){
                    fileName.add("18路面平整度.xlsx");;
                }else {
                    fileName.add("61互通平整度-"+lxbs+".xlsx");
                }
            }
            boolean remove = jjgZdhPzdService.remove(queryWrapper);
            if (remove){
                if (fileName != null && fileName.size()>0){
                    for (String s : fileName) {
                        File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+s);
                        if (f.exists()){
                            f.delete();
                        }
                    }
                }
            }
        }
        return Result.ok();

    }

    @ApiOperation("查看平整度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgZdhPzdService.lookJdbjgPage(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("平整度模板文件导出")
    @GetMapping("exportpzd")
    public void exportpzd(HttpServletResponse response,@RequestParam String cd) throws IOException {
        jjgZdhPzdService.exportpzd(response,cd);
    }


    @ApiOperation(value = "平整度数据文件导入")
    @PostMapping("importpzd")
    public Result importpzd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgZdhPzdService.importpzd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgZdhPzd jjgZdhpzd) {
        //创建page对象
        Page<JjgZdhPzd> pageParam = new Page<>(current, limit);
        if (jjgZdhpzd != null) {
            QueryWrapper<JjgZdhPzd> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgZdhpzd.getProname());
            wrapper.like("htd", jjgZdhpzd.getHtd());

            String userid = jjgZdhpzd.getUserid();
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
                    wrapper.in("userid",idlist);
                }else if ("3".equals(type)){
                    //普通用户
                    //String username = project.getUsername();
                    wrapper.like("userid",userid);
                }
            }
            if (!StringUtils.isEmpty(jjgZdhpzd.getQdzh())) {
                wrapper.eq("qdzh", jjgZdhpzd.getQdzh());
            }
            if (!StringUtils.isEmpty(jjgZdhpzd.getZdzh())) {
                wrapper.eq("zdzh", jjgZdhpzd.getZdzh());
            }

            //调用方法分页查询
            IPage<JjgZdhPzd> pageModel = jjgZdhPzdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgZdhPzdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getpzd/{id}")
    public Result getpzd(@PathVariable String id) {
        JjgZdhPzd user = jjgZdhPzdService.getById(id);
        return Result.ok(user);
    }

    @PostMapping("update")
    public Result update(@RequestBody JjgZdhPzd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgZdhPzdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setTitle("平整度");
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

