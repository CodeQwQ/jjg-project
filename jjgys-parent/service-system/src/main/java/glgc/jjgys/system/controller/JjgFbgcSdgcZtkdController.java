package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcSdgcZtkd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcSdgcZtkdService;
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
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  隧道总体宽度
 * </p>
 *
 * @author wq
 * @since 2023-03-26
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/ztkd")
@CrossOrigin
public class JjgFbgcSdgcZtkdController {

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "41隧道总体宽度.xlsx";
        String p = filespath+"/"+proname+"/"+htd+"/"+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成隧道总体宽度鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcSdgcZtkdService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgFbgcSdgcZtkd> queryWrapper = new QueryWrapper<>();
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
            boolean remove = jjgFbgcSdgcZtkdService.remove(queryWrapper);
            if (remove){
                File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"41隧道总体宽度.xlsx");
                if (f.exists()){
                    f.delete();
                }
            }
        }
        return Result.ok();

    }

    @ApiOperation("查看隧道总体宽度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcZtkdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道总体宽度模板文件导出")
    @GetMapping("exportsdztkd")
    public void exportsdztkd(HttpServletResponse response){
        jjgFbgcSdgcZtkdService.exportsdztkd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "隧道总体宽度数据文件导入")
    @PostMapping("importsdztkd")
    public Result importsdztkd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcZtkdService.importsdztkd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcZtkd jjgFbgcSdgcZtkd) throws ParseException {
        //创建page对象
        Page<JjgFbgcSdgcZtkd> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcZtkd != null){
            QueryWrapper<JjgFbgcSdgcZtkd> wrapper=new QueryWrapper<>();
            wrapper.eq("proname",jjgFbgcSdgcZtkd.getProname());
            wrapper.eq("htd",jjgFbgcSdgcZtkd.getHtd());
            wrapper.eq("fbgc",jjgFbgcSdgcZtkd.getFbgc());
            String userid = jjgFbgcSdgcZtkd.getUserid();
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
            if (!StringUtils.isEmpty(jjgFbgcSdgcZtkd.getJcsj())) {
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jjgFbgcSdgcZtkd.getJcsj().toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj", jjgFbgcSdgcZtkd.getJcsj());
            }
            if (!StringUtils.isEmpty(jjgFbgcSdgcZtkd.getSdmc())) {
                wrapper.like("sdmc", jjgFbgcSdgcZtkd.getSdmc());
            }
            if (!StringUtils.isEmpty(jjgFbgcSdgcZtkd.getZh())) {
                wrapper.like("zh", jjgFbgcSdgcZtkd.getZh());
            }

            //调用方法分页查询
            IPage<JjgFbgcSdgcZtkd> pageModel = jjgFbgcSdgcZtkdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除隧道总体宽度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcZtkdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSdhp{id}")
    public Result getSdhp(@PathVariable String id) {
        JjgFbgcSdgcZtkd user = jjgFbgcSdgcZtkdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道总体宽度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcZtkd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcZtkdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("隧道总体宽度数据");
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

    @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcSdgcZtkd> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcSdgcZtkdService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

