package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcSdgcSdhpService;
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
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-27
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/sdhp")
public class JjgFbgcSdgcSdhpController {

    @Autowired
    private JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd,String fbgc,String userid) throws IOException {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhpService.selectsdmc(proname,htd,fbgc,userid);
        List list = new ArrayList<>();
        for (int i=0;i<sdmclist.size();i++){
            list.add(sdmclist.get(i).get("sdmc"));
        }
        String zipName = "55隧道横坡";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }

    @ApiOperation("生成隧道横坡鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcSdgcSdhpService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看隧道横坡鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcSdhpService.lookJdbjgPage(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgFbgcSdgcSdhp> queryWrapper = new QueryWrapper<>();
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
            List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhpService.selectsdmc(proname,htd,"隧道工程",userid);
            List<String> list = new ArrayList<>();
            for (int i=0;i<sdmclist.size();i++){
                list.add("55隧道横坡-"+sdmclist.get(i).get("sdmc")+".xlsx");
            }
            boolean remove = jjgFbgcSdgcSdhpService.remove(queryWrapper);
            if (remove){
                if (list != null && list.size()>0){
                    for (String s : list) {
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

    @ApiOperation("隧道横坡模板文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response){
        jjgFbgcSdgcSdhpService.export(response);
    }


    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "隧道横坡数据文件导入")
    @PostMapping("importsdhp")
    public Result importsdhp(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcSdhpService.importsdhp(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcSdhp jjgFbgcSdgcSdhp) throws ParseException {
        //创建page对象
        Page<JjgFbgcSdgcSdhp> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcSdhp != null){
            QueryWrapper<JjgFbgcSdgcSdhp> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcSdhp.getProname());
            wrapper.like("htd",jjgFbgcSdgcSdhp.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcSdhp.getFbgc());
            String userid = jjgFbgcSdgcSdhp.getUserid();
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
            Date jcsj = jjgFbgcSdgcSdhp.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcSdgcSdhp.getSdmc())) {
                wrapper.like("sdmc", jjgFbgcSdgcSdhp.getSdmc());
            }
            if (!StringUtils.isEmpty(jjgFbgcSdgcSdhp.getZh())) {
                wrapper.like("zh", jjgFbgcSdgcSdhp.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcSdhp> pageModel = jjgFbgcSdgcSdhpService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除隧道横坡数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcSdgcSdhpService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSdhp{id}")
    public Result getSdhp(@PathVariable String id) {
        JjgFbgcSdgcSdhp user = jjgFbgcSdgcSdhpService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道横坡数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcSdhp user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcSdhpService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("隧道横坡数据");
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
    public Result createManyRecords(@RequestBody List<JjgFbgcSdgcSdhp> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcSdgcSdhpService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }


}

