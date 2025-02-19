package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcSdgcLmgzsdsgpsf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcSdgcLmgzsdsgpsfService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
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
 * @since 2023-05-03
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/lmgzsdsgpsf")
public class JjgFbgcSdgcLmgzsdsgpsfController {

    @Autowired
    private JjgFbgcSdgcLmgzsdsgpsfService jjgFbgcSdgcLmgzsdsgpsfService;

    @Autowired
    private SysUserService sysUserService;


    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request,HttpServletResponse response, String proname, String htd,String fbgc,String userid) throws IOException {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcLmgzsdsgpsfService.selectsdmc(proname,htd,fbgc,userid);
        List list = new ArrayList<>();
        for (int i=0;i<sdmclist.size();i++){
            list.add(sdmclist.get(i).get("sdmc"));
        }
        String zipName = "51构造深度手工铺沙法";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }


    @ApiOperation("生成构造深度手工铺沙法鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcSdgcLmgzsdsgpsfService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }
    @ApiOperation("查看构造深度手工铺沙法鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcLmgzsdsgpsfService.lookJdbjgPage(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("构造深度手工铺沙法模板文件导出")
    @GetMapping("exportlmgzsdsgpsf")
    public void exportlmgzsdsgpsf(HttpServletResponse response){
        jjgFbgcSdgcLmgzsdsgpsfService.exportlmgzsdsgpsf(response);
    }


    @ApiOperation(value = "构造深度手工铺沙法数据文件导入")
    @PostMapping("importlmgzsdsgpsf")
    public Result importlmgzsdsgpsf(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcLmgzsdsgpsfService.importlmgzsdsgpsf(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcLmgzsdsgpsf jjgFbgcSdgcLmgzsdsgpsf) throws ParseException {
        //创建page对象
        Page<JjgFbgcSdgcLmgzsdsgpsf> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcLmgzsdsgpsf != null){
            QueryWrapper<JjgFbgcSdgcLmgzsdsgpsf> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcLmgzsdsgpsf.getProname());
            wrapper.like("htd",jjgFbgcSdgcLmgzsdsgpsf.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcLmgzsdsgpsf.getFbgc());
            String userid = jjgFbgcSdgcLmgzsdsgpsf.getUserid();
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
            Date jcsj = jjgFbgcSdgcLmgzsdsgpsf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcSdgcLmgzsdsgpsf.getSdmc())) {
                wrapper.like("sdmc", jjgFbgcSdgcLmgzsdsgpsf.getSdmc());
            }
            if (!StringUtils.isEmpty(jjgFbgcSdgcLmgzsdsgpsf.getZh())) {
                wrapper.like("zh", jjgFbgcSdgcLmgzsdsgpsf.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcLmgzsdsgpsf> pageModel = jjgFbgcSdgcLmgzsdsgpsfService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删构造深度手工铺沙法数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcLmgzsdsgpsfService.removeByIds(idList);
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
        QueryWrapper<JjgFbgcSdgcLmgzsdsgpsf> queryWrapper = new QueryWrapper<>();
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
            List<Map<String,Object>> sdmclist = jjgFbgcSdgcLmgzsdsgpsfService.selectsdmc(proname,htd,"隧道工程",userid);
            List<String> list = new ArrayList<>();
            for (int i=0;i<sdmclist.size();i++){
                list.add("51构造深度手工铺沙法-"+sdmclist.get(i).get("sdmc")+".xlsx");
            }
            boolean remove = jjgFbgcSdgcLmgzsdsgpsfService.remove(queryWrapper);
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


    @ApiOperation("根据id查询")
    @GetMapping("getLmhp/{id}")
    public Result getLmhp(@PathVariable String id) {
        JjgFbgcSdgcLmgzsdsgpsf user = jjgFbgcSdgcLmgzsdsgpsfService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥面构造深度手工铺沙法数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcLmgzsdsgpsf user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcLmgzsdsgpsfService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("桥面构造深度手工铺沙法数据");
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
    public Result createManyRecords(@RequestBody List<JjgFbgcSdgcLmgzsdsgpsf> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcSdgcLmgzsdsgpsfService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

