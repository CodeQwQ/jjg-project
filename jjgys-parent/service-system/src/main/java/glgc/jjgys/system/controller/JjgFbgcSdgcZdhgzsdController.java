package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcSdgcZdhgzsd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcSdgcZdhgzsdService;
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
 * @since 2023-10-23
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/zdhgzsd")
public class JjgFbgcSdgcZdhgzsdController {

    @Autowired
    private JjgFbgcSdgcZdhgzsdService jjgFbgcSdgcZdhgzsdService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd,String userid) throws IOException {
        List<Map<String,Object>> lxlist = jjgFbgcSdgcZdhgzsdService.selectlx(proname,htd,userid);
        List<String> fileName = new ArrayList<>();
        for (Map<String, Object> map : lxlist) {
            String lxbs = map.get("lxbs").toString();
            fileName.add("51隧道路面构造深度-"+lxbs);
        }
        String zipname = "隧道路面构造深度鉴定表";
        JjgFbgcCommonUtils.batchDowndFile(response,zipname,fileName,filespath+ File.separator+proname+File.separator+htd);
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgFbgcSdgcZdhgzsd> queryWrapper = new QueryWrapper<>();
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
            List<Map<String,Object>> lxlist = jjgFbgcSdgcZdhgzsdService.selectlx(proname,htd,userid);
            List<String> list = new ArrayList<>();
            for (Map<String, Object> map : lxlist) {
                String lxbs = map.get("lxbs").toString();
                list.add("51隧道路面构造深度-"+lxbs+".xlsx");
            }
            boolean remove = jjgFbgcSdgcZdhgzsdService.remove(queryWrapper);
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

    @ApiOperation("生成隧道构造深度鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcSdgcZdhgzsdService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看隧道构造深度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcZdhgzsdService.lookJdbjgPage(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道构造深度模板文件导出")
    @GetMapping("exportgzsd")
    public void exportgzsd(HttpServletResponse response, @RequestParam String cd) throws IOException {
        jjgFbgcSdgcZdhgzsdService.exportgzsd(response,cd);
    }


    @ApiOperation(value = "隧道构造深度数据文件导入")
    @PostMapping("importgzsd")
    public Result importgzsd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgFbgcSdgcZdhgzsdService.importgzsd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcZdhgzsd jjgZdhGzsd) {
        //创建page对象
        Page<JjgFbgcSdgcZdhgzsd> pageParam = new Page<>(current, limit);
        if (jjgZdhGzsd != null) {
            QueryWrapper<JjgFbgcSdgcZdhgzsd> wrapper = new QueryWrapper<>();
            wrapper.eq("proname", jjgZdhGzsd.getProname());
            wrapper.eq("htd", jjgZdhGzsd.getHtd());

            String userid = jjgZdhGzsd.getUserid();
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
            if (!StringUtils.isEmpty(jjgZdhGzsd.getQdzh())) {
                wrapper.like("qdzh", jjgZdhGzsd.getQdzh());
            }
            if (!StringUtils.isEmpty(jjgZdhGzsd.getZdzh())) {
                wrapper.like("zdzh", jjgZdhGzsd.getZdzh());
            }
            if (!StringUtils.isEmpty(jjgZdhGzsd.getLxbs())) {
                wrapper.like("lxbs", jjgZdhGzsd.getLxbs());
            }

            //调用方法分页查询
            IPage<JjgFbgcSdgcZdhgzsd> pageModel = jjgFbgcSdgcZdhgzsdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }


    @ApiOperation("批量删除隧道构造深度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcZdhgzsdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getmcxs/{id}")
    public Result getmcxs(@PathVariable String id) {
        JjgFbgcSdgcZdhgzsd user = jjgFbgcSdgcZdhgzsdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道构造深度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcZdhgzsd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcZdhgzsdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setTitle("隧道构造深度");
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

