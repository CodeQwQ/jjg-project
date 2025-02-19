package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhl;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxfhlService;
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
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jabxfhl")
@CrossOrigin
public class JjgFbgcJtaqssJabxfhlController {

    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "58交安钢防护栏.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成交安波形防护栏鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcJtaqssJabxfhlService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看交安波形防护栏鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJabxfhlService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安波形防护模板文件导出")
    @GetMapping("exportjabxfhl")
    public void exportjabxfhl(HttpServletResponse response) throws IOException {
        jjgFbgcJtaqssJabxfhlService.exportjabxfhl(response);
    }


    @ApiOperation(value = "交安波形防护数据文件导入")
    @PostMapping("importjabxfhl")
    public Result importjabxfhl(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgFbgcJtaqssJabxfhlService.importjabxfhl(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabxfhl jjgFbgcJtaqssJabxfhl) {
        //创建page对象
        Page<JjgFbgcJtaqssJabxfhl> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabxfhl != null) {
            QueryWrapper<JjgFbgcJtaqssJabxfhl> wrapper = new QueryWrapper<>();
            wrapper.eq("proname", jjgFbgcJtaqssJabxfhl.getProname());
            wrapper.eq("htd", jjgFbgcJtaqssJabxfhl.getHtd());
            wrapper.eq("fbgc", jjgFbgcJtaqssJabxfhl.getFbgc());

            String userid = jjgFbgcJtaqssJabxfhl.getUserid();
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

            Date jcsj = jjgFbgcJtaqssJabxfhl.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcJtaqssJabxfhl.getWzjlx())) {
                wrapper.like("wzjlx", jjgFbgcJtaqssJabxfhl.getWzjlx());
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabxfhl> pageModel = jjgFbgcJtaqssJabxfhlService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgFbgcJtaqssJabxfhl> queryWrapper = new QueryWrapper<>();
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
            boolean remove = jjgFbgcJtaqssJabxfhlService.remove(queryWrapper);
            if (remove){
                File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"58交安钢防护栏.xlsx");
                if (f.exists()){
                    f.delete();
                }

            }
        }
        return Result.ok();

    }

    @ApiOperation("批量删除交安波形防护数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabxfhlService.removeByIds(idList);
        return hd ? Result.ok() : Result.fail().message("删除失败！");


    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabxfhl/{id}")
    public Result getJabxfhl(@PathVariable String id) {
        JjgFbgcJtaqssJabxfhl user = jjgFbgcJtaqssJabxfhlService.getById(id);
        return Result.ok(user);
    }

    //@Log(title = "交安波形防护数据",businessType = BusinessType.UPDATE)
    @ApiOperation("修改交安波形防护数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabxfhl user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcJtaqssJabxfhlService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("交安波形防护数据");
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
    public Result createManyRecords(@RequestBody List<JjgFbgcJtaqssJabxfhl> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcJtaqssJabxfhlService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

