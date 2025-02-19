package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcQlgcSbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcQlgcSbTqdService;
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
 *  桥梁上部砼强度
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/sb/tqd")
@CrossOrigin
public class JjgFbgcQlgcSbTqdController {

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Autowired
    private SysUserService sysUserService;



    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "29桥梁上部砼强度.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成桥梁上部砼强度鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcQlgcSbTqdService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看桥梁上部砼强度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcSbTqdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥梁上部砼强度模板文件导出")
    @GetMapping("exportqlsbtqd")
    public void exportqlsbtqd(HttpServletResponse response){
        jjgFbgcQlgcSbTqdService.exportqlsbtqd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥梁上部砼强度数据文件导入")
    @PostMapping("importqlsbtqd")
    public Result importqlsbtqd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcSbTqdService.importqlsbtqd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcSbTqd jjgFbgcQlgcSbTqd) throws ParseException {
        //创建page对象
        Page<JjgFbgcQlgcSbTqd> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcSbTqd != null){
            QueryWrapper<JjgFbgcQlgcSbTqd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcSbTqd.getProname());
            wrapper.like("htd",jjgFbgcQlgcSbTqd.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcSbTqd.getFbgc());
            String userid = jjgFbgcQlgcSbTqd.getUserid();
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
            Date jcsj = jjgFbgcQlgcSbTqd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcQlgcSbTqd.getQlmc())) {
                wrapper.like("qlmc", jjgFbgcQlgcSbTqd.getQlmc());
            }
            if (!StringUtils.isEmpty(jjgFbgcQlgcSbTqd.getBw1())) {
                wrapper.like("bw1", jjgFbgcQlgcSbTqd.getBw1());
            }
            if (!StringUtils.isEmpty(jjgFbgcQlgcSbTqd.getBw2())) {
                wrapper.like("bw2", jjgFbgcQlgcSbTqd.getBw2());
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcSbTqd> pageModel = jjgFbgcQlgcSbTqdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除桥梁上部砼强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcQlgcSbTqdService.removeByIds(idList);
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
        QueryWrapper<JjgFbgcQlgcSbTqd> queryWrapper = new QueryWrapper<>();
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
            boolean remove = jjgFbgcQlgcSbTqdService.remove(queryWrapper);
            if (remove){
                File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"29桥梁上部砼强度.xlsx");
                if (f.exists()){
                    f.delete();
                }
            }
        }
        return Result.ok();

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSbTqd{id}")
    public Result getSbTqd(@PathVariable String id) {
        JjgFbgcQlgcSbTqd user = jjgFbgcQlgcSbTqdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥梁上部砼强度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcSbTqd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcQlgcSbTqdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("桥梁上部砼强度数据");
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
    public Result createManyRecords(@RequestBody List<JjgFbgcQlgcSbTqd> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcQlgcSbTqdService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }


}

