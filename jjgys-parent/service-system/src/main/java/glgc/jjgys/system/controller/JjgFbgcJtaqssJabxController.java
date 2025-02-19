package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.service.SysUserService;
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
import java.io.*;
import java.net.URLEncoder;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jabx")
@CrossOrigin
public class JjgFbgcJtaqssJabxController {

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd) throws IOException {
        String fpath= filespath+File.separator+proname+File.separator+htd;
        String fileName1 = "57交安标线厚度.xlsx";
        String fileName2 = "57交安标线逆反射系数.xlsx";
        //String fileName3 = "57交安标线黄线逆反射系数.xlsx";
        List list = new ArrayList();
        list.add(fileName1);
        list.add(fileName2);
        //list.add(fileName3);
        //设置响应头信息csc
        /*response.reset();
        response.setCharacterEncoding("utf-8");
        response.setContentType("multipart/form-data");*/
        //设置压缩包的名字
        String downloadName = URLEncoder.encode("57交安标线.zip", "UTF-8");
        //String userAgent = request.getHeader("User-Agent");
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + downloadName);
        response.setContentType("application/zip;charset=utf-8");
        //response.setCharacterEncoding("utf-8");
        //返回客户端浏览器的版本号、类型
        /*String agent = request.getHeader("USER-AGENT");
        try {
            //针对IE或者以IE为内核的浏览器：
            if (agent.contains("MSIE") || agent.contains("Trident")) {
                downloadName = java.net.URLEncoder.encode(downloadName, "UTF-8");
            } else {
                //非IE浏览器的处理：
                downloadName = new String(downloadName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
            }
            System.out.println(downloadName);
            System.out.println("111");
        } catch (Exception e) {
            e.printStackTrace();
        }*/
        //response.setHeader("Content-Disposition", "attachment;fileName=\"" + downloadName + "\"");
        //设置压缩流：直接写入response，实现边压缩边下载
        ZipOutputStream zipOs = null;
        //循环将文件写入压缩流
        DataOutputStream os = null;
        //文件
        File file;
        try {
            zipOs = new ZipOutputStream(new BufferedOutputStream(response.getOutputStream()));
            //设置压缩方法
            zipOs.setMethod(ZipOutputStream.DEFLATED);
            //遍历文件信息（主要获取文件名/文件路径等）
            for (int i=0;i<list.size();i++) {
                String name = list.get(i).toString();
                String path = fpath+File.separator+name;
                file = new File(path);
                if (!file.exists()) {
                    break;
                }
                //添加ZipEntry，并将ZipEntry写入文件流
                zipOs.putNextEntry(new ZipEntry(name));
                os = new DataOutputStream(zipOs);
                FileInputStream fs = new FileInputStream(file);
                byte[] b = new byte[100];
                int length;
                //读入需要下载的文件的内容，打包到zip文件
                while ((length = fs.read(b)) != -1) {
                    os.write(b, 0, length);
                }
                //关闭流
                fs.close();
                zipOs.closeEntry();

            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //关闭流
            try {
                if (os != null) {
                    os.flush();
                    os.close();
                }
                if (zipOs != null) {
                    zipOs.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    @ApiOperation("生成交安标线鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcJtaqssJabxService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");

    }

    @ApiOperation("查看交安标线鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安标线模板文件导出")
    @GetMapping("exportjabx")
    public void exportjabx(HttpServletResponse response){
        jjgFbgcJtaqssJabxService.exportjabx(response);
    }

    @ApiOperation(value = "交安标线数据文件导入")
    @PostMapping("importjabx")
    public Result importjabx(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJabxService.importjabx(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabx jjgFbgcJtaqssJabx) throws ParseException {
        //创建page对象
        Page<JjgFbgcJtaqssJabx> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabx != null) {
            QueryWrapper<JjgFbgcJtaqssJabx> wrapper = new QueryWrapper<>();
            wrapper.eq("proname", jjgFbgcJtaqssJabx.getProname());
            wrapper.eq("htd", jjgFbgcJtaqssJabx.getHtd());
            wrapper.eq("fbgc", jjgFbgcJtaqssJabx.getFbgc());

            String userid = jjgFbgcJtaqssJabx.getUserid();
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
            Date jcsj = jjgFbgcJtaqssJabx.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcJtaqssJabx.getWz())) {
                wrapper.like("wz", jjgFbgcJtaqssJabx.getWz());
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabx> pageModel = jjgFbgcJtaqssJabxService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }


    @ApiOperation("批量删除交安标线数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabxService.removeByIds(idList);
        return hd ? Result.ok():Result.fail().message("删除失败！");
    }

    @ApiOperation("全部删除")
    @DeleteMapping("removeAll")
    public Result removeAll(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String username = commonInfoVo.getUsername();
        QueryWrapper<JjgFbgcJtaqssJabx> queryWrapper = new QueryWrapper<>();
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
            boolean remove = jjgFbgcJtaqssJabxService.remove(queryWrapper);
            if (remove){
                File f = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"57交安标线厚度.xlsx");
                File f1 = new File(filespath+File.separator+proname+File.separator+htd+File.separator+"57交安标线逆反射系数.xlsx");
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
    @GetMapping("getJabx/{id}")
    public Result getJabx(@PathVariable String id) {
        JjgFbgcJtaqssJabx user = jjgFbgcJtaqssJabxService.getById(id);
        return Result.ok(user);
    }

    //@Log(title = "交安标线数据",businessType = BusinessType.UPDATE)
    @ApiOperation("修改交安标线数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabx user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcJtaqssJabxService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("交安标线数据");
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
/*
    * @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcLmgcLqlmysd> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcLmgcLqlmysdService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }
    * */

    @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcJtaqssJabx> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcJtaqssJabxService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

