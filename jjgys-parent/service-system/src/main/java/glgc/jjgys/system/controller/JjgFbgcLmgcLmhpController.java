package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcLmhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFbgcLmgcLmhpService;
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
 * @since 2023-04-22
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/lmhp")
public class JjgFbgcLmgcLmhpController {

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Autowired
    private SysUserService sysUserService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd, String fbgc) throws IOException {
        List<Map<String,String>> mclist = jjgFbgcLmgcLmhpService.selectmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<mclist.size();i++){
            list.add(mclist.get(i).get("lxlx"));
        }
        String zipName = "24路面横坡";
        JjgFbgcCommonUtils.batchDownloadlmhpFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }


    @ApiOperation("生成路面横坡鉴定表")
    @PostMapping("generateJdb")
    public Result generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        boolean is_Success = jjgFbgcLmgcLmhpService.generateJdb(commonInfoVo);
        return is_Success ? Result.ok().code(200).message("成功生成鉴定表") : Result.fail().code(201).message("生成鉴定表失败");
    }
    @ApiOperation("查看路面横坡鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,String>> jdjg = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("路面横坡模板文件导出")
    @GetMapping("exportlmhp")
    public void exportlmhp(HttpServletResponse response){
        jjgFbgcLmgcLmhpService.exportlmhp(response);
    }


    @ApiOperation(value = "路面横坡数据文件导入")
    @PostMapping("importlmhp")
    public Result importlmhp(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLmhpService.importlmhp(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcLmhp jjgFbgcLmgcLmhp) throws ParseException {
        //创建page对象
        Page<JjgFbgcLmgcLmhp> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLmhp != null){
            QueryWrapper<JjgFbgcLmgcLmhp> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLmhp.getProname());
            wrapper.like("htd",jjgFbgcLmgcLmhp.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcLmhp.getFbgc());
            String userid = jjgFbgcLmgcLmhp.getUserid();
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
            Date jcsj = jjgFbgcLmgcLmhp.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                Date date = sdf.parse(jcsj.toString());
                SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
                String newFormat = outputFormat.format(date);
                wrapper.apply("DATE_FORMAT(jcsj, '%Y-%m-%d') = {0}", newFormat);
                //wrapper.like("jcsj",jcsj);
            }
            if (!StringUtils.isEmpty(jjgFbgcLmgcLmhp.getZh())) {
                wrapper.eq("zh", jjgFbgcLmgcLmhp.getZh());
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLmhp> pageModel = jjgFbgcLmgcLmhpService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删路面横坡强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLmhpService.removeByIds(idList);
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
        QueryWrapper<JjgFbgcLmgcLmhp> queryWrapper = new QueryWrapper<>();
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
            List<Map<String,String>> mclist = jjgFbgcLmgcLmhpService.selectmc(proname,htd,"路面工程");
            List<String> list = new ArrayList<>();
            for (int i=0;i<mclist.size();i++){
                String lxlx = mclist.get(i).get("lxlx");
                if (lxlx.equals("主线")){
                    list.add("24路面横坡.xlsx");
                }else {
                    list.add("24路面横坡-"+mclist.get(i).get("lxlx")+".xlsx");
                }

            }
            boolean remove = jjgFbgcLmgcLmhpService.remove(queryWrapper);
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
        JjgFbgcLmgcLmhp user = jjgFbgcLmgcLmhpService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路面横坡数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcLmhp user) {
        boolean is_Success = jjgFbgcLmgcLmhpService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @PostMapping("createManyRecords")
    public Result createManyRecords(@RequestBody List<JjgFbgcLmgcLmhp> rawData, HttpServletRequest request) {
        String userID = ReceiveUtils.getUserIDFromRequest(request);
        if (userID == null) {
            return Result.fail("用户未登录");
        }
        int res = jjgFbgcLmgcLmhpService.createMoreRecords(rawData, userID);
        return Result.ok(res);
    }

}

