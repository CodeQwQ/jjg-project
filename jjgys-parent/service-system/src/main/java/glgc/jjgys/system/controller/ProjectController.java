package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.*;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletRequest;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Api(tags = "项目管理接口")
@RestController
@Transactional
@RequestMapping(value = "/project")
@CrossOrigin//解决跨域
public class ProjectController {

    @Autowired
    private ProjectService projectService;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private OperLogService operLogService;
     
    @Autowired
    private SysMenuService sysMenuService;

    @Autowired
    private JjgLqsQlService jjgLqsQlService;


    @GetMapping("selectuser")
    public Result selectuser(String companyId){
        QueryWrapper<SysUser> sysUserQueryWrapper = new QueryWrapper<>();
        sysUserQueryWrapper.eq("company_id",companyId);
        List<SysUser> list = sysUserService.list(sysUserQueryWrapper);
        return Result.ok(list);
    }

    @ApiOperation("根据id查询")
    @GetMapping("getproject/{id}")
    public Result getJabx(@PathVariable String id) {
        Project user = projectService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改项目数据")
    @PostMapping("update")
    public Result update(@RequestBody Project user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = projectService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProName());
            sysOperLog.setHtd("-");
            sysOperLog.setFbgc("-");
            sysOperLog.setTitle("交工项目数据");
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



    /**
     * 新增项目 包括新增合同段信息 路桥隧信息
     */
    @PostMapping("/add")
    public Result add(@RequestBody Project project){
        String proName = project.getProName();
        String userid = project.getUserid();
        String participant = project.getParticipant();
        projectService.addOtherInfo(project,proName,userid);
        projectService.addptyhqx(userid,participant,proName);
        boolean res = projectService.save(project);
        return res ? Result.ok().message("增加成功") : Result.fail().message("增加失败！");

    }

    @ApiOperation("校验项目")
    @GetMapping("checkProname/{proname}")
    public Result checkProname(@PathVariable String proname) {
        QueryWrapper<Project> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<Project> list = projectService.list(wrapper);
        return list == null || list.isEmpty() ? Result.ok().message("校验成功"):Result.ok().message("校验失败");
    }

    /**
     * 删除项目
     */
    @DeleteMapping("/remove/{id}")
    public Result remove(@PathVariable String id){
        projectService.removeById(id);
        return Result.ok(null);
    }

    @ApiOperation("批量删除项目信息")
    @Transactional
    @DeleteMapping("removeBatch")
    //传json数组[1,2,3]，用List接收
    public Result removeBeatch(@RequestBody List<String> idList){
        projectService.deleteOtherInfo(idList);
        boolean isSuccess = projectService.removeByIds(idList);
        return Result.ok();
    }

    /**
     * 查询所有的项目
     */
    @GetMapping("findAll")
    public Result findAllProject(){
        List<Project> list = projectService.list();
        return Result.ok(list);
    }

    /**
     * 分页查询
     */
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody Project project){
        //创建page对象
        Page<Project> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(project == null){
            IPage<Project> pageModel = projectService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            /**
             *要是admin，查询全部的项目
             * 要是公司管理员，查询本公司的项目
             * 要是公司的普通用户，查询当前用户的
             */
            String proName = project.getProName();
            QueryWrapper<Project> wrapper=new QueryWrapper<>();

            String userid = project.getUserid();
            if ("1".equals(userid)){


            }else {
                //查询用户类型
                QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
                wrapperuser.eq("id",userid);
                SysUser one = sysUserService.getOne(wrapperuser);
                String type = one.getType();
                if ("4".equals(type)){
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

                }else if ("3".equals(type) || "2".equals(type) || "4".equals(type)){
                    //普通用户
                    String username = project.getUsername();
                    wrapper.like("participant",username);

                }
            }


            if (!StringUtils.isEmpty(proName)){
                wrapper.like("proname",proName);
            }
            if (!StringUtils.isEmpty(project.getParticipant())){
                wrapper.like("participant",project.getParticipant());
            }
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<Project> pageModel = projectService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }
    @GetMapping("getAllProject4App")
    public Result getAllProject4App() {
        List<SysMenu> nodes = sysMenuService.findNodes();
        List<SysMenu> sysMenus = projectService.projectsTree(nodes);
        return Result.ok(sysMenus);
    }

    @GetMapping("ql")
    public Result getQlInfo(@RequestParam("proname") String proname, @RequestParam("htd") String htd) {
        List<String> jjgLqsQls = jjgLqsQlService.getPureQlName(proname, htd);
        return Result.ok(jjgLqsQls);
    }
}
