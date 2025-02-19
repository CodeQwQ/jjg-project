package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFileInfo;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.service.JjgFileInfoService;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.service.SysUserService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * <p>
 * 文件资源表 前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-10-15
 */
@RestController
@RequestMapping("/jjg/file/info")
public class JjgFileInfoController {

    @Autowired
    private JjgFileInfoService jjgFileInfoService;

    @Autowired
    private ProjectService projectService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @ApiOperation("查看文件列表")
    @GetMapping("filelist/{userid}")
    public Result getfilelist(@PathVariable String userid) {
        List<JjgFileInfo> list = jjgFileInfoService.getfilelist(userid);
        return Result.ok(list);

    }

    //下载
    @ApiOperation("下载")
    @PostMapping("download")
    public void download(HttpServletResponse response, @RequestBody List<JjgFileInfo> list) throws IOException {
        jjgFileInfoService.download(response,list);

    }

    //删除
    @ApiOperation("删除")
    @PostMapping("delete")
    public void delete(@RequestBody List<JjgFileInfo> list) throws IOException {
        if (list != null){
            for (JjgFileInfo jjgFileInfo : list) {
                Path path = Paths.get(jjgFileInfo.getPath());
                Files.walkFileTree(path,
                        new SimpleFileVisitor<Path>() {
                            // 先去遍历删除文件
                            @Override
                            public FileVisitResult visitFile(Path file,
                                                             BasicFileAttributes attrs) throws IOException {
                                Files.delete(file);
                                System.out.printf("文件被删除 : %s%n", file);
                                return FileVisitResult.CONTINUE;
                            }

                            // 再去遍历删除目录
                            @Override
                            public FileVisitResult postVisitDirectory(Path dir,
                                                                      IOException exc) throws IOException {
                                Files.delete(dir);
                                System.out.printf("文件夹被删除: %s%n", dir);
                                return FileVisitResult.CONTINUE;
                            }

                        }
                );
            }
        }

    }

    @Autowired
    private SysUserService sysUserService;

    @ApiOperation("查询项目")
    @PostMapping("getxm/{userid}")
    public Result getxm(@PathVariable String userid){
        if (!"1".equals(userid)){
            //根据userid查询这个用户所属的公司id
            QueryWrapper<SysUser> wrapper1=new QueryWrapper<>();
            wrapper1.eq("id",userid);
            SysUser userData = sysUserService.getOne(wrapper1);
            Long companyId = userData.getCompanyId();
            QueryWrapper<Project> prowrapper=new QueryWrapper<>();
            List<Project> list = projectService.list(prowrapper);
            if (list != null){
                for (Project project : list) {
                    String xmuserid = project.getUserid();
                    QueryWrapper<SysUser> xmwrapper=new QueryWrapper<>();
                    xmwrapper.eq("id",xmuserid);
                    SysUser xmData = sysUserService.getOne(xmwrapper);
                    if (xmData != null){
                        Long companyId1 = xmData.getCompanyId();
                        project.setUsername(String.valueOf(companyId1));
                    }
                }
            }
            List<String> pronameList = new ArrayList<>();
            if (list != null){
                for (Project project : list) {
                    String username = project.getUsername();
                    String s = String.valueOf(companyId);
                    if (username != null && username.equals(s)){
                        pronameList.add(project.getProName());
                    }
                }
            }

            QueryWrapper<Project> wrapper = new QueryWrapper<>();
            List<Project> list1 = projectService.list(wrapper);
            if (list1 != null){
                Iterator<Project> iterator = list1.iterator();
                while (iterator.hasNext()) {
                    Project project = iterator.next();
                    String proName = project.getProName();
                    if (!pronameList.contains(proName)) {
                        iterator.remove(); // 使用迭代器的 remove 方法安全删除元素
                    }
                }
            }
            return Result.ok(list1);
        }else {
            QueryWrapper<Project> wrapper = new QueryWrapper<>();
            List<Project> list1 = projectService.list(wrapper);
            return Result.ok(list1);
        }
    }

    //上传
    @ApiOperation("上传")
    @PostMapping("upload")
    public Result upload(@RequestParam("file") MultipartFile file, Project project) {
        // 检查文件是否为空
        if (file.isEmpty()) {
            return Result.fail();
        }try {
            System.out.println(project.getProName());
            // 获取文件名
            String fileName = StringUtils.cleanPath(file.getOriginalFilename());
            String filepath = filespath+ File.separator+project.getProName() + File.separator + "设计图";
            // 设置文件存储路径
            Path uploadPath = Paths.get(filepath);
            if (!Files.exists(uploadPath)) {
                Files.createDirectories(uploadPath);
            }
            // 处理文件上传
            Path filePath = uploadPath.resolve(fileName);
            Files.copy(file.getInputStream(), filePath);
            return Result.ok("文件上传成功!");
        } catch (IOException e) {
            e.printStackTrace();
            return Result.fail();
        }
    }

}

