package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.JjgFileInfo;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.mapper.JjgFileInfoMapper;
import glgc.jjgys.system.service.JjgFileInfoService;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * <p>
 * 文件资源表 服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-10-15
 */
@Service
public class JjgFileInfoServiceImpl extends ServiceImpl<JjgFileInfoMapper, JjgFileInfo> implements JjgFileInfoService {

    @Value(value = "${jjgys.path.file}")
    private String filepath;

    @Autowired
    private SysUserService sysUserService;

    @Autowired
    private ProjectService projectService;

    @Override
    public List<JjgFileInfo> getfilelist(String userid) {
        List<JjgFileInfo> fileList = new ArrayList<>();
        File directory = new File(filepath);
        if (directory.exists() && directory.isDirectory()) {
            File[] files = directory.listFiles();
            for (File file : files) {
                JjgFileInfo jjgFileInfo = new JjgFileInfo();
                jjgFileInfo.setPath(file.getAbsolutePath());
                jjgFileInfo.setName(file.getName());
                if (file.getName().equals("files")){
                    jjgFileInfo.setFileName("交工文件");
                }else if (file.getName().equals("jgfiles")){
                    jjgFileInfo.setFileName("竣工文件");
                }else {
                    jjgFileInfo.setFileName(file.getName());
                }
                jjgFileInfo.setIsFile(file.isFile());
                jjgFileInfo.setIsDir(file.isDirectory());
                if (file.isDirectory()) {
                    jjgFileInfo.setChildren(getChildFiles(file.getAbsolutePath()));
                }
                fileList.add(jjgFileInfo);
            }
        }
        /**
         * 需要过滤一下fileList中的数据
         * 根据当前用户，判断是哪个公司的用户，然后存储该公司项目的数据。
         */
        if (!"1".equals(userid)){
            //根据userid查询这个用户所属的公司id
            QueryWrapper<SysUser> wrapper=new QueryWrapper<>();
            wrapper.eq("id",userid);
            SysUser userData = sysUserService.getOne(wrapper);
            Long companyId = userData.getCompanyId();
            //获取所有项目数据，项目数据中只有创建人的userid，需要根据userid再去查询公司id
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
            /**
             * 此时的list中，username属性存储的是项目所属的公司id
             *
             * 现在就是要根据companyId，获取到list中的项目名称，然后过滤fileList。
             */
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

            //过滤fileList中chilren中的项目数据，只保留pronameList中的项目
            if (fileList != null){
                for (JjgFileInfo jjgFileInfo : fileList) {
                    List<JjgFileInfo> children = jjgFileInfo.getChildren();
                    if (children != null){
                        Iterator<JjgFileInfo> iterator = children.iterator();
                        while (iterator.hasNext()) {
                            JjgFileInfo child = iterator.next();
                            String name = child.getName();
                            if (!pronameList.contains(name)) {
                                iterator.remove(); // 使用迭代器的 remove 方法安全删除元素
                            }
                        }
                    }
                }
            }
        }



        return fileList;

    }

    @Override
    public void download(HttpServletResponse response, List<JjgFileInfo> downloadPath) throws IOException {
        System.out.println(downloadPath);
        if (downloadPath.size() == 1 && downloadPath.get(0).getIsDir()){
            JjgFbgcCommonUtils.Downloadfile(response,downloadPath.get(0).getName(),downloadPath.get(0).getPath());
        } else {
            // Download file
            if (downloadPath.size() == 1){
                // Download single file
                JjgFbgcCommonUtils.download(response,downloadPath.get(0).getPath(),downloadPath.get(0).getName());
            }else {
                // Download batch file
                String zipname = "文件";
                //String path = downloadPath.get(0).getPath();
                List<String> list = new ArrayList<>();
                for (JjgFileInfo jjgFileInfo : downloadPath) {
                    list.add(jjgFileInfo.getPath());
                }
                System.out.println(list+"wq");
                JjgFbgcCommonUtils.DownloadBatch(response,zipname,list);
                //downloadBatch(downloadPath);
            }
        }

    }

    /**
     *
     * @param files
     * @throws IOException
     */
    public void downloadBatch(List<JjgFileInfo> files) throws IOException {
        String zipFileName = "文件.zip";
        ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(zipFileName));

        byte[] buffer = new byte[4096];
        for (JjgFileInfo fileInfo : files) {
            if (fileInfo.getIsDir()) {
                zipDirectory(new File(fileInfo.getPath()), fileInfo.getFileName(), zipOut);
            } else {
                File file = new File(fileInfo.getPath());
                FileInputStream fileIn = new FileInputStream(file);
                ZipEntry zipEntry = new ZipEntry(fileInfo.getFileName());
                zipOut.putNextEntry(zipEntry);

                int length;
                while ((length = fileIn.read(buffer)) > 0) {
                    zipOut.write(buffer, 0, length);
                }

                fileIn.close();
            }
        }
        zipOut.close();

    }

    /**
     *
     * @param folder
     * @param parentFolder
     * @param zipOut
     * @throws IOException
     */
    private void zipDirectory(File folder, String parentFolder, ZipOutputStream zipOut) throws IOException {
        byte[] buffer = new byte[4096];
        File[] files = folder.listFiles();

        for (File file : files) {
            if (file.isDirectory()) {
                zipDirectory(file, parentFolder + "/" + file.getName(), zipOut);
                continue;
            }

            FileInputStream fileIn = new FileInputStream(file);
            ZipEntry zipEntry = new ZipEntry(parentFolder + "/" + file.getName());
            zipOut.putNextEntry(zipEntry);

            int length;
            while ((length = fileIn.read(buffer)) > 0) {
                zipOut.write(buffer, 0, length);
            }

            fileIn.close();
        }
    }

    /**
     *
     * @param directoryPath
     * @return
     */
    private List<JjgFileInfo> getChildFiles(String directoryPath) {
        List<JjgFileInfo> fileList = new ArrayList<>();
        File directory = new File(directoryPath);
        if (directory.exists() && directory.isDirectory()) {
            File[] files = directory.listFiles();
            for (File file : files) {
                JjgFileInfo jjgFileInfo = new JjgFileInfo();
                jjgFileInfo.setPath(file.getAbsolutePath());
                jjgFileInfo.setName(file.getName());
                jjgFileInfo.setFileName(file.getName());
                jjgFileInfo.setIsFile(file.isFile());
                jjgFileInfo.setIsDir(file.isDirectory());
                if (file.isDirectory()) {
                    jjgFileInfo.setChildren(getChildFiles(file.getAbsolutePath()));
                }
                fileList.add(jjgFileInfo);
            }
        }
        return fileList;
    }
}
