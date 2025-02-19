package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.SysMenu;

import java.util.List;


public interface ProjectService extends IService<Project> {

    void addOtherInfo(Project project, String proName, String userid);

    Integer getlevel(String proname);

    void deleteOtherInfo(List<String> idList);

    String getGrade(String proname);

    List<SysMenu> projectsTree(List<SysMenu> nodes);

    void addptyhqx(String userid, String participant, String proName);
}
