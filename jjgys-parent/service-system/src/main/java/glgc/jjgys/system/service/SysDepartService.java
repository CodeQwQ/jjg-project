package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.SysDept;

import javax.servlet.http.HttpServletRequest;
import java.util.List;


public interface SysDepartService extends IService<SysDept> {

    String selectParentDeptName(String deptId);

    List<SysDept> selectdeptTree(HttpServletRequest request);

    boolean checkDeptNameUnique(SysDept dept);

    int insertDept(SysDept dept);

    boolean hasChildByDeptId(Long deptId);

    boolean checkDeptExistUser(Long deptId);

    int deleteDeptById(Long deptId);

    List<SysDept> selectdeptinfo();
}
