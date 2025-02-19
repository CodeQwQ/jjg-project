package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.system.SysDept;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface SysDepartMapper extends BaseMapper<SysDept> {

    SysDept checkDeptNameUnique(String name);

    SysDept selectDeptById(Long parentId);

    int insertDept(SysDept dept);


    int hasChildByDeptId(Long deptId);

    int checkDeptExistUser(Long deptId);

}
