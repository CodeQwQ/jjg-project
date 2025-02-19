package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.LambdaQueryWrapper;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.system.SysDept;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.mapper.SysDepartMapper;
import glgc.jjgys.system.service.SysDepartService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import java.util.*;

@Service
public class SysDepartServiceImpl extends ServiceImpl<SysDepartMapper, SysDept> implements SysDepartService {

    @Autowired
    private SysDepartMapper sysDepartMapper;

    @Override
    public String selectParentDeptName(String deptId) {
        // 根据用户表中的部门id，查询在部门表中的信息
        QueryWrapper<SysDept> wrapper = new QueryWrapper<>();
        if(!StringUtils.isEmpty(deptId)) {
            wrapper.like("id",deptId);
        }
        SysDept sysDept = baseMapper.selectOne(wrapper);
        //获取当前部门id下的treePath
        String treePath = sysDept.getTreePath();
        List<String> list = Arrays.asList(treePath.split(","));
        List<SysDept> sysDeptParent = baseMapper.selectBatchIds(list);
        StringBuffer sb = new StringBuffer();
        for (SysDept dept : sysDeptParent) {
            String name = dept.getName();
            sb.append(name).append(",");
        }
        String str = sb.deleteCharAt(sb.length() - 1).toString();
        return str;
    }

    @Override
    public List<SysDept> selectdeptTree(HttpServletRequest request) {
        SysUser user = getUserInfoByRequest(request);
        if ("2".equals(user.getType()) || "4".equals(user.getType())){
            // 公司管理员
            LambdaQueryWrapper<SysDept> eq = new LambdaQueryWrapper<SysDept>().eq(SysDept::getCompanyId, user.getCompanyId());
            List<SysDept> list = list(eq);
            return getTreeList(list);
        }else if ("1".equals(user.getType())){
            // 系统管理员
            List<SysDept> list = list();
            return getTreeList(list);
        }
        return new ArrayList<>();
    }

    @Override
    public boolean checkDeptNameUnique(SysDept dept) {

        SysDept info = sysDepartMapper.checkDeptNameUnique(dept.getName());
        if (StringUtils.isNotNull(info)){
            return false;
        }
        return true;
       /* Long deptId = StringUtils.isNull(dept.getId()) ? -1L : dept.getId();

        SysDept info = sysDepartMapper.checkDeptNameUnique(dept.getName(), dept.getParentId());
        if (StringUtils.isNotNull(info) && info.getId().longValue() != deptId.longValue())
        {
            return false;
        }
        return true;*/
    }

    /**
     * 新增
     * @param dept
     * @return
     */
    @Override
    public int insertDept(SysDept dept) {
        return sysDepartMapper.insertDept(dept);
    }


    @Override
    public boolean hasChildByDeptId(Long deptId) {
        int result = sysDepartMapper.hasChildByDeptId(deptId);
        return result > 0;
    }

    @Override
    public boolean checkDeptExistUser(Long deptId) {
        int result = sysDepartMapper.checkDeptExistUser(deptId);
        return result > 0;
    }

    @Override
    public int deleteDeptById(Long deptId) {
        return sysDepartMapper.deleteById(deptId);
    }

    @Override
    public List<SysDept> selectdeptinfo() {
        List<SysDept> list = list();
        return getTreeList(list);
    }

    private List<SysDept> getTreeList(List<SysDept> sysDepts) {
        Map<Long, SysDept> map = new HashMap<>();
        List<SysDept> tree = new ArrayList<>();

        for (SysDept node : sysDepts) {
            map.put(node.getId(), node);
        }

        for (SysDept node : sysDepts) {
            SysDept parent = map.get(node.getParentId());
            if (parent != null) {
                List<SysDept> children = parent.getChildren();
                if (children == null){
                    children = new ArrayList<SysDept>();
                }
                children.add(node);
                parent.setChildren(children);
            } else {
                tree.add(node);
            }
        }

        return tree;
    }

    @Resource
    SysUserService sysUserService;
    private SysUser getUserInfoByRequest(HttpServletRequest request){
        //获取请求头token字符串,因为把token放到请求头不存在跨域
        String token = request.getHeader("token");
        //从token字符串获取用户名称（id）
        String userId = JwtHelper.getUserId(token);
        //根据用户名称获取用户信息
        return sysUserService.getById(userId);
    }
}
