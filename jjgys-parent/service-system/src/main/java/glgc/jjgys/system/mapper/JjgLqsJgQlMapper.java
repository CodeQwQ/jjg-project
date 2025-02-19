package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgLqsJgQl;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-09-22
 */
@Mapper
public interface JjgLqsJgQlMapper extends BaseMapper<JjgLqsJgQl> {

    //主线桥左幅
    List<JjgLqsJgQl> selectqlzf(String proname, String htdzhq, String htdzhz, String lf);

    //主线桥右幅
    List<JjgLqsJgQl> selectqlyf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsJgQl> selectqlList(String proname, String zhq, String zhz, String bz,String wz,String zdlf);

    List<JjgLqsJgQl> selectqlListz(String proname, String zhq, String zhz, String bz, String wz, String zdlf);

    List<JjgLqsJgQl> selectqlListx(String proname, String zhq, String zhz, String wz, String zdlf);

}
