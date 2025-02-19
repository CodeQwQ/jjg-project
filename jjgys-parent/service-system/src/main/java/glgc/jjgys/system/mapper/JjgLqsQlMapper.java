package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgLqsQl;
import org.apache.ibatis.annotations.Select;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2022-12-14
 */
@Repository
@Mapper
public interface JjgLqsQlMapper extends BaseMapper<JjgLqsQl> {

    //主线桥左幅
    List<JjgLqsQl> selectqlzf(String proname, String htdzhq, String htdzhz, String lf);

    //主线桥右幅
    List<JjgLqsQl> selectqlyf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsQl> selectqlList(String proname, String zhq, String zhz, String bz,String wz,String zdlf);

    List<JjgLqsQl> getqlName(String proname, String htd);

    List<JjgLqsQl> selectqlListx(String proname, String zhq, String zhz, String wz, String zdlf);

    
    @Select("select qlname from jjg_lqs_ql where proname = #{proname} and htd = #{htd}")
    List<String> getPureQlName(String proname, String htd);

}
