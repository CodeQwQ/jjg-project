package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcQlgcQmpzd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Mapper
public interface JjgFbgcQlgcQmpzdMapper extends BaseMapper<JjgFbgcQlgcQmpzd> {

    List<Map<String, Object>> selectqlmc(String proname, String htd, String fbgc, List<Long> idlist);

    List<Map<String, Object>> selectqlmcno(String proname, String htd);

    List<Map<String, Object>> selectqlmc2(String proname, String htd, List<Long> idlist);

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    List<Map<String, Object>> getqlName(String proname, String htd);

}
