package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-22
 */
@Mapper
public interface JjgFbgcQlgcXbTqdMapper extends BaseMapper<JjgFbgcQlgcXbTqd> {

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    List<Map<String, Object>> getqlName(String proname, String htd);
}
