package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcLcfJgfc;
import org.mapstruct.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2024-04-30
 */
@Mapper
public interface JjgFbgcLmgcLmwcLcfJgfcMapper extends BaseMapper<JjgFbgcLmgcLmwcLcfJgfc> {

    List<Map<String, Object>> selecthtd(String proname);

    List<String> gethtd(String proname);

}
