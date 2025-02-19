package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxJgfc;
import org.mapstruct.Mapper;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2024-05-20
 */
@Mapper
public interface JjgFbgcJtaqssJabxJgfcMapper extends BaseMapper<JjgFbgcJtaqssJabxJgfc> {

    List<String> gethtd(String proname);

    List<JjgFbgcJtaqssJabxJgfc> selectbxnfsxs(String proname, String htd);

    List<JjgFbgcJtaqssJabxJgfc> selecthxnfsxs(String proname, String htd);
}
