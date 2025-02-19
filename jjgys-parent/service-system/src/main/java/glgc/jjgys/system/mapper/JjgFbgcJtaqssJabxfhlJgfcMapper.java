package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhlJgfc;
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
public interface JjgFbgcJtaqssJabxfhlJgfcMapper extends BaseMapper<JjgFbgcJtaqssJabxfhlJgfc> {

    List<String> gethtd(String proname);
}
