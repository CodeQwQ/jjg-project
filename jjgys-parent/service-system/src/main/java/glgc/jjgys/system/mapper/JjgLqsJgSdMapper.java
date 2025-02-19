package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgLqsJgSd;
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
public interface JjgLqsJgSdMapper extends BaseMapper<JjgLqsJgSd> {

    //主线隧道左幅
    List<JjgLqsJgSd> selectsdzf(String proname, String htdzhq, String htdzhz, String lf);

    //主线隧道右幅
    List<JjgLqsJgSd> selectsdyf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsJgSd> selectsdList(String proname, String zhq, String zhz,String bz, String wz,String zdlf);

    List<JjgLqsJgSd> selectsdListz(String proname, String zhq, String zhz, String bz, String wz, String zdlf);

    List<JjgLqsJgSd> selectsdListx(String proname, String zhq, String zhz, String wz, String zdlf);


}
