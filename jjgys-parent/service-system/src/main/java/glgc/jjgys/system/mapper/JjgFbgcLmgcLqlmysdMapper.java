package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-10
 */
@Mapper
public interface JjgFbgcLmgcLqlmysdMapper extends BaseMapper<JjgFbgcLmgcLqlmysd> {

    Map<String, String> selectsffl(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectzd(String proname, String htd, String fbgc, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectljx(String proname, String htd, String fbgc, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectljxsd(String proname, String htd, String fbgc, List<Long> idlist);

    //沥青路面压实度主线左幅上面层
    List<JjgFbgcLmgcLqlmysd> selectzxsmc(String proname, String htd, String fbgc, String sffl, String cw, List<Long> idlist);

    //沥青路面压实度主线左幅中下面层
    List<JjgFbgcLmgcLqlmysd> selectzxzxmc(String proname, String htd, String fbgc, String sffl, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectzxyfsmc(String proname, String htd, String fbgc, String sffl, String cw, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectzxyfzxmc(String proname, String htd, String fbgc, String sffl, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectsdzfsmc(String proname, String htd, String fbgc, String sffl, String cw, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectsdyfsmc(String proname, String htd, String fbgc, String sffl, String cw, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectsdzfzxmc(String proname, String htd, String fbgc, String sffl, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectsdyfzxmc(String proname, String htd, String fbgc, String sffl, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectzdsmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectzdsxmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectljxsmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectljxzxmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectljxsdsmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectljxsdzxmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    List<Map<String, Object>> selectsfl(String proname, String htd, List<Long> idlist);

    List<JjgFbgcLmgcLqlmysd> selectzdyfsmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectzdyfsxmc(String proname, String htd, String fbgc, List<Long> idlist, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectzdyf(String proname, String htd, String fbgc, List<Long> idlist);
}
