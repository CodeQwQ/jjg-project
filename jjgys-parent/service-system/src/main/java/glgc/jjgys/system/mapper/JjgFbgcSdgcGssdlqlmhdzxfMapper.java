package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcGssdlqlmhdzxf;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-05-20
 */
@Mapper
public interface JjgFbgcSdgcGssdlqlmhdzxfMapper extends BaseMapper<JjgFbgcSdgcGssdlqlmhdzxf> {

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);

    List<Map<String, Object>> selectsdmcid(String proname, String htd, String fbgc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectzxzf(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectzxyf(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectsdzf(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectsdyf(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectqlzf(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectqlyf(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectzd(String proname, String htd, String fbgc, String sdmc, List<Long> idlist);

    List<Map<String, Object>> selectsdmc1(String proname, String htd);

    List<Map<String, Object>> selectsdmc1id(String proname, String htd,List<Long> idlist);

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    List<Map<String, Object>> getsdnum(String proname, String htd);
}
