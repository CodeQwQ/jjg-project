package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcZdhldhd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-10-23
 */
@Mapper
public interface JjgFbgcSdgcZdhldhdMapper extends BaseMapper<JjgFbgcSdgcZdhldhd> {

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String, Object>> selectzfList(String proname, String htd, String result, String s);

    List<Map<String, Object>> selectzfListyh(String proname, String htd, String zx, String result, List<Long> idlist);

    List<Map<String, Object>> selectyfList(String proname, String htd, String result, String s);

    List<Map<String, Object>> selectyfListyh(String proname, String htd, String zx, String result, List<Long> idlist);

    List<Map<String, Object>> seletcfhlmzfData(String proname, String htd, String zhq, String zhz);

    List<Map<String, Object>>seletcfhlmzfDatayh(String proname, String htd, String zhq, String zhz, List<Long> idlist);

    List<Map<String, Object>>seletcfhlmyfData(String proname, String htd, String zhq, String zhz);

    List<Map<String, Object>>seletcfhlmyfDatayh(String proname, String htd, String zhq, String zhz, List<Long> idlist);

    List<Map<String, Object>> selectlxid(String proname, String htd, List<Long> idlist);

    int selectcdnumid(String proname, String htd, String zx, List<Long> idlist);

}
