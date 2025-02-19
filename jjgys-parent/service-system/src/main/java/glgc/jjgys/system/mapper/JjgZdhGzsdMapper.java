package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgZdhGzsd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-06-15
 */
@Mapper
public interface JjgZdhGzsdMapper extends BaseMapper<JjgZdhGzsd> {

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String, Object>> selectzfList(String proname, String htd, String zx);

    List<Map<String, Object>> selectzfListyh(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String, Object>> selectyfList(String proname, String htd, String zx);

    List<Map<String, Object>> selectyfListyh(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String, Object>> selectSdZfData(String proname, String htd, String zx,String zhq, String zhz);

    List<Map<String, Object>> selectSdZfDatayh(String proname, String htd, String zx, String zhq, String zhz, List<Long> idlist);

    List<Map<String, Object>> selectSdyfData(String proname, String htd, String zx,String zhq, String zhz);

    List<Map<String, Object>> selectSdyfDatayh(String proname, String htd, String zx, String zhq, String zhz, List<Long> idlist);

    List<Map<String, Object>> selectQlZfData(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectQlZfDatayh(String proname, String htd, String zx, String zhq, String zhz, List<Long> idlist);

    List<Map<String, Object>> selectQlYfData(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectQlYfDatayh(String proname, String htd, String zx, String zhq, String zhz, List<Long> idlist);

    List<Map<String, Object>> selectsdgzsd(String proname,String bz,String lf, String sdzhq, String sdzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selectsdgzsdyh(String proname, String bz, String lf, String sdzhq, String sdzhz, String zx, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String, Object>> selectqlgzsd(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selectqlgzsdyh(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1, List<Long> idlist);

    int selectnum(String proname, String htd);

    List<Map<String, Object>> selectlxid(String proname, String htd, List<Long> idlist);

    int selectcdnumid(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String, Object>> selectsdgzsdyhx(String proname, String lf, String wz, String sdzhq, String sdzhz, String sdname, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String, Object>> selectqlgzsdyhx(String proname, String lf, String wz, String qlzhq, String qlzhz, String qlname, String zhq1, String zhz1, List<Long> idlist);

}
