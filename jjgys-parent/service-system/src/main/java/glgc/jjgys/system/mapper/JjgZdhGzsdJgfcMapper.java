package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgZdhGzsdJgfc;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-09-23
 */
@Mapper
public interface JjgZdhGzsdJgfcMapper extends BaseMapper<JjgZdhGzsdJgfc> {

    List<String> gethtd(String proname);

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String, Object>> selecthtd(String proname);

    List<Map<String, Object>> selectzfListyh(String proname, String htd, String zx);

    List<Map<String, Object>> selectyfListyh(String proname, String htd, String zx);

    List<Map<String, Object>> selectSdZfDatayh(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectSdyfDatayh(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectQlZfDatayh(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectQlYfDatayh(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectsdgzsdyhx(String proname, String lf, String wz, String sdzhq, String sdzhz, String sdname, String zhq1, String zhz1);

    List<Map<String, Object>> selectqlgzsdyhx(String proname, String lf, String wz, String qlzhq, String qlzhz, String qlname, String zhq1, String zhz1);

    List<Map<String, Object>> selectsdgzsdyh(String proname, String bz, String lf, String sdzhq, String sdzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selectqlgzsdyh(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);



}
