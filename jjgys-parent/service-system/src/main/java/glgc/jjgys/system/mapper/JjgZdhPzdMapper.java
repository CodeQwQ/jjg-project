package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgZdhPzd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-06-17
 */
@Mapper
public interface JjgZdhPzdMapper extends BaseMapper<JjgZdhPzd> {

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String,Object>> selectSdZfData(String proname, String htd, String zx,String zhq, String zhz,String result);

    List<Map<String,Object>> selectSdZfDatayh(String proname, String htd, String zx, String zhq, String zhz, String result, List<Long> idlist);

    List<Map<String,Object>>  selectSdyfData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String,Object>>  selectSdyfDatayh(String proname, String htd, String zx, String zhq, String zhz, String result, List<Long> idlist);

    List<Map<String,Object>> selectQlZfData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String,Object>> selectQlZfDatayh(String proname, String htd, String zx, String zhq, String zhz, String result, List<Long> idlist);

    List<Map<String,Object>> selectQlYfData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String,Object>> selectQlYfDatayh(String proname, String htd, String zx, String zhq, String zhz, String result, List<Long> idlist);

    List<Map<String, Object>> selectzfList(String proname, String htd, String zx,String result);

    List<Map<String, Object>> selectzfListyh(String proname, String htd, String zx, String result,List<Long> idlist);

    List<Map<String, Object>> selectyfList(String proname, String htd, String zx,String result);

    List<Map<String, Object>> selectyfListyh(String proname, String htd, String zx, String result, List<Long> idlist);

    List<Map<String, Object>> selectSdData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String, Object>> selectsdpzd(String proname, String bz, String lf, String zx, String zhq1, String zhz1,String sdzhz);

    List<Map<String, Object>> selectsdpzdyh(String proname, String bz, String lf, String zx, String zhq1, String zhz1, String sdzhz, List<Long> idlist);

    List<Map<String, Object>> selectqlpzd(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx,String qlzhzj);

    List<Map<String, Object>> selectqlpzdyh(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String qlzhzj, List<Long> idlist);

    List<Map<String, Object>> selectsdpzd1(String proname, String bz, String lf, String zx, String zhq1, String zhz1, String sdzhq, String sdzhz);

    List<Map<String, Object>> selectsdpzd1yh(String proname, String bz, String lf, String zx, String zhq1, String zhz1, String sdzhq, String sdzhz, List<Long> idlist);

    List<Map<String, Object>> selectqlpzd1(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selectqlpzd1yh(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1, List<Long> idlist);

    int selectnum(String proname, String htd);

    List<Map<String, Object>> selectlxid(String proname, String htd, List<Long> idlist);

    int selectcdnumid(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String, Object>> selectsdpzd1yhx(String proname, String lf, String wz, String zhq1, String zhz1, String sdname, String sdzhq, String sdzhz, List<Long> idlist);

    List<Map<String, Object>> selectqlpzd1yhx(String proname, String lf, String wz, String qlzhq, String qlzhz, String qlname, String zhq1, String zhz1, List<Long> idlist);

}
