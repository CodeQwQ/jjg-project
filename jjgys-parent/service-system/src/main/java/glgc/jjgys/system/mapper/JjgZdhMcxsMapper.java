package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgZdhMcxs;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-06-06
 */
@Mapper
public interface JjgZdhMcxsMapper extends BaseMapper<JjgZdhMcxs> {

    List<Map<String,Object>> selectSdZfData(String proname, String htd, String zhq, String zhz,String zhq1,String zhz1);

    List<Map<String,Object>> selectSdZfDatayh(String proname, String htd, String zhq, String zhz, String zhq1, String zhz1, List<Long> idlist);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String,Object>> selectSdyfData(String proname, String htd, String zhq, String zhz,String zhq1,String zhz1);

    List<Map<String,Object>> selectSdyfDatayh(String proname, String htd, String zhq, String zhz, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String,Object>>  selectQlZfData(String proname, String htd, String zhq, String zhz,String zhq1,String zhz1);

    List<Map<String,Object>>  selectQlZfDatayh(String proname, String htd, String zhq, String zhz, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String,Object>>  selectQlYfData(String proname, String htd, String zhq, String zhz,String zhq1,String zhz1);

    List<Map<String,Object>>  selectQlYfDatayh(String proname, String htd, String zhq, String zhz, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String,Object>> selectzfList(String proname, String htd,String zx);

    //主线左幅
    List<Map<String,Object>> selectzfListyh(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String,Object>> selectyfList(String proname, String htd,String zx);

    //主线右幅
    List<Map<String,Object>> selectyfListyh(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String, Object>> selectlx(String proname, String htd);

    List<Map<String, Object>> selectsdmcxs(String proname,String bz,String lf, String sdzhq, String sdzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selectsdmcxsyh(String proname, String bz, String lf, String sdzhq, String sdzhz, String zx, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String, Object>> selectqlmcxs(String proname,String bz,String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selectqlmcxsyh(String proname,String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1,List<Long> idlist);

    int selectnum(String proname, String htd);

    List<Map<String, Object>> selectlxid(String proname, String htd, List<Long> idlist);

    int selectcdnumid(String proname, String htd, String zx, List<Long> idlist);

    List<Map<String, Object>> selectqlmcxsyhx(String proname, String lf,String wz, String qlzhq, String qlzhz, String qlname, String zhq1, String zhz1, List<Long> idlist);

    List<Map<String, Object>>  selectsdmcxsyhx(String proname, String lf, String wz, String sdzhq, String sdzhz, String sdname, String zhq1, String zhz1, List<Long> idlist);

}
