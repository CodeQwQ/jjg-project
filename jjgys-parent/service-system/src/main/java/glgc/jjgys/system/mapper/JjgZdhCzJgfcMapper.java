package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgZdhCzJgfc;
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
public interface JjgZdhCzJgfcMapper extends BaseMapper<JjgZdhCzJgfc> {

    List<String> gethtd(String proname);

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String, Object>> selectzfList(String proname, String htd, String zx);
    List<Map<String, Object>> selectzfListyh(String proname, String htd, String zx, String username);

    List<Map<String, Object>> selectyfList(String proname, String htd, String zx);
    List<Map<String, Object>> selectyfListyh(String proname, String htd, String zx, String username);

    /**
     * 主线左幅隧道
     * @param proname
     * @param htd
     * @param zx
     * @param zhq
     * @param zhz
     * @return
     */
    List<Map<String, Object>> selectSdZfData(String proname, String htd, String zx, String zhq, String zhz);
    
    List<Map<String, Object>> selectSdZfDatayh(String proname, String htd, String zx, String zhq, String zhz, String username);

    /**
     * 主线右幅隧道
     * @param proname
     * @param htd
     * @param zx
     * @param zhq
     * @param zhz
     * @return
     */
    List<Map<String, Object>> selectSdyfData(String proname, String htd, String zx, String zhq, String zhz);
    
    List<Map<String, Object>> selectSdyfDatayh(String proname, String htd, String zx, String zhq, String zhz, String username);

    /**
     * 主线左幅桥梁
     * @param proname
     * @param htd
     * @param zx
     * @param zhq
     * @param zhz
     * @return
     */
    List<Map<String, Object>> selectQlZfData(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectQlZfDatayh(String proname, String htd, String zx, String zhq, String zhz, String username);

    /**
     * 主线右幅桥梁
     * @param proname
     * @param htd
     * @param zx
     * @param zhq
     * @param zhz
     * @return
     */
    List<Map<String, Object>> selectQlYfData(String proname, String htd, String zx, String zhq, String zhz);

    List<Map<String, Object>> selectQlYfDatayh(String proname, String htd, String zx, String zhq, String zhz, String username);

    /**
     * 连接线隧道
     * @param proname
     * @param lf
     * @param wz
     * @param sdzhq
     * @param sdz
     * @param sdname
     * @param zhq1
     * @param zhz1
     * @return
     */
    List<Map<String, Object>> selectsdcz(String proname, String lf, String wz,String sdzhq, String sdz, String sdname, String zhq1, String zhz1);

    /**
     * 匝道隧道
     * @param proname
     * @param bz
     * @param lf
     * @param sdzhq
     * @param sdz
     * @param zx
     * @param zhq1
     * @param zhz1
     * @return
     */
    List<Map<String, Object>> selectsdczyh(String proname, String bz, String lf, String sdzhq, String sdz, String zx, String zhq1, String zhz1);

    /**
     * 连接线桥梁
     * @param proname
     * @param lf
     * @param qlzhq
     * @param qlzhz
     * @param qlname
     * @param zhq1
     * @param zhz1
     * @return
     */
    List<Map<String, Object>> selectqlcz(String proname, String lf, String wz, String qlzhq, String qlzhz, String qlname, String zhq1, String zhz1);

    /**
     * 匝道桥梁
     * @param proname
     * @param bz
     * @param lf
     * @param qlzhq
     * @param qlzhz
     * @param zx
     * @param zhq1
     * @param zhz1
     * @return
     */
    List<Map<String, Object>> selectqlczyh(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);

    List<Map<String, Object>> selecthtd(String proname);

}
