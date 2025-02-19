package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcSdgcTlmxlbgc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-05-04
 */
public interface JjgFbgcSdgcTlmxlbgcService extends IService<JjgFbgcSdgcTlmxlbgc> {

    boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportsdxlbgs(HttpServletResponse response);

    void importsdxlbgs(MultipartFile file, CommonInfoVo commonInfoVo);

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc, String userid);

    List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo, String value) throws IOException;

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    List<Map<String, Object>> lookJdbjgPage(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> getsdnum(String proname, String htd);

    int createMoreRecords(List<JjgFbgcSdgcTlmxlbgc> data, String userID);

}
