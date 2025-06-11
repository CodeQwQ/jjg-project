package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
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
 * @since 2023-04-10
 */
public interface JjgFbgcLmgcLqlmysdService extends IService<JjgFbgcLmgcLqlmysd> {

    boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportlqlmysd(HttpServletResponse response);

    void importlqlmysd(MultipartFile file, CommonInfoVo commonInfoVo);

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    int createMoreRecords(List<JjgFbgcLmgcLqlmysd> data, String userID);

    List<Map<String, Object>> selectsfl(String proname, String htd, List<Long> idlist);

    List<Map<String, Object>> lookJdbjgbgzbg(CommonInfoVo commonInfoVo, int flag) throws IOException;
}
