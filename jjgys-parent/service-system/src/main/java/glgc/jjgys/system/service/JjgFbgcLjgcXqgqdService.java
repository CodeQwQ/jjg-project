package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcXqgqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-02-13
 */
public interface JjgFbgcLjgcXqgqdService extends IService<JjgFbgcLjgcXqgqd> {

    boolean generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportxqgqd(HttpServletResponse response);

    void importxqgqd(MultipartFile file,CommonInfoVo commonInfoVo);

    List<String> selectsjqd(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);

    int selectnum(String proname, String htd);

    int selectnumname(String proname);

    int createMoreRecords(List<JjgFbgcLjgcXqgqd> data, String userID);

}
