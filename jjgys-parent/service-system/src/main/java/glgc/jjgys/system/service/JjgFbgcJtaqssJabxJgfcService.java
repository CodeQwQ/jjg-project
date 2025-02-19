package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxJgfc;
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
 * @since 2024-05-20
 */
public interface JjgFbgcJtaqssJabxJgfcService extends IService<JjgFbgcJtaqssJabxJgfc> {

    boolean generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportjabx(HttpServletResponse response);

    void importjabx(MultipartFile file, CommonInfoVo commonInfoVo);

    List<String> selecthtd(String proname);
}
