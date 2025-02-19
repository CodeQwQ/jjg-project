package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcJgfc;
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
 * @since 2024-04-30
 */
public interface JjgFbgcLmgcLmwcJgfcService extends IService<JjgFbgcLmgcLmwcJgfc> {

    List<Map<String, Object>> selecthtd(String proname);

    boolean generateJdb(String commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportlmwc(HttpServletResponse response);

    void importlmwc(MultipartFile file, CommonInfoVo commonInfoVo);

}
