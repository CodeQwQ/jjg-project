package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhlJgfc;
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
 * @since 2024-05-20
 */
public interface JjgFbgcJtaqssJabxfhlJgfcService extends IService<JjgFbgcJtaqssJabxfhlJgfc> {

    boolean generateJdb(CommonInfoVo commonInfoVo);

    void importjabxfhl(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException;

    void exportjabxfhl(HttpServletResponse response) throws IOException;


    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;


    List<String> selecthtd(String proname);
}
