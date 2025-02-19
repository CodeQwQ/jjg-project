package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLmgcTlmxlbgcJgfc;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
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
 * @since 2023-09-23
 */
public interface JjgFbgcLmgcTlmxlbgcJgfcService extends IService<JjgFbgcLmgcTlmxlbgcJgfc> {

    void exportxlbgs(HttpServletResponse response);

    void importxlbgs(MultipartFile file, String proname);

    boolean generateJdb(String proname) throws IOException, ParseException;

    List<Map<String, Object>> selecthtd(String proname);

    int createMoreRecords(List<JjgFbgcLmgcTlmxlbgcJgfc> data, String userID);

    List<Map<String, Object>> lookJdbjg(String proname) throws IOException;
}
