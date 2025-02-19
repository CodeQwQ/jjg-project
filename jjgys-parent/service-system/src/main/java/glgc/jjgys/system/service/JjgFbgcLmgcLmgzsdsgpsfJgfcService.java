package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLmgcLmgzsdsgpsfJgfc;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
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
public interface JjgFbgcLmgcLmgzsdsgpsfJgfcService extends IService<JjgFbgcLmgcLmgzsdsgpsfJgfc> {

    boolean generateJdb(String proname) throws IOException;

    void exportlmgzsdsgpsf(HttpServletResponse response);

    void importlmgzsdsgpsf(MultipartFile file, String proname, String username);

    List<Map<String, Object>> selecthtd(String proname);

    int createMoreRecords(List<JjgFbgcLmgcLmgzsdsgpsfJgfc> data, String userID);

    List<Map<String, Object>> lookJdbjg(String proname) throws IOException;

}
