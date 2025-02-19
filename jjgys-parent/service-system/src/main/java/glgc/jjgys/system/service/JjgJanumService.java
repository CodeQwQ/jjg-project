package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgJanum;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-10-21
 */
public interface JjgJanumService extends IService<JjgJanum> {

    void export(HttpServletResponse response, String projectname);

    void importinfo(MultipartFile file, String proname);

    int selectbznum(String proname, String htd);

    int selectbxnum(String proname, String htd);

    int selectfhlnum(String proname, String htd);
}
