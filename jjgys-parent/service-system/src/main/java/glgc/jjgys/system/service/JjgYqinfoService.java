package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgYqinfo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-12-05
 */
public interface JjgYqinfoService extends IService<JjgYqinfo> {

    void export(HttpServletResponse response, String projectname);

    void importyqxx(MultipartFile file, String proname);

}
