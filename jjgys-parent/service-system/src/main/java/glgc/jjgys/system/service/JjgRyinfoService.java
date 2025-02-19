package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgRyinfo;
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
public interface JjgRyinfoService extends IService<JjgRyinfo> {

    void export(HttpServletResponse response, String projectname);

    void importryxx(MultipartFile file, String proname);

}
