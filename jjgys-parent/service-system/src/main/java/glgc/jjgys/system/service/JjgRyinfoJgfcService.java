package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgRyinfoJgfc;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2024-05-02
 */
public interface JjgRyinfoJgfcService extends IService<JjgRyinfoJgfc> {

    void export(HttpServletResponse response, String projectname);

    void importryxx(MultipartFile file, String proname);
}
