package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgJgjcsj;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-09-25
 */
public interface JjgJgjcsjService extends IService<JjgJgjcsj> {

    void exportjgjcdata(HttpServletResponse response, String proname) throws IOException;


    void importjgsj(MultipartFile file, String projectname);

    boolean generatepdb(String projectname);

    boolean generatepdbOld(String proname);

    boolean generateword(String proname) throws IOException;

}
