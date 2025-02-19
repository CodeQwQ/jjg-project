package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdSl;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
public interface JjgFbgcLjgcLjtsfysdSlService extends IService<JjgFbgcLjgcLjtsfysdSl> {

    List<JjgFbgcLjgcLjtsfysdSl> getdata(String proname, String htd, String fbgc, List<Long> idlist);

    void exportysdsl(HttpServletResponse response) throws IOException;

    void importysdsl(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException;

    LinkedHashMap<String, ArrayList<String>> writeAndGetData(XSSFWorkbook wb, String proname, String htd, String fbgc, List<Long> idlist);

    int selectnum(String proname, String htd);

    int createMoreRecords(List<JjgFbgcLjgcLjtsfysdSl> data, String userID);
}
