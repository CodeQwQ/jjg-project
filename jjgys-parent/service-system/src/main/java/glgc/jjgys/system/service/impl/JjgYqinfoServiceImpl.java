package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgYqinfo;
import glgc.jjgys.model.projectvo.lqs.JjgYqinfoVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgYqinfoMapper;
import glgc.jjgys.system.service.JjgYqinfoService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.List;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-12-05
 */
@Service
public class JjgYqinfoServiceImpl extends ServiceImpl<JjgYqinfoMapper, JjgYqinfo> implements JjgYqinfoService {

    @Autowired
    private JjgYqinfoMapper jjgYqinfoMapper;

    @Override
    public void export(HttpServletResponse response, String projectname) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = projectname+"仪器信息";
            String sheetName = "仪器信息";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgYqinfoVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+projectname+"仪器信息.xlsx");
            if (t.exists()){
                t.delete();
            }
        }


    }

    @Override
    @Transactional
    public void importyqxx(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgYqinfoVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgYqinfoVo>(JjgYqinfoVo.class) {
                                @Override
                                public void handle(List<JjgYqinfoVo> dataList) {
                                    for(JjgYqinfoVo ryinfo: dataList)
                                    {
                                        JjgYqinfo ry = new JjgYqinfo();
                                        BeanUtils.copyProperties(ryinfo,ry);
                                        ry.setProname(proname);
                                        ry.setCreateTime(new Date());
                                        jjgYqinfoMapper.insert(ry);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }


    }
}
