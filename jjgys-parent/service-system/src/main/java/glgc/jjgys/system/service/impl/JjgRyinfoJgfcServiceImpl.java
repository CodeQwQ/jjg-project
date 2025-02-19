package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgRyinfoJgfc;
import glgc.jjgys.model.projectvo.lqs.JjgRyinfoJgfcVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgRyinfoJgfcMapper;
import glgc.jjgys.system.service.JjgRyinfoJgfcService;
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
 * @since 2024-05-02
 */
@Service
public class JjgRyinfoJgfcServiceImpl extends ServiceImpl<JjgRyinfoJgfcMapper, JjgRyinfoJgfc> implements JjgRyinfoJgfcService {

    @Autowired
    private JjgRyinfoJgfcMapper jjgRyinfoJgfcMapper;

    @Override
    public void export(HttpServletResponse response, String projectname) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = projectname+"人员信息";
            String sheetName = "人员信息";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgRyinfoJgfcVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+projectname+"人员信息.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    @Transactional
    public void importryxx(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgRyinfoJgfcVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgRyinfoJgfcVo>(JjgRyinfoJgfcVo.class) {
                                @Override
                                public void handle(List<JjgRyinfoJgfcVo> dataList) {
                                    for(JjgRyinfoJgfcVo ryinfo: dataList)
                                    {
                                        if (ryinfo.getName() != null && ryinfo.getZw() !=null && ryinfo.getJszc()!=null && ryinfo.getZgzsbh()!=null){
                                            JjgRyinfoJgfc ry = new JjgRyinfoJgfc();
                                            BeanUtils.copyProperties(ryinfo,ry);
                                            ry.setProname(proname);
                                            ry.setCreateTime(new Date());
                                            jjgRyinfoJgfcMapper.insert(ry);
                                        }

                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }
    }
}
