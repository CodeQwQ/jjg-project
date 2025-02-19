package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsJgSd;
import glgc.jjgys.model.projectvo.lqs.SdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsJgSdMapper;
import glgc.jjgys.system.service.JjgLqsJgSdService;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
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
 * @since 2023-09-22
 */
@Service
public class JjgLqsJgSdServiceImpl extends ServiceImpl<JjgLqsJgSdMapper, JjgLqsJgSd> implements JjgLqsJgSdService {

    @Autowired
    private JjgLqsJgSdMapper jjgLqsJgSdMapper;

    @Override
    public void exportSD(HttpServletResponse response, String projectname) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = projectname+"隧道清单";
            String sheetName = "隧道清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new SdVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+projectname+"隧道清单.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    public void importSD(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(SdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<SdVo>(SdVo.class) {
                                @Override
                                public void handle(List<SdVo> dataList) {
                                    int rowNumber=2;
                                    for(SdVo sd: dataList)
                                    {
                                        if (StringUtils.isEmpty(sd.getZhq())) {
                                            throw new JjgysException(20001, "您上传隧道清单的第"+rowNumber+"行的数据中，桩号起为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sd.getZhz())) {
                                            throw new JjgysException(20001, "您上传隧道清单的第"+rowNumber+"行的数据中，桩号止为空，请修改后重新上传");
                                        }
                                        JjgLqsJgSd jjgLqsSd = new JjgLqsJgSd();
                                        BeanUtils.copyProperties(sd,jjgLqsSd);
                                        jjgLqsSd.setZhq(Double.valueOf(sd.getZhq()));
                                        jjgLqsSd.setZhz(Double.valueOf(sd.getZhz()));
                                        jjgLqsSd.setProname(proname);
                                        jjgLqsSd.setCreateTime(new Date());
                                        jjgLqsJgSdMapper.insert(jjgLqsSd);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
