package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsJgQl;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsJgQlMapper;
import glgc.jjgys.system.service.JjgLqsJgQlService;
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
public class JjgLqsJgQlServiceImpl extends ServiceImpl<JjgLqsJgQlMapper, JjgLqsJgQl> implements JjgLqsJgQlService {

    @Autowired
    private JjgLqsJgQlMapper jjgLqsJgQlMapper;

    @Override
    public void exportQL(HttpServletResponse response, String proname) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = proname+"桥梁清单";
            String sheetName = "桥梁清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new QlVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+proname+"桥梁清单.xlsx");
            if (t.exists()){
                t.delete();
            }
        }
    }

    @Override
    public void importQL(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(QlVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<QlVo>(QlVo.class) {
                                @Override
                                public void handle(List<QlVo> dataList) {
                                    int rowNumber=2;
                                    for(QlVo ql: dataList)
                                    {
                                        if (StringUtils.isEmpty(ql.getZhq())) {
                                            throw new JjgysException(20001, "您上传桥梁清单的第"+rowNumber+"行的数据中，桩号起为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(ql.getZhz())) {
                                            throw new JjgysException(20001, "您上传桥梁清单的第"+rowNumber+"行的数据中，桩号止为空，请修改后重新上传");
                                        }
                                        JjgLqsJgQl jjgLqsQl = new JjgLqsJgQl();
                                        BeanUtils.copyProperties(ql,jjgLqsQl);
                                        jjgLqsQl.setZhq(Double.valueOf(ql.getZhq()));
                                        jjgLqsQl.setZhz(Double.valueOf(ql.getZhz()));
                                        jjgLqsQl.setProname(proname);
                                        jjgLqsQl.setCreateTime(new Date());
                                        jjgLqsJgQlMapper.insert(jjgLqsQl);
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
