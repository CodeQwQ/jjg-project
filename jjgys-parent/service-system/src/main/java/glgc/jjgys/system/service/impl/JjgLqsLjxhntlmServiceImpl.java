package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLjxhntlm;
import glgc.jjgys.model.projectvo.lqs.LjxhntlmVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsLjxhntlmMapper;
import glgc.jjgys.system.service.JjgLqsLjxhntlmService;
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
 * @since 2023-01-10
 */
@Service
public class JjgLqsLjxhntlmServiceImpl extends ServiceImpl<JjgLqsLjxhntlmMapper, JjgLjxhntlm> implements JjgLqsLjxhntlmService {

    @Autowired
    private JjgLqsLjxhntlmMapper jjgLqsLjxhntlmMapper;

    @Override
    public void exportLjxhntlm(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "连接线混凝土路面清单";
            String sheetName = "连接线混凝土路面清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new LjxhntlmVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"连接线混凝土路面清单.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    public void importLjxhntlm(MultipartFile file, String proname) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(LjxhntlmVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<LjxhntlmVo>(LjxhntlmVo.class) {
                                @Override
                                public void handle(List<LjxhntlmVo> dataList) {
                                    for(LjxhntlmVo ljxhntlmVo: dataList)
                                    {
                                        JjgLjxhntlm jjgLjxhntlm = new JjgLjxhntlm();
                                        BeanUtils.copyProperties(ljxhntlmVo,jjgLjxhntlm);
                                        jjgLjxhntlm.setZhq(Double.valueOf(ljxhntlmVo.getZhq()));
                                        jjgLjxhntlm.setZhz(Double.valueOf(ljxhntlmVo.getZhz()));
                                        jjgLjxhntlm.setProname(proname);
                                        jjgLjxhntlm.setCreateTime(new Date());
                                        jjgLqsLjxhntlmMapper.insert(jjgLjxhntlm);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }
}
