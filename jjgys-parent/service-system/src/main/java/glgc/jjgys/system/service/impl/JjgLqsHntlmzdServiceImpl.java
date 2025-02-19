package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsHntlmzd;
import glgc.jjgys.model.projectvo.lqs.HntlmzdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsHntlmzdMapper;
import glgc.jjgys.system.service.JjgLqsHntlmzdService;
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
 * @since 2023-01-10
 */
@Service
public class JjgLqsHntlmzdServiceImpl extends ServiceImpl<JjgLqsHntlmzdMapper, JjgLqsHntlmzd> implements JjgLqsHntlmzdService {

    @Autowired
    private JjgLqsHntlmzdMapper jjgLqsHntlmzdMapper;

    @Override
    public void export(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "混凝土路面及匝道清单";
            String sheetName = "混凝土路面及匝道清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new HntlmzdVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"混凝土路面及匝道清单.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    @Transactional
    public void importhntlmzd(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(HntlmzdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<HntlmzdVo>(HntlmzdVo.class) {
                                @Override
                                public void handle(List<HntlmzdVo> dataList) {
                                    for (HntlmzdVo hntlmzdVo : dataList) {
                                        JjgLqsHntlmzd jjgLqsHntlmzd = new JjgLqsHntlmzd();
                                        BeanUtils.copyProperties(hntlmzdVo, jjgLqsHntlmzd);
                                        jjgLqsHntlmzd.setProname(proname);
                                        jjgLqsHntlmzd.setCreateTime(new Date());
                                        jjgLqsHntlmzdMapper.insert(jjgLqsHntlmzd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001, "解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"请检查桩号是否正确或删除文件最后的空数据，然后重新上传");
        }
    }
}
