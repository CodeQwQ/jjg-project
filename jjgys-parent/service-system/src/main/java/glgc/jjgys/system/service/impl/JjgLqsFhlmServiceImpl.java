package glgc.jjgys.system.service.impl;


import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsFhlm;
import glgc.jjgys.model.projectvo.lqs.FhlmVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsFhlmMapper;
import glgc.jjgys.system.service.JjgLqsFhlmService;
import org.apache.commons.lang3.StringUtils;
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
public class JjgLqsFhlmServiceImpl extends ServiceImpl<JjgLqsFhlmMapper, JjgLqsFhlm> implements JjgLqsFhlmService {

    @Autowired
    private JjgLqsFhlmMapper jjgLqsFhlmMapper;

    @Override
    public void export(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "复合路面清单";
            String sheetName = "复合路面清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new FhlmVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"复合路面清单.xlsx");
            if (t.exists()){
                t.delete();
            }
        }
    }

    @Override
    @Transactional
    public void importFhlm(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(FhlmVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<FhlmVo>(FhlmVo.class) {
                                @Override
                                public void handle(List<FhlmVo> dataList) {
                                    int rowNumber=2;
                                    for(FhlmVo fhlm: dataList)
                                    {
                                        if (StringUtils.isEmpty(fhlm.getZhq())) {
                                            throw new JjgysException(20001, "您上传复合路面清单的第"+rowNumber+"行的数据中，桩号起为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(fhlm.getZhz())) {
                                            throw new JjgysException(20001, "您上传复合路面清单的第"+rowNumber+"行的数据中，桩号止为空，请修改后重新上传");
                                        }
                                        JjgLqsFhlm jjgLqsFhlm = new JjgLqsFhlm();
                                        BeanUtils.copyProperties(fhlm,jjgLqsFhlm);
                                        jjgLqsFhlm.setZhq(Double.valueOf(fhlm.getZhq()));
                                        jjgLqsFhlm.setZhz(Double.valueOf(fhlm.getZhz()));
                                        jjgLqsFhlm.setProname(proname);
                                        jjgLqsFhlm.setCreateTime(new Date());
                                        jjgLqsFhlmMapper.insert(jjgLqsFhlm);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"请检查桩号是否正确或删除文件最后的空数据，然后重新上传");
        }

    }
}
