package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgSfz;
import glgc.jjgys.model.projectvo.lqs.SfzVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsSfzMapper;
import glgc.jjgys.system.service.JjgLqsSfzService;
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
public class JjgLqsSfzServiceImpl extends ServiceImpl<JjgLqsSfzMapper, JjgSfz> implements JjgLqsSfzService {

    @Autowired
    private JjgLqsSfzMapper jjgLqsSfzMapper;

    @Override
    public void exportSD(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "收费站清单";
            String sheetName = "收费站清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new SfzVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"收费站清单.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    @Transactional
    public void importSD(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(SfzVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<SfzVo>(SfzVo.class) {
                                @Override
                                public void handle(List<SfzVo> dataList) {
                                    int rowNumber=2;
                                    for(SfzVo sfz: dataList)
                                    {
                                        if (StringUtils.isEmpty(sfz.getZhq())) {
                                            throw new JjgysException(20001, "您上传收费站清单的第"+rowNumber+"行的数据中，桩号起为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sfz.getZhz())) {
                                            throw new JjgysException(20001, "您上传收费站清单的第"+rowNumber+"行的数据中，桩号止为空，请修改后重新上传");
                                        }
                                        JjgSfz jjgSfz = new JjgSfz();
                                        BeanUtils.copyProperties(sfz,jjgSfz);
                                        String zhq = sfz.getZhq();
                                        String zhz = sfz.getZhz();
                                        jjgSfz.setZhq(Double.valueOf(zhq));
                                        jjgSfz.setZhz(Double.valueOf(zhz));
                                        jjgSfz.setProname(proname);
                                        jjgSfz.setCreateTime(new Date());
                                        jjgLqsSfzMapper.insert(jjgSfz);
                                        rowNumber++;
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
