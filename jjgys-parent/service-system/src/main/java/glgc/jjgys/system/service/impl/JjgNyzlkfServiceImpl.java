package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgNyzlkf;
import glgc.jjgys.model.projectvo.jggl.JjgNyzlkfVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgNyzlkfMapper;
import glgc.jjgys.system.service.JjgNyzlkfService;
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
 * @since 2023-09-25
 */
@Service
public class JjgNyzlkfServiceImpl extends ServiceImpl<JjgNyzlkfMapper, JjgNyzlkf> implements JjgNyzlkfService {

    @Autowired
    private JjgNyzlkfMapper jjgNyzlkfMapper;

    @Override
    public void export(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "内页资料扣分";
            String sheetName = "内页资料扣分";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgNyzlkfVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"内页资料扣分.xlsx");
            if (t.exists()){
                t.delete();
            }
        }


    }

    @Override
    @Transactional
    public void importnyzl(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgNyzlkfVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgNyzlkfVo>(JjgNyzlkfVo.class) {
                                @Override
                                public void handle(List<JjgNyzlkfVo> dataList) {
                                    for(JjgNyzlkfVo ql: dataList)
                                    {
                                        JjgNyzlkf jjgLqsQl = new JjgNyzlkf();
                                        BeanUtils.copyProperties(ql,jjgLqsQl);
                                        jjgLqsQl.setProname(proname);
                                        jjgLqsQl.setCreateTime(new Date());
                                        jjgNyzlkfMapper.insert(jjgLqsQl);
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
