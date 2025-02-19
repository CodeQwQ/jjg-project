package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsQlMapper;
import glgc.jjgys.system.service.JjgLqsQlService;
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
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2022-12-14
 */
@Service
public class JjgLqsQlServiceImpl extends ServiceImpl<JjgLqsQlMapper, JjgLqsQl> implements JjgLqsQlService {

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;


    @Override
    @Transactional
    public void importQL(MultipartFile file,String proname)  {

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
                                        if (StringUtils.isEmpty(ql.getQlqc())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，桥梁全长为空，请修改后重新上传");
                                        }
                                        JjgLqsQl jjgLqsQl = new JjgLqsQl();
                                        BeanUtils.copyProperties(ql,jjgLqsQl);
                                        jjgLqsQl.setZhq(Double.valueOf(ql.getZhq()));
                                        jjgLqsQl.setZhz(Double.valueOf(ql.getZhz()));
                                        jjgLqsQl.setProname(proname);
                                        jjgLqsQl.setCreateTime(new Date());
                                        jjgLqsQlMapper.insert(jjgLqsQl);
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

    @Override
    public void exportQL(HttpServletResponse response,String proname) {
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
    public List<JjgLqsQl> getqlName(String proname, String htd) {
        List<JjgLqsQl> list = jjgLqsQlMapper.getqlName(proname,htd);
        return list;
    }

    @Override
    public List<String> getPureQlName(String proname, String htd) {

        List<String> qlName = jjgLqsQlMapper.getPureQlName(proname,htd);
        return qlName.stream().distinct().collect(Collectors.toList());
    }
}
