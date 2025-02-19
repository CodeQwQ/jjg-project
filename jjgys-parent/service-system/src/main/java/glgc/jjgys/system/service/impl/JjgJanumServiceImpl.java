package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.project.JjgJanum;
import glgc.jjgys.model.projectvo.lqs.JjgJanumVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgJanumMapper;
import glgc.jjgys.system.service.JjgHtdService;
import glgc.jjgys.system.service.JjgJanumService;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-10-21
 */
@Service
public class JjgJanumServiceImpl extends ServiceImpl<JjgJanumMapper, JjgJanum> implements JjgJanumService {

    @Autowired
    private JjgJanumMapper jjgJanumMapper;

    @Autowired
    private JjgHtdService jjgHtdService;


    @Override
    public void export(HttpServletResponse response, String proname) {
        String fileName = proname+"交安设施数量";
        String sheetName = "交安设施数量";
        List<JjgJanumVo> list = new ArrayList<>();
        String[] jalist = {"单悬","双悬","附着","门架","单柱","双柱"};
        String[] jalist1 = {"标志(实有数)","标线(公里数)","防护栏(实有数)"};


        List<JjgHtd> gethtd = jjgHtdService.gethtd(proname);
        for (JjgHtd jjgHtd : gethtd) {
            String lx = jjgHtd.getLx();
            String htd = jjgHtd.getName();
            if (lx.contains("交安工程")){
                for (String s : jalist) {
                    JjgJanumVo janum = new JjgJanumVo();
                    janum.setProname(proname);
                    janum.setHtd(htd);
                    janum.setFbgc("交安工程");
                    janum.setZb(s);
                    list.add(janum);
                }
                for (String s : jalist1) {
                    JjgJanumVo janum = new JjgJanumVo();
                    janum.setProname(proname);
                    janum.setHtd(htd);
                    janum.setFbgc("交安工程");
                    janum.setZb(s);
                    list.add(janum);
                }

            }
        }
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName, new JjgJanumVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+proname+"交安设施数量.xlsx");
            if (t.exists()){
                t.delete();
            }
        }



    }

    @Override
    @Transactional
    public void importinfo(MultipartFile file, String proname) {
        try {
            QueryWrapper<JjgJanum> wrapper = new QueryWrapper<>();
            wrapper.eq("proname",proname);
            List<JjgJanum> jjgPlaninfos = jjgJanumMapper.selectList(wrapper);
            if (jjgPlaninfos == null || jjgPlaninfos.size() <= 0){
                EasyExcel.read(file.getInputStream())
                        .sheet(0)
                        .head(JjgJanumVo.class)
                        .headRowNumber(1)
                        .registerReadListener(
                                new ExcelHandler<JjgJanumVo>(JjgJanumVo.class) {
                                    @Override
                                    public void handle(List<JjgJanumVo> dataList) {
                                        int rowNumber = 2;
                                        for(JjgJanumVo ql: dataList)
                                        {
                                            if (StringUtils.isEmpty(ql.getNum())) {
                                                //第2行的数据中，部位1值为空，请修改后重新上传
                                                throw new JjgysException(20001, "第"+rowNumber+"行的数据中，数量为空，请修改后重新上传");
                                            }
                                            JjgJanum jjgPlaninfo = new JjgJanum();
                                            BeanUtils.copyProperties(ql,jjgPlaninfo);
                                            jjgPlaninfo.setCreateTime(new Date());
                                            jjgJanumMapper.insert(jjgPlaninfo);
                                            rowNumber++;
                                        }
                                    }
                                }
                        ).doRead();
            }else {
                throw new JjgysException(20001,"数据已经存在，请删除后重新上传");
            }
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public int selectbznum(String proname, String htd) {
        QueryWrapper<JjgJanum> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc","交安工程");
        wrapper.eq("zb","标志(实有数)");
        JjgJanum jjgJanum = jjgJanumMapper.selectOne(wrapper);
        return  jjgJanum !=null && !"".equals(jjgJanum.getNum())  ? Integer.valueOf(jjgJanum.getNum()) : 0;
    }

    @Override
    public int selectbxnum(String proname, String htd) {
        QueryWrapper<JjgJanum> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc","交安工程");
        wrapper.eq("zb","标线(公里数)");
        JjgJanum jjgJanum = jjgJanumMapper.selectOne(wrapper);
        return  jjgJanum !=null && !"".equals(jjgJanum.getNum())   ? Integer.valueOf(jjgJanum.getNum()) : 0;
    }

    @Override
    public int selectfhlnum(String proname, String htd) {
        QueryWrapper<JjgJanum> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc","交安工程");
        wrapper.eq("zb","防护栏(实有数)");
        JjgJanum jjgJanum = jjgJanumMapper.selectOne(wrapper);
        return  jjgJanum !=null && !"".equals(jjgJanum.getNum())  ? Integer.valueOf(jjgJanum.getNum()) : 0;
    }
}
