package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.project.JjgLjfrnum;
import glgc.jjgys.model.projectvo.lqs.JjgLjfrnumVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLjfrnumMapper;
import glgc.jjgys.system.service.JjgHtdService;
import glgc.jjgys.system.service.JjgLjfrnumService;
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
public class JjgLjfrnumServiceImpl extends ServiceImpl<JjgLjfrnumMapper, JjgLjfrnum> implements JjgLjfrnumService {

    @Autowired
    private JjgLjfrnumMapper jjgLjfrnumMapper;

    @Autowired
    private JjgHtdService jjgHtdService;

    @Override
    public void export(HttpServletResponse response, String proname) {
        String fileName = proname+"路基数量";
        String sheetName = "路基数量";
        List<JjgLjfrnumVo> list = new ArrayList<>();
        String[] ljlist = {"支挡","涵洞","通道","小桥"};


        List<JjgHtd> gethtd = jjgHtdService.gethtd(proname);
        for (JjgHtd jjgHtd : gethtd) {
            String lx = jjgHtd.getLx();
            String htd = jjgHtd.getName();
            if (lx.contains("路基工程")){
                for (String s : ljlist) {
                    JjgLjfrnumVo janum = new JjgLjfrnumVo();
                    janum.setProname(proname);
                    janum.setHtd(htd);
                    janum.setFbgc("路基工程");
                    janum.setZb(s);
                    list.add(janum);
                }
            }
        }
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName, new JjgLjfrnumVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+proname+"路基数量.xlsx");
            if (t.exists()){
                t.delete();
            }
        }


    }

    @Override
    @Transactional
    public void importlj(MultipartFile file, String proname) {
        try {
            QueryWrapper<JjgLjfrnum> wrapper = new QueryWrapper<>();
            wrapper.eq("proname",proname);
            List<JjgLjfrnum> jjgPlaninfos = jjgLjfrnumMapper.selectList(wrapper);
            if (jjgPlaninfos == null || jjgPlaninfos.size() <= 0){
                EasyExcel.read(file.getInputStream())
                        .sheet(0)
                        .head(JjgLjfrnumVo.class)
                        .headRowNumber(1)
                        .registerReadListener(
                                new ExcelHandler<JjgLjfrnumVo>(JjgLjfrnumVo.class) {
                                    @Override
                                    public void handle(List<JjgLjfrnumVo> dataList) {
                                        for(JjgLjfrnumVo ql: dataList)
                                        {
                                            JjgLjfrnum jjgPlaninfo = new JjgLjfrnum();
                                            BeanUtils.copyProperties(ql,jjgPlaninfo);
                                            jjgPlaninfo.setCreateTime(new Date());
                                            jjgLjfrnumMapper.insert(jjgPlaninfo);
                                        }
                                    }
                                }
                        ).doRead();
            }else {
                throw new JjgysException(20001,"数据已经存在，请删除后重新上传");
            }
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }
    }

    @Override
    public int selecthdnum(String proname, String htd) {
        QueryWrapper<JjgLjfrnum> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc","路基工程");
        wrapper.eq("zb","涵洞");
        JjgLjfrnum ljfrnum = jjgLjfrnumMapper.selectOne(wrapper);
        return  ljfrnum !=null && !"".equals(ljfrnum.getNum()) ? Integer.valueOf(ljfrnum.getNum()) : 0;

    }

    @Override
    public int selectzdnum(String proname, String htd) {
        QueryWrapper<JjgLjfrnum> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc","路基工程");
        wrapper.eq("zb","支挡");
        JjgLjfrnum ljfrnum = jjgLjfrnumMapper.selectOne(wrapper);
        return  ljfrnum !=null && !"".equals(ljfrnum.getNum()) ? Integer.valueOf(ljfrnum.getNum()) : 0;

    }

    @Override
    public int selectxqnum(String proname, String htd) {
        QueryWrapper<JjgLjfrnum> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc","路基工程");
        wrapper.eq("zb","小桥");
        JjgLjfrnum ljfrnum = jjgLjfrnumMapper.selectOne(wrapper);
        return  ljfrnum !=null && !"".equals(ljfrnum.getNum())  ? Integer.valueOf(ljfrnum.getNum()) : 0;
    }
}
