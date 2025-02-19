package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.spire.doc.Table;
import com.spire.doc.*;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.TableRowHeightType;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.documents.VerticalAlignment;
import com.spire.doc.fields.TextRange;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.jggl.JCSJVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgJgHtdinfoMapper;
import glgc.jjgys.system.mapper.JjgJgjcsjMapper;
import glgc.jjgys.system.mapper.JjgLqsJgQlMapper;
import glgc.jjgys.system.mapper.JjgLqsJgSdMapper;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import glgc.jjgys.system.utils.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-09-25
 */
@Service
public class JjgJgjcsjServiceImpl extends ServiceImpl<JjgJgjcsjMapper, JjgJgjcsj> implements JjgJgjcsjService {

    @Autowired
    private JjgJgHtdinfoMapper jjgJgHtdinfoMapper;

    @Autowired
    private JjgLqsJgSdMapper jjgLqsJgSdMapper;

    @Autowired
    private JjgLqsJgQlMapper jjgLqsJgQlMapper;

    @Autowired
    private JjgJgjcsjMapper jjgJgjcsjMapper;

    @Autowired
    private JjgWgkfService jjgWgkfService;

    @Autowired
    private JjgDwgctzeService jjgDwgctzeService;

    @Autowired
    private JjgNyzlkfService jjgNyzlkfService;

    @Autowired
    private JjgFbgcLmgcLmwcJgfcService jjgFbgcLmgcLmwcJgfcService;

    @Autowired
    private JjgFbgcLmgcLmwcLcfJgfcService jjgFbgcLmgcLmwcLcfJgfcService;

    @Autowired
    private JjgZdhCzJgfcService jjgZdhCzJgfcService;

    @Autowired
    private JjgZdhPzdJgfcService jjgZdhPzdJgfcService;

    @Autowired
    private JjgZdhGzsdJgfcService jjgZdhGzsdJgfcService;

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfJgfcService jjgFbgcLmgcLmgzsdsgpsfJgfcService;

    @Autowired
    private JjgZdhMcxsJgfcService jjgZdhMcxsJgfcService;

    @Autowired
    private JjgFbgcLmgcTlmxlbgcJgfcService jjgFbgcLmgcTlmxlbgcJgfcService;

    @Autowired
    private JjgJgProjectinfoService projectinfoService;

    @Autowired
    private JjgJgHtdinfoService jjgJgHtdinfoService;

    @Autowired
    private JjgFbgcJtaqssJabxJgfcService jjgFbgcJtaqssJabxJgfcService;

    @Autowired
    private JjgFbgcJtaqssJabxfhlJgfcService jjgFbgcJtaqssJabxfhlJgfcService;

    @Autowired
    private JjgNyzlkfJgService jjgNyzlkfJgService;



    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @Override
    public void exportjgjcdata(HttpServletResponse response, String proname) throws IOException {
        QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgJgHtdinfo> htdinfo = jjgJgHtdinfoMapper.selectList(wrapper);
        List<JCSJVo> result = new ArrayList<>();
        if (htdinfo != null && !htdinfo.isEmpty()){
            for (JjgJgHtdinfo jjgJgHtdinfo : htdinfo) {
                String htd = jjgJgHtdinfo.getName();
                String lx = jjgJgHtdinfo.getLx();
                if (lx.contains("路基工程")){
                    List<JCSJVo> ljdata = getLjdata(htd);
                    result.addAll(ljdata);
                }
                if (lx.contains("路面工程")){
                    List<JCSJVo> lmdata = getLmdata(proname,htd);
                    result.addAll(lmdata);

                }
                if (lx.contains("桥梁工程")){
                    List<JCSJVo> qldata = getQldata(proname,htd);
                    result.addAll(qldata);

                }
                if (lx.contains("隧道工程")){
                    List<JCSJVo> sddata = getSddata(proname,htd);
                    result.addAll(sddata);

                }
                if (lx.contains("交安工程")){
                    List<JCSJVo> jadata = getJadata(htd);
                    result.addAll(jadata);

                }
            }
        }
        //往excel中写入
        exportExcel(response,proname,result);
    }

    @Override
    @Transactional
    public void importjgsj(MultipartFile file, String projectname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JCSJVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JCSJVo>(JCSJVo.class) {
                                @Override
                                public void handle(List<JCSJVo> dataList) {
                                    for(JCSJVo ql: dataList)
                                    {
                                        JjgJgjcsj jcsj = new JjgJgjcsj();
                                        BeanUtils.copyProperties(ql,jcsj);
                                        jcsj.setProname(projectname);
                                        jcsj.setCreateTime(new Date());
                                        jjgJgjcsjMapper.insert(jcsj);
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

    @Override
    public boolean generatepdb(String projectname) {
        QueryWrapper<JjgJgjcsj> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",projectname);
        List<JjgJgjcsj> jjgJgjcsjs = jjgJgjcsjMapper.selectList(wrapper);
        //filename 合格率

        //先按合同段，然后按单位工程分别写入到不同的工作簿
        if (jjgJgjcsjs != null && !jjgJgjcsjs.isEmpty()){
            List<Map<String,Object>> data = getdata(jjgJgjcsjs);
            Map<String, List<Map<String,Object>>> groupedData = data.stream()
                    .collect(Collectors.groupingBy(m -> m.get("htd").toString()));
            groupedData.forEach((group, grouphtdData) -> {
                try {
                    writeExceldata(group,grouphtdData,projectname);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
            return true;
        }else {
            return false;
        }
    }

    @Override
    public boolean generatepdbOld(String proname) {
        QueryWrapper<JjgJgjcsj> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgJgjcsj> jjgJgjcsjs = jjgJgjcsjMapper.selectList(wrapper);
        //filename 合格率

        //先按合同段，然后按单位工程分别写入到不同的工作簿
        if (jjgJgjcsjs != null && !jjgJgjcsjs.isEmpty()){
            List<Map<String,Object>> data = getdata(jjgJgjcsjs);
            Map<String, List<Map<String,Object>>> groupedData = data.stream()
                    .collect(Collectors.groupingBy(m -> m.get("htd").toString()));
            groupedData.forEach((group, grouphtdData) -> {
                try {
                    writeExceldataOld(group,grouphtdData,proname);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
            return true;
        }else {
            return false;
        }

    }



    @Override
    public boolean generateword(String proname) throws IOException {
        File f = new File(jgfilepath + File.separator + proname + File.separator + "高速公路竣工验收工程质量检测报告.docx");
        File fdir = new File(jgfilepath + File.separator + proname);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        //替换项目全称
        QueryWrapper<JjgJgProjectinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname", proname);
        JjgJgProjectinfo pro = projectinfoService.getOne(wrapper);

        InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/竣工报告6月.docx");

        Document xw = new Document(inputStream);

        xw.replace("${项目全称}", pro.getXmqc(), false, true);
        xw.replace("${项目名称}", pro.getProname(), false, true);


        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

        //表1 人员信息
        CreateTableryData(f,xw,proname);

        //${表2.2-1}
        CreateTableryData2(f,xw,proname);

        //${表1.3-1}
        DBtoLjqshtd(f,xw,proname);

        //${表1.3-2}
        DBtoLmhtd(f,xw,proname);

        //${表1.3-2}
        DBtojahtd(f,xw,proname);

        //竣工复测指标结果
        DBtoDocxJgfc(f,xw,proname);

        //表4.1-2  沥青路面弯沉复测结果统计表
        DBtoDocxlqwcfcjg(f,xw,proname);

        //表4.1-3  沥青路面车辙检测结果统计表
        DBtoDocxlqlmcz(f,xw,proname);

        //表4.1-4（1）  沥青路面平整度复测结果统计表
        DBtoDocxlqlmpzd(f,xw,proname);

        //表4.1-4（2）  混凝土路面平整度复测结果统计表
        DBtoDocxhntlmpzd(f,xw,proname);

        //表4.1-5  路面抗滑性能复测结果统计表（SFC）
        DBtoDocxlmmcxs(f,xw,proname);

        //表4.1-6（1）  沥青路面抗滑性能复测结果统计表（TD） 构造深度
        DBtoDocxlmgzsd(f,xw,proname);

        //表4.1-6（2）  混凝土路面抗滑性能复测结果统计表（TD）
        DBtoDocxhntlmgzsd(f,xw,proname);

        //表4.1-7（1）  沥青桥面铺装平整度复测结果统计表
        DBtoDocxlqqmpzd(f,xw,proname);

        //表4.1-7（2）  混凝土桥面铺装平整度复测结果统计表
        DBtoDocxhntqmpzd(f,xw,proname);

        //表4.1-8（1）  沥青桥面抗滑性能复测结果统计表（TD）
        DBtoDocxqmgzsd(f,xw,proname);

        //表4.1-8（2）  混凝土桥面抗滑性能复测结果统计表（TD）

        //表4.1-9  桥面抗滑性能复测结果统计表（SFC）
        DBtoDocxqmmcxs(f,xw,proname);

        //表4.1-10  隧道沥青路面平整度复测结果统计表
        DBtoDocxlqsdpzd(f,xw,proname);

        //表4.1-11  隧道路面抗滑性能复测结果统计表（SFC）
        DBtoDocxsdmcxs(f,xw,proname);

        //表4.1-12  隧道路面抗滑性能复测结果统计表（TD）
        DBtoDocxsdgzsd(f,xw,proname);

        //表4.1-13  隧道沥青路面车辙检测结果统计表
        DBtoDocxlqsdcz(f,xw,proname);

        //表4.1-14  砼路面相邻板高差检测结果统计表
        DBtoDocxtlmxlbgc(f,xw,proname);

        //表4.1-2
        DBtofcdb(f,xw,proname);

        //表4.1.3-1
        DBtojabx(f,xw,proname);

        //表4.1.3-2
        DBtojafhl(f,xw,proname);

        //表11.1.1-1
        DBtoljpdhzb(f,xw,proname);

        //表11.1.2-1
        DBtolmpdhzb(f,xw,proname);

        //表11.1.3-1
        DBtoqlpdhzb(f,xw,proname);

        //表11.1.4-1
        DBtosdpdhzb(f,xw,proname);

        //表11.1.5-1
        DBtojapdhzb(f,xw,proname);

        //表11.2-1
        DBtohtdpdhzb(f,xw,proname);

        //附表1桥梁清单
        CreateTableqlqdData(f, xw, proname);

        //附表2隧道清单
        CreateTablesdqdData(f, xw, proname);

        //外观检查
        CreateTablewgjcl(f, xw, proname);

        //内页资料
        CreateTablenyzl(f, xw, proname);


        return true;
    }


    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablewgjcl(File f, Document xw, String proname) {
        QueryWrapper<JjgWgkf> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd");
        List<JjgWgkf> list = jjgWgkfService.list(wrapper);

        //按单位工程分类
        if (list != null && list.size()>0) {
            List<JjgWgkf> ljgc = new ArrayList<>();
            List<JjgWgkf> lmgc = new ArrayList<>();
            List<JjgWgkf> sdgc = new ArrayList<>();
            List<JjgWgkf> qlgc = new ArrayList<>();
            List<JjgWgkf> jagc = new ArrayList<>();
            for (JjgWgkf jjgWgjc : list) {
                String dwgc = jjgWgjc.getDwgc();
                if (dwgc.contains("路基工程")) {
                    ljgc.add(jjgWgjc);

                } else if (dwgc.contains("隧道工程")) {
                    sdgc.add(jjgWgjc);

                } else if (dwgc.contains("路面工程")) {
                    lmgc.add(jjgWgjc);

                } else if (dwgc.contains("桥梁工程")) {
                    qlgc.add(jjgWgjc);

                } else if (dwgc.contains("交通安全设施工程")) {
                    jagc.add(jjgWgjc);

                }
            }
            extractedlj(f, xw, ljgc);
            /*extractedlm(f, xw, lmgc);
            extractedqlgc(f, xw, qlgc);
            extractedsd(f, xw, sdgc);
            extractedja(f, xw, jagc);*/
        }

    }

    /**
     *
     * @param f
     * @param xw
     * @param ljgc
     */
    private static void extractedlj(File f, Document xw, List<JjgWgkf> ljgc) {
        if (ljgc !=null && ljgc.size()>0) {
            String[] headerlj1 = {"合同段", "分部工程", "外观检查结果","","","","","工程量","分部工程扣分"};
            String[] headerlj2 = {"", "", "检查部位","缺陷描述","扣分单位","缺陷数量","扣分","",""};
            String lj = "${路基外观检查}";
            TextSelection textSelectionlj = xw.findString(lj, true, true);
            if (textSelectionlj != null){
                Table table = new Table(xw, true);
                table.resetCells(ljgc.size() + 2, 9);
                Paragraph paragraph = textSelectionlj.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < headerlj1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(headerlj1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 2) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 2, 6);
                    }

                }
                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < headerlj2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(headerlj2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0 || i == 1 || i == 7 || i == 8) {
                        table.applyVerticalMerge(i, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < ljgc.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    JjgWgkf rowData = ljgc.get(rowIdx);

                    for (int colIdx = 0; colIdx < headerlj1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.getHtd() + "合同段").getCharacterFormat().setFontSize(9f);

                                break;
                            case 1:
                                p.appendText(rowData.getFbgc()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                //p.appendText(rowData.getBhms()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                //p.appendText(rowData.getBhsl()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }
        }
    }


    @Autowired
    private JjgLqsJgSdService jjgLqsJgSdService;

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablesdqdData(File f, Document xw, String proname) {
        QueryWrapper<JjgLqsJgSd> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd","name");
        List<JjgLqsJgSd> list = jjgLqsJgSdService.list(wrapper);
        if (list != null){
            String[] header = {"合同段", "序号", "位置", "隧道名称", "起讫桩号","长度(m)","备注"};
            String s = "{附表2隧道清单}";
            //TextSelection textSelection = xw.findString(s, true, true);
            TextSelection textSelection = findStringInPages(xw,s,30,50);
            // 检查是否找到字符串
            if (textSelection != null) {
                Table table=new Table(xw,true);
                table.resetCells(list.size()+1, header.length);
                Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                Body body=paragraph.ownerTextBody();
                int index=body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    //row.getCells().get(i).setWidth(400);
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < list.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);

                    JjgLqsJgSd rowData = list.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.getHtd());
                                break;
                            case 1:
                                p.appendText(String.valueOf(rowIdx+1));
                                break;
                            case 2:
                                p.appendText(rowData.getLf());
                                break;
                            case 3:
                                p.appendText(rowData.getSdname());
                                break;
                            case 4:
                                p.appendText((int) Math.round(rowData.getZhq())+"~"+(int) Math.round(rowData.getZhz()));
                                break;
                            case 5:
                                p.appendText(rowData.getSdqc());
                                break;
                            case 6:
                                p.appendText("");
                                break;
                            default:
                                break;
                        }

                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index,table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    @Autowired
    private JjgLqsJgQlService jjgLqsJgQlService;

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTableqlqdData(File f, Document xw, String proname) {
        QueryWrapper<JjgLqsJgQl> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd","qlname");
        List<JjgLqsJgQl> list = jjgLqsJgQlService.list(wrapper);
        if (list != null){
            String[] header = {"合同段", "序号", "左幅/右幅", "桥梁名称", "起讫桩号","跨径组合（孔-米）","孔数","桥梁全长(米)","桥梁类型","位置","备注"};
            String s = "{附表1桥梁清单}";
            //TextSelection textSelection = xw.findString(s, true, true);
            TextSelection textSelection = findStringInPages(xw,s,31,50);
            // 检查是否找到字符串
            if (textSelection != null) {
                Table table=new Table(xw,true);
                table.resetCells(list.size()+1, header.length);
                Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                Body body=paragraph.ownerTextBody();
                int index=body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    //row.getCells().get(i).setWidth(400);
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < list.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);

                    JjgLqsJgQl rowData = list.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.getHtd());
                                break;
                            case 1:
                                p.appendText(String.valueOf(rowIdx+1));
                                break;
                            case 2:
                                p.appendText(rowData.getLf());
                                break;
                            case 3:
                                p.appendText(rowData.getQlname());
                                break;
                            case 4:
                                p.appendText((int) Math.round(rowData.getZhq())+"~"+(int) Math.round(rowData.getZhz()));
                                break;
                            case 5:
                                p.appendText(rowData.getDkkj());
                                break;
                            case 6:
                                p.appendText("");
                                break;
                            case 7:
                                p.appendText(rowData.getQlqc());
                                break;
                            case 8:
                                p.appendText(rowData.getPzlx());
                                break;
                            case 9:
                                p.appendText(rowData.getWz());
                                break;
                            case 10:
                                p.appendText("");
                                break;
                            default:
                                break;
                        }

                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index,table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }
    }


    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablenyzl(File f, Document xw, String proname) {
        QueryWrapper<JjgNyzlkfJg> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd");
        List<JjgNyzlkfJg> list = jjgNyzlkfJgService.list(wrapper);

        String[] header = {"合同段", "存在问题", "扣分"};

        String s = "${附表4内业资料}";
        //TextSelection textSelection = xw.findString(s, true, true);
        TextSelection textSelection = findStringInPages(xw, s, 30, 50);
        // 检查是否找到字符串
        if (textSelection != null) {
            Table table = new Table(xw, true);

            table.resetCells(list.size() + 1, header.length);
            Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
            Body body = paragraph.ownerTextBody();
            int index = body.getChildObjects().indexOf(paragraph);

            //将第一行设置为表格标题
            TableRow row = table.getRows().get(0);
            row.isHeader(true);
            row.setHeight(40);
            row.setHeightType(TableRowHeightType.Exactly);
            for (int i = 0; i < header.length; i++) {
                row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                Paragraph p = row.getCells().get(i).addParagraph();
                p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                TextRange txtRange = p.appendText(header[i]);
                txtRange.getCharacterFormat().setBold(true);
            }
            //添加数据
            for (int rowIdx = 0; rowIdx < list.size(); rowIdx++) {
                TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                row1.setHeightType(TableRowHeightType.Exactly);

                JjgNyzlkfJg rowData = list.get(rowIdx);

                for (int colIdx = 0; colIdx < header.length; colIdx++) {
                    row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(colIdx).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                    // 根据表头的不同，设置相应的数据
                    switch (colIdx) {
                        case 0:
                            p.appendText(rowData.getHtd());
                            break;
                        case 1:
                            p.appendText(rowData.getWt());
                            break;
                        case 2:
                            p.appendText(rowData.getKf());
                            break;
                        default:
                            break;
                    }
                }
            }

            body.getChildObjects().remove(paragraph);
            body.getChildObjects().insert(index, table);
            //列宽自动适应内容
            table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
            xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
        }


    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtohtdpdhzb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gethtdzlpdData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "", "", "", "", "单位工程", "", "", ""};
            String[] header2 = {"名称", "评分", "资料扣分", "质量得分", "质量等级", "名称", "得分", "投资额（万元）", "质量等级"};
            String s = "${表11.2-1}";
            TextSelection textSelection = findStringInPages(xw, s, 20, 45);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 9);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 0, 4);
                    }
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 5, 8);
                    }
                }
                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）

                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                break;
                            case 1:
                                p.appendText(rowData.get("htdpf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.get("zlkf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.get("zldf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.get("htdzldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("dwgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText(rowData.get("df").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText(rowData.get("tze").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }

                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);


            }
        }

    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gethtdzlpdData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        QueryWrapper<JjgJgHtdinfo> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String htd = jjgJgHtdinfo.getName();
                File ff = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                if (!ff.exists()) {
                    break;
                }else {
                    XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                    XSSFSheet slSheet1 = xwb.getSheet("合同段");
                    for (int i = 0; i < slSheet1.getPhysicalNumberOfRows(); i++){
                        CellType cellTypeEnum = slSheet1.getRow(i).getCell(0).getCellTypeEnum();


                        slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                        slSheet1.getRow(i).getCell(5).setCellType(CellType.STRING);//项目名称

                        String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                        String bzxmmc = slSheet1.getRow(i).getCell(5).getStringCellValue();

                        if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                            Map<String,Object> map = new HashMap<>();
                            slSheet1.getRow(i+25).getCell(4).setCellType(CellType.STRING);
                            slSheet1.getRow(i+25).getCell(6).setCellType(CellType.STRING);
                            slSheet1.getRow(i+26).getCell(4).setCellType(CellType.STRING);
                            slSheet1.getRow(i+26).getCell(6).setCellType(CellType.STRING);

                            map.put("htdpf",slSheet1.getRow(i+25).getCell(4).getStringCellValue());
                            map.put("zlkf",slSheet1.getRow(i+25).getCell(6).getStringCellValue());
                            map.put("zldf",slSheet1.getRow(i+26).getCell(4).getStringCellValue());
                            map.put("htdzldj",slSheet1.getRow(i+26).getCell(6).getStringCellValue());
                            resultList1.add(map);
                        }
                        if (cellTypeEnum == CellType.NUMERIC){
                            String value1 = slSheet1.getRow(i).getCell(1).getStringCellValue();
                            double value2 = slSheet1.getRow(i).getCell(3).getNumericCellValue();
                            double value3 = slSheet1.getRow(i).getCell(4).getNumericCellValue();
                            String value4 = slSheet1.getRow(i).getCell(6).getStringCellValue();
                            Map<String,Object> map = new HashMap<>();
                            map.put("htd",htd);
                            map.put("dwgc",value1);
                            map.put("df",value2);
                            map.put("tze",value3);
                            map.put("zldj",value4);
                            resultList.add(map);
                        }


                    }
                }
            }
        }
        if (resultList1 != null){
            for (Map<String, Object> objectMap : resultList1) {
                String htd = objectMap.get("htdpf").toString();
                String dwgc = objectMap.get("zlkf").toString();
                String dwgcdf = objectMap.get("zldf").toString();
                String dwgczldj = objectMap.get("htdzldj").toString();
                if (resultList != null){
                    for (Map<String, Object> stringObjectMap : resultList) {
                        stringObjectMap.put("htdpf",htd);
                        stringObjectMap.put("zlkf",dwgc);
                        stringObjectMap.put("zldf",dwgcdf);
                        stringObjectMap.put("htdzldj",dwgczldj);

                    }
                }
            }
        }

        //按合同段排序
        Collections.sort(resultList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("htd").toString();
                String name2 = o2.get("htd").toString();
                return name1.compareTo(name2);
            }
        });
        return resultList;

    }

    /**
     *
     * @param strNum
     * @return
     */
    public static boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        return strNum.matches("-?\\d+(\\.\\d+)?");
    }


    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtojapdhzb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getjapdbData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "单位工程", "", "", "分部工程", "", "", "", "", ""};
            String[] header2 = {"", "名称", "得分", "质量等级", "名称", "权值", "得分", "实测得分", "外观扣分", "质量等级"};
            String s = "${表11.1.5-1}";
            TextSelection textSelection = findStringInPages(xw, s, 20, 45);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 10);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 1, 3);
                    }
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 9);
                    }
                }
                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        table.applyVerticalMerge(0, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                if ((rowIdx ) % 3 == 0) {
                                    p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 1:
                                if ((rowIdx ) % 3 == 0) {
                                    p.appendText(rowData.get("dwgc").toString()).getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 2:
                                if ((rowIdx ) % 3 == 0) {
                                    p.appendText(rowData.get("dwgcdf").toString()).getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 3:
                                if ((rowIdx ) % 3 == 0) {
                                    p.appendText(rowData.get("dwgczldj").toString()).getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 4:
                                p.appendText(rowData.get("fbgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("qz").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText(rowData.get("df").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText(rowData.get("scdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText(rowData.get("wgkf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 9:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }
                int count = table.getRows().getCount();
                int startRow = 2; // 开始行（跳过表头或其他不想合并的行）
                while (startRow + 2 < count) { // 确保有足够的行进行合并
                    table.applyVerticalMerge(0, startRow, startRow + 2); // 仅合并第一列
                    table.applyVerticalMerge(1, startRow, startRow + 2); // 仅合并第一列
                    table.applyVerticalMerge(2, startRow, startRow + 2); // 仅合并第一列
                    table.applyVerticalMerge(3, startRow, startRow + 2); // 仅合并第一列
                    startRow += 3; // 移动到下一个要合并的5行组的起始行
                }



                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }
        }
    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getjapdbData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        QueryWrapper<JjgJgHtdinfo> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("交安工程")){
                    File ff = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    if (!ff.exists()) {
                        break;
                    }else {
                        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                        XSSFSheet slSheet1 = xwb.getSheet("分部-交通安全设施");
                        for (int i = 0; i < slSheet1.getPhysicalNumberOfRows()-20; i++){
                            slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                            slSheet1.getRow(i).getCell(15).setCellType(CellType.STRING);//项目名称
                            slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);//分部工程名称

                            String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                            String bzxmmc = slSheet1.getRow(i).getCell(15).getStringCellValue();

                            if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                                slSheet1.getRow(i+20).getCell(6).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(10).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(17).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","交通安全设施");
                                map.put("fbgc",fbgcname);
                                map.put("scdf",slSheet1.getRow(i+20).getCell(6).getStringCellValue());
                                map.put("wgkf",slSheet1.getRow(i+20).getCell(10).getStringCellValue());
                                map.put("df",slSheet1.getRow(i+20).getCell(17).getStringCellValue());
                                map.put("zldj",slSheet1.getRow(i+20).getCell(20).getStringCellValue());
                                if (fbgcname.equals("标志") || fbgcname.equals("标线") ){
                                    map.put("qz",1);
                                }else if (fbgcname.equals("防护栏")){
                                    map.put("qz",2);
                                }else {
                                    map.put("qz",0);
                                }
                                resultList.add(map);
                            }
                        }
                        XSSFSheet slSheet2 = xwb.getSheet("单位工程");
                        for (int i = 0; i < slSheet2.getPhysicalNumberOfRows()-26; i++){
                            slSheet2.getRow(i).getCell(1).setCellType(CellType.STRING);

                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (value.equals("交通安全设施")){
                                slSheet2.getRow(i + 25).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 25).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 25).getCell(1).getStringCellValue();
                                String value2 = slSheet2.getRow(i + 25).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","交通安全设施");
                                map.put("dwgcdf",value1);
                                map.put("dwgczldj",value2);
                                resultList1.add(map);
                            }
                        }
                    }
                }
            }
        }
        if (resultList1 != null){
            for (Map<String, Object> objectMap : resultList1) {
                String htd = objectMap.get("htd").toString();
                String dwgc = objectMap.get("dwgc").toString();
                String dwgcdf = objectMap.get("dwgcdf").toString();
                String dwgczldj = objectMap.get("dwgczldj").toString();
                if (resultList != null){
                    for (Map<String, Object> stringObjectMap : resultList) {
                        String htdname = stringObjectMap.get("htd").toString();
                        String dwgcname = stringObjectMap.get("dwgc").toString();
                        if (htd.equals(htdname) && dwgc.equals(dwgcname)){
                            stringObjectMap.put("dwgcdf",dwgcdf);
                            stringObjectMap.put("dwgczldj",dwgczldj);
                        }

                    }
                }
            }
        }

        //按合同段排序
        Collections.sort(resultList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("htd").toString();
                String name2 = o2.get("htd").toString();
                return name1.compareTo(name2);
            }
        });
        return resultList;
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtosdpdhzb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getsdpdbData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "单位工程", "", "", "分部工程", "", "", "", "", ""};
            String[] header2 = {"", "名称", "得分", "质量等级", "名称", "权值", "得分", "实测得分", "外观扣分", "质量等级"};
            String s = "${表11.1.4-1}";
            TextSelection textSelection = findStringInPages(xw,s,20,45);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 10);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 1, 3);
                    }
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 9);
                    }
                }
                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        table.applyVerticalMerge(0, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                break;
                            case 1:
                                p.appendText(rowData.get("dwgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.get("dwgcdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.get("dwgczldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.get("fbgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("qz").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText(rowData.get("df").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText(rowData.get("scdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText(rowData.get("wgkf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 9:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }

                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }
    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getsdpdbData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        QueryWrapper<JjgJgHtdinfo> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("隧道工程")){
                    File ff = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    if (!ff.exists()) {
                        break;
                    }else {
                        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                        for (int j = 0; j < xwb.getNumberOfSheets(); j++) {
                            String sheetName = xwb.getSheetName(j);
                            if (sheetName.contains("隧道")){
                                XSSFSheet slSheet1 = xwb.getSheet(sheetName);
                                for (int i = 0; i < slSheet1.getPhysicalNumberOfRows()-20; i++){
                                    slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                                    slSheet1.getRow(i).getCell(15).setCellType(CellType.STRING);//项目名称
                                    slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);//分部工程名称

                                    String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                                    String bzxmmc = slSheet1.getRow(i).getCell(15).getStringCellValue();

                                    if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                                        slSheet1.getRow(i+20).getCell(6).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(10).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(17).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                        String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                        if (!fbgcname.equals("隧道路面")){
                                            Map<String,Object> map = new HashMap<>();
                                            map.put("htd",htd);
                                            map.put("dwgc",sheetName.substring(3));
                                            map.put("fbgc",fbgcname);
                                            map.put("scdf",slSheet1.getRow(i+20).getCell(6).getStringCellValue());
                                            map.put("wgkf",slSheet1.getRow(i+20).getCell(10).getStringCellValue());
                                            map.put("df",slSheet1.getRow(i+20).getCell(17).getStringCellValue());
                                            map.put("zldj",slSheet1.getRow(i+20).getCell(20).getStringCellValue());
                                            if (fbgcname.equals("衬砌")){
                                                map.put("qz",3);
                                            }else if (fbgcname.equals("总体")){
                                                map.put("qz",1);
                                            }else {
                                                map.put("qz",0);
                                            }
                                            resultList.add(map);
                                        }
                                    }
                                }
                            }
                        }
                        XSSFSheet slSheet2 = xwb.getSheet("单位工程");
                        for (int i = 0; i < slSheet2.getPhysicalNumberOfRows()-26; i++){
                            slSheet2.getRow(i).getCell(1).setCellType(CellType.STRING);

                            String dwvalue = slSheet2.getRow(i).getCell(0).getStringCellValue();
                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (dwvalue.equals("单位工程名称：") && value.contains("隧道")){
                                slSheet2.getRow(i + 25).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 25).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 25).getCell(1).getStringCellValue();
                                String value2 = slSheet2.getRow(i + 25).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc",value);
                                map.put("dwgcdf",value1);
                                map.put("dwgczldj",value2);
                                resultList1.add(map);
                            }
                        }
                    }
                }
            }
        }
        if (resultList1 != null){
            for (Map<String, Object> objectMap : resultList1) {
                String htd = objectMap.get("htd").toString();
                String dwgc = objectMap.get("dwgc").toString();
                String dwgcdf = objectMap.get("dwgcdf").toString();
                String dwgczldj = objectMap.get("dwgczldj").toString();
                if (resultList != null){
                    for (Map<String, Object> stringObjectMap : resultList) {
                        String htdname = stringObjectMap.get("htd").toString();
                        String dwgcname = stringObjectMap.get("dwgc").toString();
                        if (htd.equals(htdname) && dwgc.equals(dwgcname)){
                            stringObjectMap.put("dwgcdf",dwgcdf);
                            stringObjectMap.put("dwgczldj",dwgczldj);
                        }

                    }
                }
            }
        }

        //按合同段排序
        Collections.sort(resultList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("htd").toString();
                String name2 = o2.get("htd").toString();
                return name1.compareTo(name2);
            }
        });
        return resultList;
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoqlpdhzb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getqlpdbData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "单位工程", "", "", "分部工程", "", "", "", "", ""};
            String[] header2 = {"", "名称", "得分", "质量等级", "名称", "权值", "得分", "实测得分", "外观扣分", "质量等级"};
            String s = "${表11.1.3-1}";
            TextSelection textSelection = findStringInPages(xw,s,20,45);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 10);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 1, 3);
                    }
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 9);
                    }
                }
                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        table.applyVerticalMerge(0, 0, 1);
                    }
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                break;
                            case 1:
                                p.appendText(rowData.get("dwgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.get("dwgcdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.get("dwgczldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.get("fbgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("qz").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText(rowData.get("df").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText(rowData.get("scdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText(rowData.get("wgkf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 9:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }

                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }
        }



    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getqlpdbData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        QueryWrapper<JjgJgHtdinfo> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("桥梁工程")){
                    File ff = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    if (!ff.exists()) {
                        break;
                    }else {
                        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                        for (int j = 0; j < xwb.getNumberOfSheets(); j++) {
                            String sheetName = xwb.getSheetName(j);
                            if (sheetName.contains("桥")){
                                XSSFSheet slSheet1 = xwb.getSheet(sheetName);
                                for (int i = 0; i < slSheet1.getPhysicalNumberOfRows()-20; i++){
                                    slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                                    slSheet1.getRow(i).getCell(15).setCellType(CellType.STRING);//项目名称
                                    slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);//分部工程名称

                                    String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                                    String bzxmmc = slSheet1.getRow(i).getCell(15).getStringCellValue();

                                    if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                                        slSheet1.getRow(i+20).getCell(6).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(10).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(17).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                        String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                        if (!fbgcname.equals("桥面系")){
                                            Map<String,Object> map = new HashMap<>();
                                            map.put("htd",htd);
                                            map.put("dwgc",sheetName.substring(3));
                                            map.put("fbgc",fbgcname);
                                            map.put("scdf",slSheet1.getRow(i+20).getCell(6).getStringCellValue());
                                            map.put("wgkf",slSheet1.getRow(i+20).getCell(10).getStringCellValue());
                                            map.put("df",slSheet1.getRow(i+20).getCell(17).getStringCellValue());
                                            map.put("zldj",slSheet1.getRow(i+20).getCell(20).getStringCellValue());
                                            if (fbgcname.equals("桥梁上部")){
                                                map.put("qz",3);
                                            }else if (fbgcname.equals("桥梁下部")){
                                                map.put("qz",2);
                                            }else {
                                                map.put("qz",0);
                                            }
                                            resultList.add(map);
                                        }
                                    }
                                }
                            }
                        }
                        XSSFSheet slSheet2 = xwb.getSheet("单位工程");
                        for (int i = 0; i < slSheet2.getPhysicalNumberOfRows()-26; i++){
                            slSheet2.getRow(i).getCell(1).setCellType(CellType.STRING);

                            String dwvalue = slSheet2.getRow(i).getCell(0).getStringCellValue();
                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (dwvalue.equals("单位工程名称：") && value.contains("桥")){
                                slSheet2.getRow(i + 25).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 25).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 25).getCell(1).getStringCellValue();
                                String value2 = slSheet2.getRow(i + 25).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc",value);
                                map.put("dwgcdf",value1);
                                map.put("dwgczldj",value2);
                                resultList1.add(map);
                            }
                        }
                    }
                }
            }
        }

        if (resultList1 != null){
            for (Map<String, Object> objectMap : resultList1) {
                String htd = objectMap.get("htd").toString();
                String dwgc = objectMap.get("dwgc").toString();
                String dwgcdf = objectMap.get("dwgcdf").toString();
                String dwgczldj = objectMap.get("dwgczldj").toString();
                if (resultList != null){
                    for (Map<String, Object> stringObjectMap : resultList) {
                        String htdname = stringObjectMap.get("htd").toString();
                        String dwgcname = stringObjectMap.get("dwgc").toString();
                        if (htd.equals(htdname) && dwgc.equals(dwgcname)){
                            stringObjectMap.put("dwgcdf",dwgcdf);
                            stringObjectMap.put("dwgczldj",dwgczldj);
                        }

                    }
                }
            }
        }

        //按合同段排序
        Collections.sort(resultList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("htd").toString();
                String name2 = o2.get("htd").toString();
                return name1.compareTo(name2);
            }
        });
        return resultList;


    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtolmpdhzb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlmpdbData(proname);
        if (data !=null && data.size()>0){
            String[] header1 = {"合同段", "单位工程","","","分部工程","","","","",""};
            String[] header2 = {"","名称", "得分","质量等级","名称","权值","得分","实测得分","外观扣分","质量等级"};
            String s = "${表11.1.2-1}";
            TextSelection textSelection = findStringInPages(xw,s,20,45);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 10);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 1, 3);
                    }
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 9);
                    }
                }


                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        table.applyVerticalMerge(0, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                break;
                            case 1:
                                p.appendText(rowData.get("dwgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.get("dwgcdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.get("dwgczldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.get("fbgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("qz").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText(rowData.get("df").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText(rowData.get("scdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText(rowData.get("wgkf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 9:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }

                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }

        }



    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getlmpdbData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        QueryWrapper<JjgJgHtdinfo> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("路面工程")){
                    File ff = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    if (!ff.exists()) {
                        break;
                    }else {
                        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                        XSSFSheet slSheet1 = xwb.getSheet("分部-路面");

                        for (int i = 0; i < slSheet1.getPhysicalNumberOfRows()-20; i++){
                            slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                            slSheet1.getRow(i).getCell(15).setCellType(CellType.STRING);//项目名称
                            slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);//分部工程名称

                            String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                            String bzxmmc = slSheet1.getRow(i).getCell(15).getStringCellValue();

                            if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                                slSheet1.getRow(i+20).getCell(6).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(10).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(17).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路面工程");
                                map.put("fbgc",fbgcname);
                                map.put("scdf",slSheet1.getRow(i+20).getCell(6).getStringCellValue());
                                map.put("wgkf",slSheet1.getRow(i+20).getCell(10).getStringCellValue());
                                map.put("df",slSheet1.getRow(i+20).getCell(17).getStringCellValue());
                                map.put("zldj",slSheet1.getRow(i+20).getCell(20).getStringCellValue());
                                map.put("qz",1);
                                resultList.add(map);
                            }
                        }
                        XSSFSheet slSheet2 = xwb.getSheet("单位工程");
                        for (int i = 0; i < slSheet2.getPhysicalNumberOfRows()-26; i++){
                            slSheet2.getRow(i).getCell(1).setCellType(CellType.STRING);

                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (value.equals("路面工程")){
                                slSheet2.getRow(i + 25).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 25).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 25).getCell(1).getStringCellValue();
                                String value2 = slSheet2.getRow(i + 25).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路面工程");
                                map.put("dwgcdf",value1);
                                map.put("dwgczldj",value2);
                                resultList1.add(map);
                            }
                        }
                    }

                }
            }
        }
        if (resultList1 != null){
            for (Map<String, Object> objectMap : resultList1) {
                String htd = objectMap.get("htd").toString();
                String dwgc = objectMap.get("dwgc").toString();
                String dwgcdf = objectMap.get("dwgcdf").toString();
                String dwgczldj = objectMap.get("dwgczldj").toString();
                if (resultList != null){
                    for (Map<String, Object> stringObjectMap : resultList) {
                        String htdname = stringObjectMap.get("htd").toString();
                        String dwgcname = stringObjectMap.get("dwgc").toString();
                        if (htd.equals(htdname) && dwgc.equals(dwgcname)){
                            stringObjectMap.put("dwgcdf",dwgcdf);
                            stringObjectMap.put("dwgczldj",dwgczldj);
                        }

                    }
                }
            }
        }

        //按合同段排序
        Collections.sort(resultList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("htd").toString();
                String name2 = o2.get("htd").toString();
                return name1.compareTo(name2);
            }
        });
        return resultList;

    }


    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoljpdhzb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getljpdbData(proname);
        if (data !=null && data.size()>0){
            String[] header1 = {"合同段", "单位工程","","","分部工程","","","","",""};
            String[] header2 = {"","名称", "得分","质量等级","名称","权值","得分","实测得分","外观扣分","质量等级"};
            String s = "${表11.1.1-1}";
            TextSelection textSelection = findStringInPages(xw,s,20,45);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 10);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 1, 3);
                    }
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 9);
                    }
                }


                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header2.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header2[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        table.applyVerticalMerge(0, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header1.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                if ((rowIdx ) % 5 == 0) {
                                    p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 1:
                                if ((rowIdx ) % 5 == 0) {
                                    p.appendText(rowData.get("dwgc").toString()).getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 2:
                                if ((rowIdx ) % 5 == 0) {
                                    p.appendText(rowData.get("dwgcdf").toString()).getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 3:
                                if ((rowIdx ) % 5 == 0) {
                                    p.appendText(rowData.get("dwgczldj").toString()).getCharacterFormat().setFontSize(9f);
                                }
                                break;
                            case 4:
                                p.appendText(rowData.get("fbgc").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("qz").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
                                p.appendText(rowData.get("df").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 7:
                                p.appendText(rowData.get("scdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 8:
                                p.appendText(rowData.get("wgkf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 9:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            default:
                                break;
                        }
                    }
                }
                /*int count = table.getRows().getCount();
                for (int i = 2; i <= count-4; i+=4) {
                    table.applyVerticalMerge(0, i, i+4);
                    table.applyVerticalMerge(1, i, i+4);
                    table.applyVerticalMerge(2, i, i+4);
                    table.applyVerticalMerge(3, i, i+4);
                }*/
                int count = table.getRows().getCount();
                int startRow = 2; // 开始行（跳过表头或其他不想合并的行）
                while (startRow + 4 < count) { // 确保有足够的行进行合并
                    table.applyVerticalMerge(0, startRow, startRow + 4); // 仅合并第一列
                    table.applyVerticalMerge(1, startRow, startRow + 4); // 仅合并第一列
                    table.applyVerticalMerge(2, startRow, startRow + 4); // 仅合并第一列
                    table.applyVerticalMerge(3, startRow, startRow + 4); // 仅合并第一列
                    startRow += 5; // 移动到下一个要合并的5行组的起始行
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);


            }
        }


    }

    private List<Map<String,Object>> getljpdbData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        //先查合同段
        QueryWrapper<JjgJgHtdinfo> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(htdinfo);
        if (list != null && list.size() > 0){
            //List<String> htdlist = new ArrayList<>();
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {

                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("路基工程")){
                    File ff = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    if (!ff.exists()) {
                        break;
                    }else {
                        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                        XSSFSheet slSheet1 = xwb.getSheet("分部-路基");
                        for (int i = 0; i < slSheet1.getPhysicalNumberOfRows()-20; i++){

                            slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                            slSheet1.getRow(i).getCell(15).setCellType(CellType.STRING);//项目名称
                            slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);//分部工程名称

                            String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                            String bzxmmc = slSheet1.getRow(i).getCell(15).getStringCellValue();


                            if (bzhtd.equals(htd) && bzxmmc.equals(proname)){

                                slSheet1.getRow(i+20).getCell(6).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(10).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(17).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路基工程");
                                map.put("fbgc",fbgcname);
                                map.put("scdf",slSheet1.getRow(i+20).getCell(6).getStringCellValue());
                                map.put("wgkf",slSheet1.getRow(i+20).getCell(10).getStringCellValue());
                                map.put("df",slSheet1.getRow(i+20).getCell(17).getStringCellValue());
                                map.put("zldj",slSheet1.getRow(i+20).getCell(20).getStringCellValue());
                                if (fbgcname.equals("路基土石方")){
                                    map.put("qz",3);
                                }else if (fbgcname.equals("排水工程")){
                                    map.put("qz",1);
                                }else if (fbgcname.equals("小桥")){
                                    map.put("qz",2);
                                }else if (fbgcname.equals("涵洞")){
                                    map.put("qz",1);
                                }else if (fbgcname.equals("支挡工程")){
                                    map.put("qz",2);
                                }else {
                                    map.put("qz",0);
                                }
                                resultList.add(map);

                            }
                        }
                        XSSFSheet slSheet2 = xwb.getSheet("单位工程");
                        for (int i = 0; i < slSheet2.getPhysicalNumberOfRows()-26; i++){
                            slSheet2.getRow(i).getCell(1).setCellType(CellType.STRING);

                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (value.equals("路基工程")){
                                slSheet2.getRow(i + 25).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 25).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 25).getCell(1).getStringCellValue();
                                String value2 = slSheet2.getRow(i + 25).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路基工程");
                                map.put("dwgcdf",value1);
                                map.put("dwgczldj",value2);
                                resultList1.add(map);
                            }
                        }
                    }
                }
            }
        }
        if (resultList1 != null){
            for (Map<String, Object> objectMap : resultList1) {
                String htd = objectMap.get("htd").toString();
                String dwgc = objectMap.get("dwgc").toString();
                String dwgcdf = objectMap.get("dwgcdf").toString();
                String dwgczldj = objectMap.get("dwgczldj").toString();
                if (resultList != null){
                    for (Map<String, Object> stringObjectMap : resultList) {
                        String htdname = stringObjectMap.get("htd").toString();
                        String dwgcname = stringObjectMap.get("dwgc").toString();
                        if (htd.equals(htdname) && dwgc.equals(dwgcname)){
                            stringObjectMap.put("dwgcdf",dwgcdf);
                            stringObjectMap.put("dwgczldj",dwgczldj);
                        }

                    }
                }
            }
        }


        //按合同段排序
        Collections.sort(resultList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("htd").toString();
                String name2 = o2.get("htd").toString();
                return name1.compareTo(name2);
            }
        });
        return resultList;
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtojafhl(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getjafhldata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "横梁中心高度","",""};
            String[] header1 = {"","检测点数", "合格点数","合格率(％)"};
            String s = "${表4.1.3-2}";
            TextSelection textSelection = findStringInPages(xw,s,20,40);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 4);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 1, 3);
                    }
                }

                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0) {
                        table.applyVerticalMerge(0, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("jabxfhljcds").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("jabxfhlhgds").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("jabxfhlhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }

                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }

        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getjafhldata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        CommonInfoVo commonInfoVo = new CommonInfoVo();
        commonInfoVo.setProname(proname);
        List<Map<String, Object>> jafhldata = jjgFbgcJtaqssJabxfhlJgfcService.lookJdbjg(commonInfoVo);
        if (jafhldata != null && jafhldata.size()>0){
            Map<Object, List<Map<String, Object>>> grouped = jafhldata.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double jabxfhlzds = 0,jabxfhlhgds = 0;
                for (Map<String, Object> map : value) {
                    String jcxm = map.get("检测项目").toString();
                    if (jcxm.contains("波形梁钢护栏横梁中心高度")){
                        jabxfhlzds += Double.valueOf(map.get("总点数").toString());
                        jabxfhlhgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }
                Map map = new HashMap<>();
                map.put("htd",key);
                map.put("jabxfhljcds",decf.format(jabxfhlzds));
                map.put("jabxfhlhgds",decf.format(jabxfhlhgds));
                map.put("jabxfhlhgl",jabxfhlzds != 0 ? df.format(jabxfhlhgds / jabxfhlzds * 100) : 0);

                resultlist.add(map);
            });
        }
        return resultlist;
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtojabx(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getjabxdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "反光标线逆反射系数","","","标线厚度","",""};
            String[] header1 = {"","检测点数", "合格点数","合格率(％)","检测点数", "合格点数","合格率(％)"};
            String s = "${表4.1.3-1}";
            TextSelection textSelection = findStringInPages(xw,s,20,40);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 2, 7);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(35);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 1){
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0,1,3);
                    }
                    if (i == 4){
                        table.applyHorizontalMerge(0,4,6);
                    }
                }

                TableRow row1 = table.getRows().get(1);
                row1.isHeader(true);
                row1.setHeight(35);
                row1.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header1.length; i++) {
                    row1.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header1[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                    if (i == 0){
                        table.applyVerticalMerge(0,0,1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("jabxxsjcds").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("jabxxshgds").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("jabxxshgl").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("jabxhdjcds").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("jabxhdhgds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("jabxhdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }

                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }

        }
    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getjabxdata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        CommonInfoVo commonInfoVo = new CommonInfoVo();
        commonInfoVo.setProname(proname);
        List<Map<String, Object>> jabxdata = jjgFbgcJtaqssJabxJgfcService.lookJdbjg(commonInfoVo);
        if (jabxdata != null && jabxdata.size()>0){
            Map<Object, List<Map<String, Object>>> grouped = jabxdata.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));

            grouped.forEach((key, value) -> {
                double jabxhdzds = 0,jabxhdhgds = 0;
                double jabxxszds = 0,jabxxshgds = 0;

                for (Map<String, Object> map : value) {
                    String jcxm = map.get("检测项目").toString();
                    if (jcxm.contains("交安标线厚度")){
                        jabxhdzds += Double.valueOf(map.get("总点数").toString());
                        jabxhdhgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    if (jcxm.contains("交安标线逆反射系数")){
                        jabxxszds += Double.valueOf(map.get("总点数").toString());
                        jabxxshgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }
                Map map = new HashMap<>();
                map.put("htd",key);
                map.put("jabxxsjcds",decf.format(jabxxszds));
                map.put("jabxxshgds",decf.format(jabxxshgds));
                map.put("jabxxshgl",jabxxszds != 0 ? df.format(jabxxshgds / jabxxszds * 100) : 0);

                map.put("jabxhdjcds",decf.format(jabxhdzds));
                map.put("jabxhdhgds",decf.format(jabxhdhgds));
                map.put("jabxhdhgl",jabxhdzds != 0 ? df.format(jabxhdhgds / jabxhdzds * 100) : 0);
                resultlist.add(map);

            });
        }
        return resultlist;
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtofcdb(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getDocxdbData(proname);
        if (data != null && data.size()>0){
            for (Map<String, Object> map : data) {
                for (Map.Entry<String, Object> entry : map.entrySet()) {
                    String s = "${"+entry.getKey()+"}";
                    String value = entry.getValue().toString();
                    xw.replace(s, value, false, true);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                }
            }

        }

    }


    /**
     *
     * @param proname
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getDocxdbData(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //弯沉
        QueryWrapper<JjgJgjcsj> wrapper = new QueryWrapper<>();
        wrapper.eq("proname", proname);
        List<JjgJgjcsj> jjgJgjcsjs = jjgJgjcsjMapper.selectList(wrapper);

        List<Map<String, Object>> data = new ArrayList<>();
        Map wcmap = new HashMap<>();
        Map czmap = new HashMap<>();
        Map sdczmap = new HashMap<>();
        /*Map pzdhntljxmap = new HashMap<>();
        Map pzdhntzdmap = new HashMap<>();
        Map pzdlqljxhtmap = new HashMap<>();
        Map pzdlqqmzxmap = new HashMap<>();*/
        Map pzdlqzxmap = new HashMap<>();
        Map pzdsdlqallmap = new HashMap<>();
        Map pzdlqqmap = new HashMap<>();
        Map pzdlqqhtmap = new HashMap<>();
        Map pzdhntqljxmap = new HashMap<>();
        Map pzdlqhtljxmap = new HashMap<>();
        Map pzdhnthtmap = new HashMap<>();

        Map mcxszxmap = new HashMap<>();
        Map mcxssdmap = new HashMap<>();
        Map mcxsqmmap = new HashMap<>();

        Map gzsdzxmap = new HashMap<>();
        Map gzsdsdmap = new HashMap<>();
        Map gzsdqmmap = new HashMap<>();
        Map tlmxlbgcmap = new HashMap<>();

        if (jjgJgjcsjs != null && jjgJgjcsjs.size()>0){
            double wczds = 0,wchgds = 0;
            double czzds = 0,czhgds = 0;
            double sdczzds = 0,sdczhgds = 0;
            /*double pzdhntljxzds = 0,pzdhntljxhgds = 0;
            double pzdhnthtzds = 0,pzdhnththgds = 0;
            double pzdljxhtzds = 0,pzdljxhthgds = 0;
            double pzdlqqmzxzds = 0,pzdlqqmzxhgds = 0;*/
            double pzdlqzxzds = 0,pzdlqzxhgds = 0;
            double pzdsdlqallzds = 0,pzdsdlqallhgds = 0;
            double pzdlqqzds = 0,pzdlqqhgds = 0;
            double pzdlqqhtzds = 0,pzdlqqhthgds = 0;
            double pzdhntqljxzds = 0,pzdhntqljxhgds = 0;

            double mcxszxzds = 0,mcxszxhgds = 0;
            double mcxssdzds = 0,mcxssdhgds = 0;
            double mcxsqmzds = 0,mcxsqmhgds = 0;

            double gzsdzxzds = 0,gzsdzxhgds = 0;
            double gzsdsdzds = 0,gzsdsdhgds = 0;
            double gzsdqmzds = 0,gzsdqmhgds = 0;

            double tlmxlbgczds = 0,tlmxlbgchgds = 0;

            for (JjgJgjcsj jjgJgjcsj : jjgJgjcsjs) {
                String dwgc = jjgJgjcsj.getDwgc();
                String fbgc = jjgJgjcsj.getFbgc();
                String ccxm = jjgJgjcsj.getCcxm();
                if ("路面工程".equals(dwgc) && "路面面层".equals(fbgc) && "△沥青路面弯沉".equals(ccxm)){
                    wczds += Double.valueOf(jjgJgjcsj.getZds());
                    wchgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                if ("路面工程".equals(dwgc) && "路面面层".equals(fbgc) && ccxm.contains("沥青路面车辙（路面面层）")){
                    czzds += Double.valueOf(jjgJgjcsj.getZds());
                    czhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                if ("路面工程".equals(dwgc) && "路面面层".equals(fbgc) && ccxm.contains("沥青路面车辙（隧道路面）")){
                    sdczzds += Double.valueOf(jjgJgjcsj.getZds());
                    sdczhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }


                /*if (dwgc.contains("连接线") && ccxm.contains("平整度") && ccxm.contains("混凝土")){
                    pzdhntljxzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdhntljxhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                if (dwgc.contains("匝道") && ccxm.contains("平整度") && ccxm.contains("混凝土")){
                    pzdhnthtzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdhnththgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                if (dwgc.contains("连接线") || dwgc.contains("匝道")  && ccxm.contains("沥青路面") && ccxm.contains("混凝土")){
                    pzdljxhtzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdljxhthgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                if (dwgc.contains("桥") && ccxm.contains("平整度") && ccxm.contains("混凝土")){
                    pzdlqqmzxzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdlqqmzxhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }*/

                //平整度 混凝土 连接线

                //平整度 混凝土 互通

                //平整度 沥青 互通及连接线


                //平整度 沥青路面  路面主线
                if (dwgc.contains("路面工程") && fbgc.contains("路面面层") && ccxm.contains("平整度（沥青路面面层）")){
                    pzdlqzxzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdlqzxhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                //平整度 沥青桥 主线
                if (dwgc.contains("路面工程") && fbgc.contains("路面面层") && ccxm.contains("平整度（桥面系沥青）") ){
                    pzdlqqzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdlqqhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                //平整度 沥青桥 互通及连接线
                if (ccxm.contains("沥青桥平整度") ){
                    pzdlqqhtzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdlqqhthgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                //平整度 混凝土桥 连接线
                if (dwgc.contains("连接线") && ccxm.contains("混凝土桥平整度")){
                    pzdhntqljxzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdhntqljxhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                //隧道路面平整度 沥青
                if (ccxm.contains("平整度（隧道路面沥青）")){
                    pzdsdlqallzds += Double.valueOf(jjgJgjcsj.getZds());
                    pzdsdlqallhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                
                //摩擦系数
                if (ccxm.contains("摩擦系数（沥青路面面层）")){
                    mcxszxzds += Double.valueOf(jjgJgjcsj.getZds());
                    mcxszxhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                if (ccxm.contains("摩擦系数（隧道路面）")){
                    mcxssdzds += Double.valueOf(jjgJgjcsj.getZds());
                    mcxssdhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                if (ccxm.contains("摩擦系数（桥面系）") || fbgc.contains("桥面系") && ccxm.contains("摩擦系数")){
                    mcxsqmzds += Double.valueOf(jjgJgjcsj.getZds());
                    mcxsqmhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                //构造深度
                if (ccxm.contains("构造深度（沥青路面面层）")){
                    gzsdzxzds += Double.valueOf(jjgJgjcsj.getZds());
                    gzsdzxhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                if (ccxm.contains("沥青桥面构造深度") || ccxm.contains("构造深度（桥面系沥青）")){
                    gzsdqmzds += Double.valueOf(jjgJgjcsj.getZds());
                    gzsdqmhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
                if (ccxm.contains("构造深度（隧道路面沥青）") || ccxm.contains("构造深度（隧道路面）")){
                    gzsdsdzds += Double.valueOf(jjgJgjcsj.getZds());
                    gzsdsdhgds += Double.valueOf(jjgJgjcsj.getHgds());
                }

                if (ccxm.contains("混凝土路面相邻板高差")){
                    tlmxlbgczds += Double.valueOf(jjgJgjcsj.getZds());
                    tlmxlbgchgds += Double.valueOf(jjgJgjcsj.getHgds());
                }
            }
            wcmap.put("wchgl",wczds != 0 ? df.format(wchgds / wczds * 100) : 0);
            czmap.put("czhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
            sdczmap.put("sdczhgl",sdczzds != 0 ? df.format(sdczhgds / sdczzds * 100) : 0);

            //pzdhntljxmap.put("pzdhntljx",pzdhntljxzds != 0 ? df.format(pzdhntljxhgds / pzdhntljxzds * 100) : 0);
            //pzdhntzdmap.put("pzdhntzd",pzdhnthtzds != 0 ? df.format(pzdhnththgds / pzdhnthtzds * 100) : 0);
            //pzdlqljxhtmap.put("pzdlqljxht",pzdljxhtzds != 0 ? df.format(pzdljxhthgds / pzdljxhtzds * 100) : 0);

            pzdlqzxmap.put("pzdlqzxhgl",pzdlqzxzds != 0 ? df.format(pzdlqzxhgds / pzdlqzxzds * 100) : 0);
            pzdsdlqallmap.put("pzdsdlqallhgl",pzdsdlqallzds != 0 ? df.format(pzdsdlqallhgds / pzdsdlqallzds * 100) : 0);
            pzdlqqmap.put("pzdlqqhgl",pzdlqqzds != 0 ? df.format(pzdlqqhgds / pzdlqqzds * 100) : 0);
            pzdlqqhtmap.put("pzdlqqhthgl",pzdlqqhtzds != 0 ? df.format(pzdlqqhthgds / pzdlqqhtzds * 100) : 0);
            pzdhntqljxmap.put("pzdhntqljxhgl",pzdhntqljxzds != 0 ? df.format(pzdhntqljxhgds / pzdhntqljxzds * 100) : 0);

            mcxszxmap.put("mcxszxhgl",mcxszxzds != 0 ? df.format(mcxszxhgds / mcxszxzds * 100) : 0);
            mcxssdmap.put("mcxssdhgl",mcxssdzds != 0 ? df.format(mcxssdhgds / mcxssdzds * 100) : 0);
            mcxsqmmap.put("mcxsqmhgl",mcxsqmzds != 0 ? df.format(mcxsqmhgds / mcxsqmzds * 100) : 0);

            gzsdzxmap.put("gzsdzxhgl",gzsdzxzds != 0 ? df.format(gzsdzxhgds / gzsdzxzds * 100) : 0);
            gzsdqmmap.put("gzsdqmhgl",gzsdqmzds != 0 ? df.format(gzsdqmhgds / gzsdqmzds * 100) : 0);
            gzsdsdmap.put("gzsdsdhgl",gzsdsdzds != 0 ? df.format(gzsdsdhgds / gzsdsdzds * 100) : 0);

            tlmxlbgcmap.put("tlmxlbgchgl",tlmxlbgczds != 0 ? df.format(tlmxlbgchgds / tlmxlbgczds * 100) : 0);

        }
        CommonInfoVo commonInfoVo = new CommonInfoVo();
        commonInfoVo.setProname(proname);
        List<Map<String, Object>> maps1 = jjgFbgcLmgcLmwcJgfcService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> maps2 = jjgFbgcLmgcLmwcLcfJgfcService.lookJdbjg(commonInfoVo);
        double wczds = 0,wxhgds = 0;
        List<String> wcpjzlist = new ArrayList<>();
        List<String> wcsjzlist = new ArrayList<>();
        List<String> wcdbzlist = new ArrayList<>();
        if (maps1 != null && maps1.size() > 0){
            for (Map<String, Object> stringObjectMap : maps1) {
                wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                wxhgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
                String pjz = stringObjectMap.get("平均值").toString();
                String dbz = stringObjectMap.get("代表值").toString();
                String gdz = stringObjectMap.get("规定值").toString();
                wcpjzlist.add(pjz.substring(1,pjz.length()-1));
                wcdbzlist.add(dbz.substring(1,dbz.length()-1));
                wcsjzlist.add(gdz.substring(1,gdz.length()-1));
            }
        }
        if (maps2 != null && maps2.size() > 0){
            for (Map<String, Object> stringObjectMap : maps2) {
                wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                wxhgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
                String pjz = stringObjectMap.get("平均值").toString();
                String dbz = stringObjectMap.get("代表值").toString();
                String gdz = stringObjectMap.get("规定值").toString();
                wcpjzlist.add(pjz.substring(1,pjz.length()-1));
                wcdbzlist.add(dbz.substring(1,dbz.length()-1));
                wcsjzlist.add(gdz.substring(1,gdz.length()-1));
            }
        }
        wcmap.put("wcgdz",getpjz(wcsjzlist));
        wcmap.put("wcpjzjg",getpjz(wcpjzlist));
        wcmap.put("wcsczbhfwjg",getbhfw(wcdbzlist));
        wcmap.put("wchgljg",wczds != 0 ? df.format(wxhgds / wczds * 100) : 0);
        data.add(wcmap);

        //车辙
        List<Map<String, Object>> czdata = jjgZdhCzJgfcService.lookJdbjg(proname);
        if (czdata != null && czdata.size() > 0){
            List<String>  pjzlist = new ArrayList<>();
            List<String>  sdpjzlist = new ArrayList<>();
            List<String>  maxlist = new ArrayList<>();
            List<String>  sdmaxlist = new ArrayList<>();
            List<String>  minlist = new ArrayList<>();
            List<String>  sdminlist = new ArrayList<>();
            double czzds=0,czhgds=0;
            double czsdzds=0,czsdhgds=0;
            for (Map<String, Object> czdatum : czdata) {
                String lmlx = czdatum.get("路面类型").toString();
                if (lmlx.contains("路面") || lmlx.contains("桥")){
                    String gdz = czdatum.get("设计值").toString();
                    czmap.put("czgdz",gdz);
                    String pjz = czdatum.get("pjz").toString();
                    pjzlist.add(pjz);
                    String Max = czdatum.get("Max").toString();
                    maxlist.add(Max);
                    String Min = czdatum.get("Min").toString();
                    minlist.add(Min);
                    czzds += Double.valueOf(czdatum.get("总点数").toString());
                    czhgds += Double.valueOf(czdatum.get("合格点数").toString());
                }
                if (lmlx.contains("隧道")){
                    String gdz = czdatum.get("设计值").toString();
                    sdczmap.put("sdczgdz",gdz);
                    String pjz = czdatum.get("pjz").toString();
                    sdpjzlist.add(pjz);

                    String Max = czdatum.get("Max").toString();
                    sdmaxlist.add(Max);
                    String Min = czdatum.get("Min").toString();
                    sdminlist.add(Min);

                    czsdzds += Double.valueOf(czdatum.get("总点数").toString());
                    czsdhgds += Double.valueOf(czdatum.get("合格点数").toString());
                }
            }
            czmap.put("czpjz",getpjz(pjzlist));
            czmap.put("czsczbhfw",getczbhfw(maxlist,minlist));
            czmap.put("czjghgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
            data.add(czmap);

            sdczmap.put("sdczpjz",getpjz(sdpjzlist));
            sdczmap.put("sdczsczbhfw",getczbhfw(sdmaxlist,sdminlist));
            sdczmap.put("sdczjghgl",czsdzds != 0 ? df.format(czsdhgds / czsdzds * 100) : 0);
            data.add(sdczmap);
        }

        //平整度
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        if (pzdlist != null && pzdlist.size() > 0){
            List<String>  pzdsdlqallpjzlist = new ArrayList<>();
            List<String>  pzdsdlqallmaxlist = new ArrayList<>();
            List<String>  pzdsdlqallminlist = new ArrayList<>();
            double pzdsdlqallzds = 0,pzdsdlqallhgds = 0;
            List<String>  pzdlqqpjzlist = new ArrayList<>();
            List<String>  pzdlqqmaxlist = new ArrayList<>();
            List<String>  pzdlqqminlist = new ArrayList<>();
            double pzdlqqzds = 0,pzdlqqhgds = 0;

            List<String>  pzdlqqhtpjzlist = new ArrayList<>();
            List<String>  pzdlqqhtmaxlist = new ArrayList<>();
            List<String>  pzdlqqhtminlist = new ArrayList<>();
            double pzdlqqhtzds = 0,pzdlqqhthgds = 0;

            List<String>  pzdhntqljxpjzlist = new ArrayList<>();
            List<String>  pzdhntqljxmaxlist = new ArrayList<>();
            List<String>  pzdhntqljxminlist = new ArrayList<>();
            double pzdhntqljxzds = 0,pzdhntqljxhgds = 0;

            List<String>  pzdlqzxpjzlist = new ArrayList<>();
            List<String>  pzdlqzxmaxlist = new ArrayList<>();
            List<String>  pzdlqzxminlist = new ArrayList<>();
            double pzdlqzxzds = 0,pzdlqzxhgds = 0;

            List<String>  pzdlqhtljxpjzlist = new ArrayList<>();
            List<String>  pzdlqhtljxmaxlist = new ArrayList<>();
            List<String>  pzdlqhtljxminlist = new ArrayList<>();
            double pzdlqhtljxzds = 0,pzdlqhtljxhgds = 0;

            List<String>  pzdhntljxpjzlist = new ArrayList<>();
            List<String>  pzdhntljxmaxlist = new ArrayList<>();
            List<String>  pzdhntljxminlist = new ArrayList<>();
            double pzdhntljxzds = 0,pzdhntljxhgds = 0;

            List<String>  pzdhnthtpjzlist = new ArrayList<>();
            List<String>  pzdhnthtmaxlist = new ArrayList<>();
            List<String>  pzdhnthtminlist = new ArrayList<>();
            double pzdhnthtzds = 0,pzdhnththgds = 0;

            Map ljxhntmap = new HashMap<>();
            for (Map<String, Object> map : pzdlist) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                //混凝土路面
                if (lmlx.contains("混凝土")){
                    if (lmlx.contains("路面")){
                        if (jcxm.contains("连接线")){
                            //路面平整度 混凝土路面 连接线
                            String sjz = map.get("设计值").toString();
                            ljxhntmap.put("ljxhntlmpzdsjz",sjz);

                            String pjz = map.get("pjz").toString();
                            pzdhntljxpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdhntljxmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdhntljxminlist.add(Min);

                            pzdhntljxzds += Double.valueOf(map.get("总点数").toString());
                            pzdhntljxhgds += Double.valueOf(map.get("合格点数").toString());

                        }else {
                            //路面平整度 混凝土路面 互通

                            String sjz = map.get("设计值").toString();
                            pzdhnthtmap.put("pzdhnthtsjz",sjz);

                            String pjz = map.get("pjz").toString();
                            pzdhnthtpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdhnthtmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdhnthtminlist.add(Min);

                            pzdhnthtzds += Double.valueOf(map.get("总点数").toString());
                            pzdhnththgds += Double.valueOf(map.get("合格点数").toString());

                        }
                    }
                    if (lmlx.contains("桥")){
                        if (jcxm.contains("连接线")){
                            //桥面平整度 混凝土桥面 连接线
                            String sjz = map.get("设计值").toString();
                            pzdhntqljxmap.put("pzdhntqljxsjz",sjz);
                            String pjz = map.get("pjz").toString();
                            pzdhntqljxpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdhntqljxmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdhntqljxminlist.add(Min);

                            pzdhntqljxzds += Double.valueOf(map.get("总点数").toString());
                            pzdhntqljxhgds += Double.valueOf(map.get("合格点数").toString());
                        }

                    }


                }

                if (lmlx.contains("沥青")){
                    if (lmlx.contains("路面")){
                        if (jcxm.contains("主线")){
                            //路面平整度 沥青路面 主线
                            String sjz = map.get("设计值").toString();
                            pzdlqzxmap.put("pzdlqzxsjz",sjz);
                            String pjz = map.get("pjz").toString();
                            pzdlqzxpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdlqzxmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdlqzxminlist.add(Min);

                            pzdlqzxzds += Double.valueOf(map.get("总点数").toString());
                            pzdlqzxhgds += Double.valueOf(map.get("合格点数").toString());

                        }else {
                            //路面平整度 沥青路面 互通及连接线
                            String sjz = map.get("设计值").toString();
                            pzdlqhtljxmap.put("pzdlqhtljxsjz",sjz);
                            String pjz = map.get("pjz").toString();
                            pzdlqhtljxpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdlqhtljxmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdlqhtljxminlist.add(Min);

                            pzdlqhtljxzds += Double.valueOf(map.get("总点数").toString());
                            pzdlqhtljxhgds += Double.valueOf(map.get("合格点数").toString());
                        }
                    }
                    if (lmlx.contains("桥")){
                        if (jcxm.contains("主线")){
                            //桥面平整度 沥青路面 主线
                            String sjz = map.get("设计值").toString();
                            pzdlqqmap.put("pzdlqqsjz",sjz);
                            String pjz = map.get("pjz").toString();
                            pzdlqqpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdlqqmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdlqqminlist.add(Min);

                            pzdlqqzds += Double.valueOf(map.get("总点数").toString());
                            pzdlqqhgds += Double.valueOf(map.get("合格点数").toString());

                        }else {
                            //桥面平整度 沥青路面 互通及连接线
                            String sjz = map.get("设计值").toString();
                            pzdlqqhtmap.put("pzdlqqhtsjz",sjz);
                            String pjz = map.get("pjz").toString();
                            pzdlqqhtpjzlist.add(pjz);
                            String Max = map.get("Max").toString();
                            pzdlqqhtmaxlist.add(Max);
                            String Min = map.get("Min").toString();
                            pzdlqqhtminlist.add(Min);

                            pzdlqqhtzds += Double.valueOf(map.get("总点数").toString());
                            pzdlqqhthgds += Double.valueOf(map.get("合格点数").toString());
                        }

                    }

                    if (lmlx.contains("隧道")){
                        //隧道 沥青路面
                        String sjz = map.get("设计值").toString();
                        pzdsdlqallmap.put("pzdsdlqallsjz",sjz);
                        String pjz = map.get("pjz").toString();
                        pzdsdlqallpjzlist.add(pjz);
                        String Max = map.get("Max").toString();
                        pzdsdlqallmaxlist.add(Max);
                        String Min = map.get("Min").toString();
                        pzdsdlqallminlist.add(Min);
                        pzdsdlqallzds += Double.valueOf(map.get("总点数").toString());
                        pzdsdlqallhgds += Double.valueOf(map.get("合格点数").toString());
                        //隧道混凝土的word无，暂时放着
                    }
                }
            }
            pzdsdlqallmap.put("pzdsdlqallpjz",getpjz(pzdsdlqallpjzlist));
            pzdsdlqallmap.put("pzdsdlqallbhfw",getczbhfw(pzdsdlqallmaxlist,pzdsdlqallminlist));
            pzdsdlqallmap.put("pzdsdlqalljghgl",pzdsdlqallzds != 0 ? df.format(pzdsdlqallhgds / pzdsdlqallzds * 100) : 0);
            data.add(pzdsdlqallmap);

            pzdlqqmap.put("pzdlqqpjz",getpjz(pzdlqqpjzlist));
            pzdlqqmap.put("pzdlqqbhfw",getczbhfw(pzdlqqmaxlist,pzdlqqminlist));
            pzdlqqmap.put("pzdlqqjghgl",pzdlqqzds != 0 ? df.format(pzdlqqhgds / pzdlqqzds * 100) : 0);
            data.add(pzdlqqmap);

            pzdlqqhtmap.put("pzdlqqhtpjz",getpjz(pzdlqqhtpjzlist));
            pzdlqqhtmap.put("pzdlqqhtbhfw",getczbhfw(pzdlqqhtmaxlist,pzdlqqhtminlist));
            pzdlqqhtmap.put("pzdlqqhtjghgl",pzdlqqhtzds != 0 ? df.format(pzdlqqhthgds / pzdlqqhtzds * 100) : 0);
            data.add(pzdlqqhtmap);

            pzdhntqljxmap.put("pzdhntqljxpjz",getpjz(pzdhntqljxpjzlist));
            pzdhntqljxmap.put("pzdhntqljxbhfw",getczbhfw(pzdhntqljxmaxlist,pzdhntqljxminlist));
            pzdhntqljxmap.put("pzdhntqljxjghgl",pzdhntqljxzds != 0 ? df.format(pzdhntqljxhgds / pzdhntqljxzds * 100) : 0);
            data.add(pzdlqqhtmap);

            pzdlqzxmap.put("pzdlqzxpjz",getpjz(pzdlqzxpjzlist));
            pzdlqzxmap.put("pzdlqzxbhfw",getczbhfw(pzdlqzxmaxlist,pzdlqzxminlist));
            pzdlqzxmap.put("pzdlqzxjghgl",pzdlqzxzds != 0 ? df.format(pzdlqzxhgds / pzdlqzxzds * 100) : 0);
            data.add(pzdlqzxmap);

            pzdlqhtljxmap.put("pzdlqhtljxpjz",getpjz(pzdlqhtljxpjzlist));
            System.out.println(pzdlqhtljxmaxlist);
            System.out.println(pzdlqhtljxminlist);
            pzdlqhtljxmap.put("pzdlqhtljxbhfw",getczbhfw(pzdlqhtljxmaxlist,pzdlqhtljxminlist));
            pzdlqhtljxmap.put("pzdlqhtljxhgl",pzdlqhtljxzds != 0 ? df.format(pzdlqhtljxhgds / pzdlqhtljxzds * 100) : 0);
            data.add(pzdlqhtljxmap);

            ljxhntmap.put("pzdhntljxpjz",getpjz(pzdhntljxpjzlist));
            ljxhntmap.put("pzdhntljxbhfw",getczbhfw(pzdhntljxmaxlist,pzdhntljxminlist));
            ljxhntmap.put("pzdhntljxhgl",pzdhntljxzds != 0 ? df.format(pzdhntljxhgds / pzdhntljxzds * 100) : 0);
            data.add(ljxhntmap);

            pzdhnthtmap.put("pzdhnthtpjz",getpjz(pzdhnthtpjzlist));
            pzdhnthtmap.put("pzdhnthtbhfw",getczbhfw(pzdhnthtmaxlist,pzdhnthtminlist));
            pzdhnthtmap.put("pzdhnththgl",pzdhnthtzds != 0 ? df.format(pzdhnththgds / pzdhnthtzds * 100) : 0);
            data.add(pzdhnthtmap);


        }

        // 摩擦系数
        List<Map<String, Object>> mcxslist = jjgZdhMcxsJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size() > 0){

            List<String>  mcxszxpjzlist = new ArrayList<>();
            List<String>  mcxszxmaxlist = new ArrayList<>();
            List<String>  mcxszxminlist = new ArrayList<>();
            double mcxszxzds = 0,mcxszxhgds = 0;

            List<String>  mcxsqmpjzlist = new ArrayList<>();
            List<String>  mcxsqmmaxlist = new ArrayList<>();
            List<String>  mcxsqmminlist = new ArrayList<>();
            double mcxsqmzds = 0,mcxsqmhgds = 0;

            List<String>  mcxssdpjzlist = new ArrayList<>();
            List<String>  mcxssdmaxlist = new ArrayList<>();
            List<String>  mcxssdminlist = new ArrayList<>();
            double mcxssdzds = 0,mcxssdhgds = 0;

            for (Map<String, Object> map : mcxslist) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("路面")){
                    String sjz = map.get("设计值").toString();
                    mcxszxmap.put("mcxszxsjz",sjz);

                    String pjz = map.get("pjz").toString();
                    mcxszxpjzlist.add(pjz);
                    String Max = map.get("Max").toString();
                    mcxszxmaxlist.add(Max);
                    String Min = map.get("Min").toString();
                    mcxszxminlist.add(Min);

                    mcxszxzds += Double.valueOf(map.get("总点数").toString());
                    mcxszxhgds += Double.valueOf(map.get("合格点数").toString());

                }
                if (lmlx.contains("隧道")){
                    String sjz = map.get("设计值").toString();
                    mcxssdmap.put("mcxssdsjz",sjz);

                    String pjz = map.get("pjz").toString();
                    mcxssdpjzlist.add(pjz);
                    String Max = map.get("Max").toString();
                    mcxssdmaxlist.add(Max);
                    String Min = map.get("Min").toString();
                    mcxssdminlist.add(Min);

                    mcxssdzds += Double.valueOf(map.get("总点数").toString());
                    mcxssdhgds += Double.valueOf(map.get("合格点数").toString());
                }
                if (lmlx.contains("桥")){
                    String sjz = map.get("设计值").toString();
                    mcxsqmmap.put("mcxsqmsjz",sjz);

                    String pjz = map.get("pjz").toString();
                    mcxsqmpjzlist.add(pjz);
                    String Max = map.get("Max").toString();
                    mcxsqmmaxlist.add(Max);
                    String Min = map.get("Min").toString();
                    mcxsqmminlist.add(Min);

                    mcxsqmzds += Double.valueOf(map.get("总点数").toString());
                    mcxsqmhgds += Double.valueOf(map.get("合格点数").toString());

                }
            }
            mcxszxmap.put("mcxszxpjz",getpjz(mcxszxpjzlist));
            mcxszxmap.put("mcxszxbhfw",getczbhfw(mcxszxmaxlist,mcxszxminlist));
            mcxszxmap.put("mcxszxjghgl",mcxszxzds != 0 ? df.format(mcxszxhgds / mcxszxzds * 100) : 0);
            data.add(mcxszxmap);

            mcxssdmap.put("mcxssdpjz",getpjz(mcxssdpjzlist));
            mcxssdmap.put("mcxssdbhfw",getczbhfw(mcxssdmaxlist,mcxssdminlist));
            mcxssdmap.put("mcxssdjghgl",mcxssdzds != 0 ? df.format(mcxssdhgds / mcxssdzds * 100) : 0);
            data.add(mcxssdmap);

            mcxsqmmap.put("mcxsqmpjz",getpjz(mcxsqmpjzlist));
            mcxsqmmap.put("mcxsqmbhfw",getczbhfw(mcxsqmmaxlist,mcxsqmminlist));
            mcxsqmmap.put("mcxsqmjghgl",mcxsqmzds != 0 ? df.format(mcxsqmhgds / mcxsqmzds * 100) : 0);
            data.add(mcxsqmmap);

        }

        List<Map<String, Object>> gzsdlist = jjgZdhGzsdJgfcService.lookJdbjg(proname);
        if (gzsdlist != null && gzsdlist.size() > 0){
            List<String>  gzsdzxpjzlist = new ArrayList<>();
            List<String>  gzsdzxmaxlist = new ArrayList<>();
            List<String>  gzsdzxminlist = new ArrayList<>();
            double gzsdzxzds = 0,gzsdzxhgds = 0;

            List<String>  gzsdqmpjzlist = new ArrayList<>();
            List<String>  gzsdqmmaxlist = new ArrayList<>();
            List<String>  gzsdqmminlist = new ArrayList<>();
            double gzsdqmzds = 0,gzsdqmhgds = 0;

            List<String>  gzsdsdpjzlist = new ArrayList<>();
            List<String>  gzsdsdmaxlist = new ArrayList<>();
            List<String>  gzsdsdminlist = new ArrayList<>();
            double gzsdsdzds = 0,gzsdsdhgds = 0;

            for (Map<String, Object> map : gzsdlist) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("路面")){
                    String sjz = map.get("设计值").toString();
                    gzsdzxmap.put("gzsdzxsjz",sjz);

                    String pjz = map.get("pjz").toString();
                    gzsdzxpjzlist.add(pjz);
                    String Max = map.get("最大值").toString();
                    gzsdzxmaxlist.add(Max);
                    String Min = map.get("最小值").toString();
                    gzsdzxminlist.add(Min);

                    gzsdzxzds += Double.valueOf(map.get("总点数").toString());
                    gzsdzxhgds += Double.valueOf(map.get("合格点数").toString());

                }
                if (lmlx.contains("隧道")){
                    String sjz = map.get("设计值").toString();
                    gzsdsdmap.put("gzsdsdsjz",sjz);

                    String pjz = map.get("pjz").toString();
                    gzsdsdpjzlist.add(pjz);
                    String Max = map.get("最大值").toString();
                    gzsdsdmaxlist.add(Max);
                    String Min = map.get("最小值").toString();
                    gzsdsdminlist.add(Min);

                    gzsdsdzds += Double.valueOf(map.get("总点数").toString());
                    gzsdsdhgds += Double.valueOf(map.get("合格点数").toString());

                }
                if (lmlx.contains("桥")){
                    String sjz = map.get("设计值").toString();
                    gzsdqmmap.put("gzsdqmsjz",sjz);

                    String pjz = map.get("pjz").toString();
                    gzsdqmpjzlist.add(pjz);
                    String Max = map.get("最大值").toString();
                    gzsdqmmaxlist.add(Max);
                    String Min = map.get("最小值").toString();
                    gzsdqmminlist.add(Min);

                    gzsdqmzds += Double.valueOf(map.get("总点数").toString());
                    gzsdqmhgds += Double.valueOf(map.get("合格点数").toString());

                }
            }
            gzsdzxmap.put("gzsdzxpjz",getpjz(gzsdzxpjzlist));
            gzsdzxmap.put("gzsdzxbhfw",getczbhfw(gzsdzxmaxlist,gzsdzxminlist));
            gzsdzxmap.put("gzsdzxjghgl",gzsdzxzds != 0 ? df.format(gzsdzxhgds / gzsdzxzds * 100) : 0);
            data.add(gzsdzxmap);

            gzsdsdmap.put("gzsdsdpjz",getpjz(gzsdsdpjzlist));
            gzsdsdmap.put("gzsdsdbhfw",getczbhfw(gzsdsdmaxlist,gzsdsdminlist));
            gzsdsdmap.put("gzsdsdjghgl",gzsdsdzds != 0 ? df.format(gzsdsdhgds / gzsdsdzds * 100) : 0);
            data.add(gzsdsdmap);

            gzsdqmmap.put("gzsdqmpjz",getpjz(gzsdqmpjzlist));
            gzsdqmmap.put("gzsdqmbhfw",getczbhfw(gzsdqmmaxlist,gzsdqmminlist));
            gzsdqmmap.put("gzsdqmjghgl",gzsdqmzds != 0 ? df.format(gzsdqmhgds / gzsdqmzds * 100) : 0);
            data.add(gzsdqmmap);

        }

        List<Map<String, Object>> tlmxlbgclist = jjgFbgcLmgcTlmxlbgcJgfcService.lookJdbjg(proname);
        if (tlmxlbgclist != null && tlmxlbgclist.size() > 0){

            double tlmxlbgczds = 0,tlmxlbgchgds = 0;
            List<String>  tlmxlbgcmaxlist = new ArrayList<>();
            List<String>  tlmxlbgcminlist = new ArrayList<>();
            for (Map<String, Object> map : tlmxlbgclist) {
                String sjz = map.get("规定值").toString();
                tlmxlbgcmap.put("tlmxlbgcsjz",sjz);

                String Max = map.get("max").toString();
                tlmxlbgcmaxlist.add(Max);
                String Min = map.get("min").toString();
                tlmxlbgcminlist.add(Min);

                tlmxlbgczds += Double.valueOf(map.get("总点数").toString());
                tlmxlbgchgds += Double.valueOf(map.get("合格点数").toString());
            }

            tlmxlbgcmap.put("tlmxlbgcpjz",getxlbgcpjz2(proname));
            tlmxlbgcmap.put("tlmxlbgcbhfw",getczbhfw(tlmxlbgcmaxlist,tlmxlbgcminlist));
            tlmxlbgcmap.put("tlmxlbgcjghgl",tlmxlbgczds != 0 ? df.format(tlmxlbgchgds / tlmxlbgczds * 100) : 0);
            data.add(tlmxlbgcmap);
        }

        return data;


    }

    /**
     *
     * @param proname
     * @return
     */
    private String getxlbgcpjz2(String proname) {
        QueryWrapper<JjgFbgcLmgcTlmxlbgcJgfc> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgFbgcLmgcTlmxlbgcJgfc> list = jjgFbgcLmgcTlmxlbgcJgfcService.list(wrapper);
        double zs = 0;
        if (list != null && list.size()>0){
            for (JjgFbgcLmgcTlmxlbgcJgfc jjgFbgcLmgcTlmxlbgcJgfc : list) {
                double scz1 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz1());
                double scz2 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz2());
                double scz3 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz3());
                double scz4 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz4());
                zs += (scz1+scz2+scz3+scz4);
            }
        }
        return String.valueOf(zs/(list.size()*4));

    }

    /**
     * 混凝土路面构造深度
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxhntlmgzsd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gethntlmgzsddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值（mm）","平均值(mm)", "实测值变化范围（mm）","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-6（2）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("gzsdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("gzsdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("gzsdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("gzsdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("gzsdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("gzsdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gethntlmgzsddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> gzsdhntlist = jjgFbgcLmgcLmgzsdsgpsfJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (gzsdhntlist != null && gzsdhntlist.size()>0){
            Map<Object, List<Map<String, Object>>> grouped = gzsdhntlist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));

            grouped.forEach((key, value) -> {
                QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                wrapper.eq("proname",proname);
                wrapper.eq("name",key);
                JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                maxlist.add(value.get(0).get("Max").toString());
                minlist.add(value.get(0).get("Min").toString());

                Map map = new HashMap<>();
                map.put("htd",key);
                map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                map.put("gzsdzds",value.get(0).get("检测点数"));
                map.put("gzsdhgds",value.get(0).get("合格点数"));
                map.put("gzsdhgl",value.get(0).get("合格率"));
                map.put("gzsdgdz",value.get(0).get("规定值"));
                map.put("gzsdpjz",getxlbgcpjz(proname,key.toString()));
                map.put("gzsdbhfw",getczbhfw(maxlist,minlist));
                resultlist.add(map);

            });
        }
        return resultlist;
    }

    /**
     * 沥青隧道构造深度
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxsdgzsd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlqsdgzsddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值（mm）","平均值(mm)", "实测值变化范围（mm）","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-12}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("lmgzsdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("lmgzsdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("lmgzsdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("lmgzsdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("lmgzsdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("lmgzsdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }


    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getlqsdgzsddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        List<Map<String, Object>> mcxslist = jjgZdhGzsdJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = mcxslist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("隧道")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("最大值").toString());
                        minlist.add(map.get("最小值").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("lmgzsdzds",decf.format(czzds));
                    map.put("lmgzsdhgds",decf.format(czhgds));
                    map.put("lmgzsdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("lmgzsdgdz",value.get(0).get("设计值"));
                    map.put("lmgzsdpjz",getpjz(pjzlist));
                    map.put("lmgzsdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }
            });
        }
        return resultlist;
    }

    /**
     * 沥青桥面构造深度
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxqmgzsd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlqqmgzsddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值（mm）","平均值(mm)", "实测值变化范围（mm）","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-8（1）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("lmgzsdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("lmgzsdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("lmgzsdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("lmgzsdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("lmgzsdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("lmgzsdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }


    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getlqqmgzsddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        List<Map<String, Object>> mcxslist = jjgZdhGzsdJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = mcxslist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("桥")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("最大值").toString());
                        minlist.add(map.get("最小值").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("lmgzsdzds",decf.format(czzds));
                    map.put("lmgzsdhgds",decf.format(czhgds));
                    map.put("lmgzsdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("lmgzsdgdz",value.get(0).get("设计值"));
                    map.put("lmgzsdpjz",getpjz(pjzlist));
                    map.put("lmgzsdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }
            });
        }
        return resultlist;
    }

    /**
     * 沥青路面构造深度
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlmgzsd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlqlmgzsddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值（mm）","平均值(mm)", "实测值变化范围（mm）","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-6（1）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("lmgzsdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("lmgzsdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("lmgzsdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("lmgzsdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("lmgzsdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("lmgzsdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getlqlmgzsddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        List<Map<String, Object>> mcxslist = jjgZdhGzsdJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = mcxslist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("路面")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("最大值").toString());
                        minlist.add(map.get("最小值").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("lmgzsdzds",decf.format(czzds));
                    map.put("lmgzsdhgds",decf.format(czhgds));
                    map.put("lmgzsdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("lmgzsdgdz",value.get(0).get("设计值"));
                    map.put("lmgzsdpjz",getpjz(pjzlist));
                    map.put("lmgzsdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }
            });
        }
        return resultlist;
    }

    /**
     * 隧道摩擦系数
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxsdmcxs(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gettsdmcxsdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值", "实测值变化范围","平均值","代表值变化范围","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-11}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("lmmcxsgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("lmmcxssczbhfw").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("lmmcxspjz").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("lmmcxsdbzbhfw").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("lmmcxszds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("lmmcxshgds").toString());
                                break;
                            case 8:
                                p.appendText(rowData.get("lmmcxshgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gettsdmcxsdata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        List<Map<String, Object>> mcxslist = jjgZdhMcxsJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = mcxslist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("隧道")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("lmmcxszds",decf.format(czzds));
                    map.put("lmmcxshgds",decf.format(czhgds));
                    map.put("lmmcxshgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("lmmcxsgdz",value.get(0).get("设计值"));
                    map.put("lmmcxspjz",getpjz(pjzlist));
                    map.put("lmmcxsdbzbhfw",getczbhfw(maxlist,minlist));
                    map.put("lmmcxssczbhfw",0);
                    resultlist.add(map);
                }
            });
        }
        return resultlist;
    }

    /**
     * 桥面摩擦系数
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxqmmcxs(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gettqmmcxsdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值", "实测值变化范围","平均值","代表值变化范围","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-9}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("lmmcxsgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("lmmcxssczbhfw").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("lmmcxspjz").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("lmmcxsdbzbhfw").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("lmmcxszds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("lmmcxshgds").toString());
                                break;
                            case 8:
                                p.appendText(rowData.get("lmmcxshgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gettqmmcxsdata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        List<Map<String, Object>> mcxslist = jjgZdhMcxsJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = mcxslist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("桥")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("lmmcxszds",decf.format(czzds));
                    map.put("lmmcxshgds",decf.format(czhgds));
                    map.put("lmmcxshgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("lmmcxsgdz",value.get(0).get("设计值"));
                    map.put("lmmcxspjz",getpjz(pjzlist));
                    map.put("lmmcxsdbzbhfw",getczbhfw(maxlist,minlist));
                    map.put("lmmcxssczbhfw",0);
                    resultlist.add(map);
                }
            });
        }
        return resultlist;
    }


    /**
     * 路面摩擦系数
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlmmcxs(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gettlmmcxsdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值", "实测值变化范围","平均值","代表值变化范围","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-5}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("lmmcxsgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("lmmcxssczbhfw").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("lmmcxspjz").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("lmmcxsdbzbhfw").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("lmmcxszds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("lmmcxshgds").toString());
                                break;
                            case 8:
                                p.appendText(rowData.get("lmmcxshgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gettlmmcxsdata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultlist = new ArrayList<>();
        List<Map<String, Object>> mcxslist = jjgZdhMcxsJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = mcxslist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("路面")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("lmmcxszds",decf.format(czzds));
                    map.put("lmmcxshgds",decf.format(czhgds));
                    map.put("lmmcxshgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("lmmcxsgdz",value.get(0).get("设计值"));
                    map.put("lmmcxspjz",getpjz(pjzlist));
                    map.put("lmmcxsdbzbhfw",getczbhfw(maxlist,minlist));
                    map.put("lmmcxssczbhfw",0);
                    resultlist.add(map);
                }
            });
        }
        return resultlist;
    }




    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxtlmxlbgc(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gettlmxlbgsdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-14}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("gcgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("gcpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("gcbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("gczds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("gchgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("gchgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }




    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gettlmxlbgsdata(String proname) throws IOException {
        List<Map<String, Object>> xlbgclist = jjgFbgcLmgcTlmxlbgcJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (xlbgclist != null && xlbgclist.size()>0){
            Map<Object, List<Map<String, Object>>> grouped = xlbgclist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            //{LJ-1=[{合格点数=24, min=1.0, 总点数=24, max=2.0, 规定值=2.0, htd=LJ-1, 合格率=100.00}]}
            grouped.forEach((key, value) -> {
                QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                wrapper.eq("proname",proname);
                wrapper.eq("name",key);
                JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                maxlist.add(value.get(0).get("max").toString());
                minlist.add(value.get(0).get("min").toString());

                Map map = new HashMap<>();
                map.put("htd",key);
                map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                map.put("gczds",value.get(0).get("总点数"));
                map.put("gchgds",value.get(0).get("合格点数"));
                map.put("gchgl",value.get(0).get("合格率"));
                map.put("gcgdz",value.get(0).get("规定值"));
                map.put("gcpjz",getxlbgcpjz(proname,key.toString()));
                map.put("gcbhfw",getczbhfw(maxlist,minlist));
                resultlist.add(map);

            });
        }
        return resultlist;

    }

    /**
     *
     * @param proname
     * @param key
     * @return
     */
    private String getxlbgcpjz(String proname, String key) {
        QueryWrapper<JjgFbgcLmgcTlmxlbgcJgfc> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",key);
        List<JjgFbgcLmgcTlmxlbgcJgfc> list = jjgFbgcLmgcTlmxlbgcJgfcService.list(wrapper);
        double zs = 0;
        if (list != null && list.size()>0){
            for (JjgFbgcLmgcTlmxlbgcJgfc jjgFbgcLmgcTlmxlbgcJgfc : list) {
                double scz1 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz1());
                double scz2 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz2());
                double scz3 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz3());
                double scz4 = Double.valueOf(jjgFbgcLmgcTlmxlbgcJgfc.getScz4());
                zs += (scz1+scz2+scz3+scz4);
            }
        }
        return String.valueOf(zs/(list.size()*4));


    }


    /**
     * 沥青隧道平整度
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlqsdpzd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlqsdpzddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-10}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("pzdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("pzdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("pzdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("pzdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("pzdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("pzdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     * 沥青隧道平整度
     * 沥青匝道隧道  沥青隧道
     * @param proname
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlqsdpzddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (pzdlist != null && pzdlist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = pzdlist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("沥青匝道隧道") || lmlx.contains("沥青隧道")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("pzdzds",decf.format(czzds));
                    map.put("pzdhgds",decf.format(czhgds));
                    map.put("pzdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("pzdgdz",value.get(0).get("设计值"));
                    map.put("pzdpjz",getpjz(pjzlist));
                    map.put("pzdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });
        }
        return resultlist;


    }


    /**
     *混凝土桥面平整度
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxhntqmpzd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gethntqmpzddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-7（2）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("pzdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("pzdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("pzdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("pzdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("pzdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("pzdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     * 混凝土桥面平整度
     * 混凝土匝道桥  混凝土桥
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gethntqmpzddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();

        if (pzdlist != null && pzdlist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = pzdlist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("混凝土匝道桥") || lmlx.contains("混凝土桥")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("pzdzds",decf.format(czzds));
                    map.put("pzdhgds",decf.format(czhgds));
                    map.put("pzdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("pzdgdz",value.get(0).get("设计值"));
                    map.put("pzdpjz",getpjz(pjzlist));
                    map.put("pzdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });
        }
        return resultlist;


    }


    /**
     *沥青桥面平整度
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlqqmpzd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlqqmpzddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-7（1）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("pzdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("pzdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("pzdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("pzdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("pzdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("pzdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }

    }

    /**
     * 沥青桥面平整度
     * 沥青匝道桥  沥青桥
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getlqqmpzddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (pzdlist != null && pzdlist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = pzdlist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("沥青匝道桥") || lmlx.contains("沥青桥")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("pzdzds",decf.format(czzds));
                    map.put("pzdhgds",decf.format(czhgds));
                    map.put("pzdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("pzdgdz",value.get(0).get("设计值"));
                    map.put("pzdpjz",getpjz(pjzlist));
                    map.put("pzdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });
        }
        return resultlist;


    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxhntlmpzd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gethntlmpzddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-4（2）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("pzdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("pzdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("pzdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("pzdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("pzdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("pzdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }

    }

    /**
     * 混凝土路面平整度
     *  混凝土路面   混凝土匝道   混凝土收费站 混凝土隧道 混凝土匝道隧道
     * @param proname
     * @return
     */
    private List<Map<String, Object>> gethntlmpzddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (pzdlist != null && pzdlist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = pzdlist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("混凝土路面") || lmlx.contains("混凝土匝道") || lmlx.contains("混凝土收费站") || lmlx.contains("混凝土隧道") || lmlx.contains("混凝土匝道隧道")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("pzdzds",decf.format(czzds));
                    map.put("pzdhgds",decf.format(czhgds));
                    map.put("pzdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("pzdgdz",value.get(0).get("设计值"));
                    map.put("pzdpjz",getpjz(pjzlist));
                    map.put("pzdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });
        }
        return resultlist;
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlqlmpzd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlqlmpzddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-4（1）}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("pzdgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("pzdpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("pzdbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("pzdzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("pzdhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("pzdhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }

    }

    /**
     * 沥青路面平整度
     * 沥青路面   沥青匝道
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getlqlmpzddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (pzdlist != null && pzdlist.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = pzdlist.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("沥青路面") || lmlx.contains("沥青匝道")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);
                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("pzdzds",decf.format(czzds));
                    map.put("pzdhgds",decf.format(czhgds));
                    map.put("pzdhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("pzdgdz",value.get(0).get("设计值"));
                    map.put("pzdpjz",getpjz(pjzlist));
                    map.put("pzdbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });
        }
        return resultlist;
    }

    /**
     * 沥青隧道车辙
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlqsdcz(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getczsddata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-13}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("czgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("czpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("czbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("czzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("czhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("czhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     * 沥青隧道车辙
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getczsddata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> data = jjgZdhCzJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (data != null && data.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = data.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("隧道")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);

                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("czzds",decf.format(czzds));
                    map.put("czhgds",decf.format(czhgds));
                    map.put("czhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("czgdz",value.get(0).get("设计值"));
                    map.put("czpjz",getpjz(pjzlist));
                    map.put("czbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });

        }
        return resultlist;
    }

    /**
     * 沥青路面车辙
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlqlmcz(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getczcdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "规定值或允许偏差(mm)", "平均值（mm）", "实测值变化范围(mm)","检测点数","合格点数","合格率（%）"};
            String s = "${表4.1-3}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, header.length);
                Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                Body body = paragraph.ownerTextBody();
                int index = body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("czgdz").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("czpjz").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("czbhfw").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("czzds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("czhgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("czhgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index, table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }


    }

    /**
     * 沥青路面车辙
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getczcdata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> data = jjgZdhCzJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (data != null && data.size()>0){
            AtomicBoolean a = new AtomicBoolean(false);
            Map<Object, List<Map<String, Object>>> grouped = data.stream()
                    .collect(Collectors.groupingBy(
                            // 提取Map中htd键的值作为分组依据
                            map -> (map.containsKey("htd")) ? map.get("htd") : null,
                            // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                            Collectors.toList()
                    ));
            grouped.forEach((key, value) -> {
                double czzds = 0,czhgds = 0;
                List<String> pjzlist = new ArrayList<>();
                List<String> maxlist = new ArrayList<>();
                List<String> minlist = new ArrayList<>();
                for (Map<String, Object> map : value) {
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("路面") || lmlx.contains("桥")){
                        a.set(true);
                        czzds += Double.valueOf(map.get("总点数").toString());
                        czhgds += Double.valueOf(map.get("合格点数").toString());
                        String pjz = map.get("pjz").toString();
                        pjzlist.add(pjz);
                        maxlist.add(map.get("Max").toString());
                        minlist.add(map.get("Min").toString());
                    }
                }
                if (a.get()){
                    QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
                    wrapper.eq("proname",proname);
                    wrapper.eq("name",key);
                    JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);

                    Map map = new HashMap<>();
                    map.put("htd",key);
                    map.put("qzzh",one.getZhq()+"~"+one.getZhz());
                    map.put("czzds",decf.format(czzds));
                    map.put("czhgds",decf.format(czhgds));
                    map.put("czhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                    map.put("czgdz",value.get(0).get("设计值"));
                    map.put("czpjz",getpjz(pjzlist));
                    map.put("czbhfw",getczbhfw(maxlist,minlist));
                    resultlist.add(map);
                }

            });

        }
        return resultlist;
    }

    /**
     *
     * @param maxlist
     * @param minlist
     * @return
     */
    private String getczbhfw(List<String> maxlist, List<String> minlist) {
        double min = Double.MAX_VALUE;
        double max = Double.MIN_VALUE;
        if (maxlist != null && !maxlist.isEmpty()){
            for (String s : maxlist) {
                Double v = Double.valueOf(s);
                if (v > max) {
                    max = v;
                }
            }
        }else {
            max = 0;
        }
        if (minlist != null && !minlist.isEmpty()){
            for (String s : minlist) {
                Double v = Double.valueOf(s);
                if (v < min) {
                    min = v;
                }
            }

        }else {
            min = 0;
        }
        return String.format("%.2f~%.2f", min, max);
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoDocxlqwcfcjg(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getwcfcdata(proname);
        if (data != null && data.size()>0){
            String[] header = {"合同段", "起讫桩号", "设计值(0.01mm)", "平均值(0.01mm)", "代表值变化范围(0.01mm)","检测单元数","合格单元数","合格率（%）"};
            String s = "${表4.1-2}";
            TextSelection textSelection = findStringInPages(xw,s,20,35);
            if (textSelection != null) {
                com.spire.doc.Table table=new Table(xw,true);
                table.resetCells(data.size()+1, header.length);
                Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                Body body=paragraph.ownerTextBody();
                int index=body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                for (int i = 0; i < header.length; i++) {
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString()+"合同段");
                                break;
                            case 1:
                                p.appendText(rowData.get("qzzh").toString());
                                break;
                            case 2:
                                p.appendText(rowData.get("sjzlist").toString());
                                break;
                            case 3:
                                p.appendText(rowData.get("pjzlist").toString());
                                break;
                            case 4:
                                p.appendText(rowData.get("dbzlist").toString());
                                break;
                            case 5:
                                p.appendText(rowData.get("wczds").toString());
                                break;
                            case 6:
                                p.appendText(rowData.get("wchgds").toString());
                                break;
                            case 7:
                                p.appendText(rowData.get("wchgl").toString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index,table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
            }
        }
    }

    /**
     *
     * @param proname
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getwcfcdata(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        CommonInfoVo commonInfoVo = new CommonInfoVo();
        commonInfoVo.setProname(proname);
        List<Map<String, Object>> maps1 = jjgFbgcLmgcLmwcJgfcService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> maps2 = jjgFbgcLmgcLmwcLcfJgfcService.lookJdbjg(commonInfoVo);
        //设计值(规定值) 平均值  代表值变化范围(list中取最大和最小值) 总点数 ...

        List<Map<String, Object>> resultlist = new ArrayList<>();
        if (maps1 != null && maps1.size() > 0){
            for (Map<String, Object> stringObjectMap : maps1) {
                String htd = stringObjectMap.get("htd").toString();
                double wczds = Double.valueOf(stringObjectMap.get("检测单元数").toString());
                double wxhgds = Double.valueOf(stringObjectMap.get("合格单元数").toString());
                String sjz = stringObjectMap.get("规定值").toString();
                String dbz = stringObjectMap.get("代表值").toString();
                String pjz = stringObjectMap.get("平均值").toString();

                Map wcmap = new HashMap<>();
                wcmap.put("htd",htd);
                wcmap.put("wczds",wczds);
                wcmap.put("wxhgds",wxhgds);
                wcmap.put("sjz",sjz);
                wcmap.put("dbz",dbz);
                wcmap.put("pjz",pjz);
                resultlist.add(wcmap);
            }
        }
        if (maps2 != null && maps2.size() > 0){
            for (Map<String, Object> stringObjectMap : maps2) {
                String htd = stringObjectMap.get("htd").toString();
                double wczds = Double.valueOf(stringObjectMap.get("检测单元数").toString());
                double wxhgds = Double.valueOf(stringObjectMap.get("合格单元数").toString());
                String sjz = stringObjectMap.get("规定值").toString();
                String dbz = stringObjectMap.get("代表值").toString();
                String pjz = stringObjectMap.get("平均值").toString();

                Map wcmap = new HashMap<>();
                wcmap.put("htd",htd);
                wcmap.put("wczds",wczds);
                wcmap.put("wxhgds",wxhgds);
                wcmap.put("sjz",sjz);
                wcmap.put("dbz",dbz);
                wcmap.put("pjz",pjz);
                resultlist.add(wcmap);
            }
        }
        List<Map<String, Object>> result = new ArrayList<>();
        //按合同段进行分组
        Map<Object, List<Map<String, Object>>> grouped = resultlist.stream()
                .collect(Collectors.groupingBy(
                        // 提取Map中htd键的值作为分组依据
                        map -> (map.containsKey("htd")) ? map.get("htd") : null,
                        // 可以使用LinkedHashMap来保持插入顺序（如果需要）
                        Collectors.toList()
                ));
        grouped.forEach((key, value) -> {

            double wczds = 0,wchgds = 0;
            List<String> pjzlist = new ArrayList<>();
            List<String> sjzlist = new ArrayList<>();
            List<String> dbzlist = new ArrayList<>();
            for (Map<String, Object> stringObjectMap : value) {

                wczds += Double.valueOf(stringObjectMap.get("wczds").toString());
                wchgds += Double.valueOf(stringObjectMap.get("wxhgds").toString());
                String pjz = stringObjectMap.get("pjz").toString();
                String sjz = stringObjectMap.get("sjz").toString();
                String dbz = stringObjectMap.get("dbz").toString();
                pjzlist.add(pjz.substring(1,pjz.length()-1));
                sjzlist.add(sjz.substring(1,sjz.length()-1));
                dbzlist.add(dbz.substring(1,dbz.length()-1));

            }
            QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
            wrapper.eq("proname",proname);
            wrapper.eq("name",key);
            JjgJgHtdinfo one = jjgJgHtdinfoService.getOne(wrapper);

            Map map = new HashMap<>();
            map.put("htd",key);
            map.put("qzzh",one.getZhq()+"~"+one.getZhz());
            map.put("wczds",decf.format(wczds));
            map.put("wchgds",decf.format(wchgds));
            map.put("wchgl",wczds != 0 ? df.format(wchgds / wczds * 100) : 0);
            map.put("pjzlist",pjzlist);
            map.put("sjzlist",sjzlist);
            map.put("dbzlist",dbzlist);
            result.add(map);

        });
        if (result != null){
            for (Map<String, Object> stringObjectMap : result) {
                String dbzstring = stringObjectMap.get("dbzlist").toString();
                String pjzliststr = stringObjectMap.get("pjzlist").toString();
                String sjzlist = stringObjectMap.get("sjzlist").toString();

                if (dbzstring != null && !dbzstring.isEmpty()) {
                    String substring = dbzstring.substring(1, dbzstring.length() - 1);
                    List<String> list = Arrays.asList(substring.split(","));
                    String getbhfw = getbhfw(list);
                    stringObjectMap.put("dbzlist",getbhfw);
                }
                if (pjzliststr != null && !pjzliststr.isEmpty()) {
                    String substring = pjzliststr.substring(1, pjzliststr.length() - 1);
                    List<String> list = Arrays.asList(substring.split(","));
                    String pjz = getpjz(list);
                    stringObjectMap.put("pjzlist",pjz);
                }
                if (sjzlist != null && !sjzlist.isEmpty()) {
                    String substring = sjzlist.substring(1, sjzlist.length() - 1).replaceAll("\\s+", "");;
                    List<String> list = Arrays.asList(substring.split(","));
                    Set<String> set = new HashSet<>(list);
                    //List<String> distinctList = new ArrayList<>(set);
                    stringObjectMap.put("sjzlist",set);
                }
            }
        }


        return result;
    }

    /**
     *
     * @param list
     * @return
     */
    private String getpjz(List<String> list) {
        if (list != null && !list.isEmpty()){
            double pjz = 0;
            for (String s : list) {
                Double v = Double.valueOf(s);
                pjz += v;
            }
            double re = pjz/list.size();
            return String.valueOf(String.format("%.2f", re));
        }else {
            return "0";
        }
    }

    /**
     *
     * @param list
     * @return
     */
    public String getbhfw(List<String> list) {
        if (list != null){
            double min = Double.MAX_VALUE;
            double max = Double.MIN_VALUE;
            for (String s : list) {
                Double v = Double.valueOf(s);
                if (v < min) {
                    min = v;
                }
                if (v > max) {
                    max = v;
                }
            }
            return String.format("%.2f~%.2f", min, max);
        }else {
            return "";
        }
    }




    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtojahtd(File f, Document xw, String proname) {
        QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(wrapper);
        if (list != null && list.size()>0) {
            List<JjgJgHtdinfo> resultList = new ArrayList<>();
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                if (lx.contains("交安工程")) {
                    resultList.add(jjgJgHtdinfo);
                }
            }
            if (resultList != null && resultList.size() > 0) {
                String[] header = {"序号", "合同段", "起讫桩号", "决算金额（万元）", "施工单位", "监理单位"};
                String s = "${表1.3-3}";
                //TextSelection textSelection = xw.findString(s, true, true);
                TextSelection textSelection = findStringInPages(xw, s, 8, 13);
                if (textSelection != null) {
                    com.spire.doc.Table table = new Table(xw, true);
                    table.resetCells(resultList.size() + 1, header.length);
                    Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                    Body body = paragraph.ownerTextBody();
                    int index = body.getChildObjects().indexOf(paragraph);
                    //将第一行设置为表格标题
                    TableRow row = table.getRows().get(0);
                    row.isHeader(true);
                    row.setHeight(50);
                    row.setHeightType(TableRowHeightType.Exactly);
                    for (int i = 0; i < header.length; i++) {
                        row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row.getCells().get(i).addParagraph();
                        p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                        TextRange txtRange = p.appendText(header[i]);
                        txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                        txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                        txtRange.getCharacterFormat().setBold(true);
                    }
                    //添加数据
                    for (int rowIdx = 0; rowIdx < resultList.size(); rowIdx++) {
                        TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                        //row1.setHeight(80);
                        row1.setHeightType(TableRowHeightType.Exactly);
                        row1.setHeight(34);
                        JjgJgHtdinfo rowData = resultList.get(rowIdx);


                        for (int colIdx = 0; colIdx < header.length; colIdx++) {
                            row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                            Paragraph p = row1.getCells().get(colIdx).addParagraph();
                            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                            // 根据表头的不同，设置相应的数据
                            switch (colIdx) {
                                case 0:
                                    p.appendText(String.valueOf(rowIdx + 1));
                                    break;
                                case 1:
                                    p.appendText(rowData.getName() + "合同段");
                                    break;
                                case 2:
                                    p.appendText(rowData.getZhq() + "~" + rowData.getZhz());
                                    break;
                                case 3:
                                    p.appendText("");
                                    break;
                                case 4:
                                    p.appendText(rowData.getSgdw());
                                    break;
                                case 5:
                                    p.appendText(rowData.getJldw());
                                    break;
                                default:
                                    break;
                            }

                        }
                    }
                    body.getChildObjects().remove(paragraph);
                    body.getChildObjects().insert(index, table);
                    //列宽自动适应内容
                    table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                }
            }
        }
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoLmhtd(File f, Document xw, String proname) {
        QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(wrapper);
        if (list != null && list.size()>0){
            List<JjgJgHtdinfo> resultList = new ArrayList<>();
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                if (lx.contains("路面工程")){
                    resultList.add(jjgJgHtdinfo);
                }
            }
            if (resultList != null && resultList.size()>0){
                String[] header = {"序号", "合同段", "起讫桩号", "决算金额（万元）", "施工单位", "监理单位"};
                String s = "${表1.3-2}";
                //TextSelection textSelection = xw.findString(s, true, true);
                TextSelection textSelection = findStringInPages(xw,s,8,13);
                if (textSelection != null) {
                    com.spire.doc.Table table=new Table(xw,true);
                    table.resetCells(resultList.size()+1, header.length);
                    Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                    Body body=paragraph.ownerTextBody();
                    int index=body.getChildObjects().indexOf(paragraph);
                    //将第一行设置为表格标题
                    TableRow row = table.getRows().get(0);
                    row.isHeader(true);
                    row.setHeight(50);
                    row.setHeightType(TableRowHeightType.Exactly);
                    for (int i = 0; i < header.length; i++) {
                        row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row.getCells().get(i).addParagraph();
                        p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                        TextRange txtRange = p.appendText(header[i]);
                        txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                        txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                        txtRange.getCharacterFormat().setBold(true);
                    }
                    //添加数据
                    for (int rowIdx = 0; rowIdx < resultList.size(); rowIdx++) {
                        TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                        //row1.setHeight(80);
                        row1.setHeightType(TableRowHeightType.Exactly);
                        row1.setHeight(34);
                        JjgJgHtdinfo rowData = resultList.get(rowIdx);


                        for (int colIdx = 0; colIdx < header.length; colIdx++) {
                            row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                            Paragraph p = row1.getCells().get(colIdx).addParagraph();
                            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                            // 根据表头的不同，设置相应的数据
                            switch (colIdx) {
                                case 0:
                                    p.appendText(String.valueOf(rowIdx+1));
                                    break;
                                case 1:
                                    p.appendText(rowData.getName()+"合同段");
                                    break;
                                case 2:
                                    p.appendText(rowData.getZhq()+"~"+rowData.getZhz());
                                    break;
                                case 3:
                                    p.appendText("");
                                    break;
                                case 4:
                                    p.appendText(rowData.getSgdw());
                                    break;
                                case 5:
                                    p.appendText(rowData.getJldw());
                                    break;
                                default:
                                    break;
                            }

                        }
                    }
                    body.getChildObjects().remove(paragraph);
                    body.getChildObjects().insert(index,table);
                    //列宽自动适应内容
                    table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                }
            }
        }
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void DBtoLjqshtd(File f, Document xw, String proname) {
        QueryWrapper<JjgJgHtdinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname", proname);
        List<JjgJgHtdinfo> list = jjgJgHtdinfoService.list(wrapper);
        if (list != null && list.size()>0){
            List<JjgJgHtdinfo> resultList = new ArrayList<>();
            for (JjgJgHtdinfo jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                if (lx.contains("路基工程") || lx.contains("桥梁工程") || lx.contains("隧道工程")){
                    resultList.add(jjgJgHtdinfo);
                }
            }
            if (resultList != null && resultList.size()>0){
                String[] header = {"序号", "合同段", "起讫桩号", "决算金额（万元）", "施工单位", "监理单位"};
                String s = "${表1.3-1}";
                //TextSelection textSelection = xw.findString(s, true, true);
                TextSelection textSelection = findStringInPages(xw,s,8,13);
                if (textSelection != null) {
                    com.spire.doc.Table table=new Table(xw,true);
                    table.resetCells(resultList.size()+1, header.length);
                    Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                    Body body=paragraph.ownerTextBody();
                    int index=body.getChildObjects().indexOf(paragraph);
                    //将第一行设置为表格标题
                    TableRow row = table.getRows().get(0);
                    row.isHeader(true);
                    row.setHeight(50);
                    row.setHeightType(TableRowHeightType.Exactly);
                    for (int i = 0; i < header.length; i++) {
                        row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row.getCells().get(i).addParagraph();
                        p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                        TextRange txtRange = p.appendText(header[i]);

                        txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                        txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                        txtRange.getCharacterFormat().setBold(true);
                    }
                    //添加数据
                    for (int rowIdx = 0; rowIdx < resultList.size(); rowIdx++) {
                        TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                        //row1.setHeight(80);
                        row1.setHeightType(TableRowHeightType.Exactly);
                        row1.setHeight(34);
                        JjgJgHtdinfo rowData = resultList.get(rowIdx);


                        for (int colIdx = 0; colIdx < header.length; colIdx++) {
                            row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                            Paragraph p = row1.getCells().get(colIdx).addParagraph();
                            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                            // 根据表头的不同，设置相应的数据
                            switch (colIdx) {
                                case 0:
                                    p.appendText(String.valueOf(rowIdx+1));
                                    break;
                                case 1:
                                    p.appendText(rowData.getName()+"合同段");
                                    break;
                                case 2:
                                    p.appendText(rowData.getZhq()+"~"+rowData.getZhz());
                                    break;
                                case 3:
                                    p.appendText("");
                                    break;
                                case 4:
                                    p.appendText(rowData.getSgdw());
                                    break;
                                case 5:
                                    p.appendText(rowData.getJldw());
                                    break;
                                default:
                                    break;
                            }

                        }
                    }
                    body.getChildObjects().remove(paragraph);
                    body.getChildObjects().insert(index,table);
                    //列宽自动适应内容
                    table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                }
            }


        }


    }


    /**
     *
     * @param f
     * @param xw
     * @param proname
     * @throws IOException
     */
    private void DBtoDocxJgfc(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> list = getDocxData(proname);
        if (list != null && list.size()>0){
            for (Map<String, Object> map : list) {
                for (Map.Entry<String, Object> entry : map.entrySet()) {
                    String s = "${"+entry.getKey()+"}";
                    String value = entry.getValue().toString();
                    xw.replace(s, value, false, true);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                }
            }

        }

    }




    /**
     *
     * @param proname
     */
    private List<Map<String, Object>> getDocxData(String proname) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //沥青路面弯沉
        CommonInfoVo commonInfoVo = new CommonInfoVo();
        commonInfoVo.setProname(proname);
        List<Map<String, Object>> maps1 = jjgFbgcLmgcLmwcJgfcService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> maps2 = jjgFbgcLmgcLmwcLcfJgfcService.lookJdbjg(commonInfoVo);
        double wczds = 0,wxhgds = 0;
        if (maps1 != null && maps1.size() > 0){
            for (Map<String, Object> stringObjectMap : maps1) {
                wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                wxhgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
            }
        }
        if (maps2 != null && maps2.size() > 0){
            for (Map<String, Object> stringObjectMap : maps2) {
                wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                wxhgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
            }
        }
        Map wcmap = new HashMap<>();
        wcmap.put("wczds",decf.format(wczds));
        wcmap.put("wchgds",decf.format(wxhgds));
        wcmap.put("wchgl",wczds != 0 ? df.format(wxhgds / wczds * 100) : 0);
        resultList.add(wcmap);

        //沥青路面车辙： 沥青路面 隧道路面
        List<Map<String, Object>> lookJdbjg = jjgZdhCzJgfcService.lookJdbjg(proname);
        if (lookJdbjg != null && lookJdbjg.size() > 0){
            double czzxzds = 0,czzxhgds = 0,czsdzds = 0,czsdhgds = 0,czqmzds = 0,czqmhgds = 0;
            for (Map<String, Object> stringObjectMap : lookJdbjg) {
                String lmlx = stringObjectMap.get("路面类型").toString();
                /*if (lmlx.contains("桥") ){
                    czqmzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    czqmhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }else */
                if (lmlx.contains("隧道")){
                    czsdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    czsdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }else if (lmlx.contains("路面") || lmlx.contains("桥")){
                    czzxzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    czzxhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }
            }
            Map czmap1 = new HashMap<>();
            czmap1.put("czlqzds",decf.format(czzxzds));
            czmap1.put("czlqhgds",decf.format(czzxhgds));
            czmap1.put("czlqhgl",czzxzds != 0 ? df.format(czzxhgds / czzxzds * 100) : 0);
            /*Map czmap2 = new HashMap<>();
            czmap2.put("czqmzds",decf.format(czqmzds));
            czmap2.put("czqmhgds",decf.format(czqmhgds));
            czmap2.put("czqmhgl",czqmzds != 0 ? df.format(czqmhgds / czqmzds * 100) : 0);*/
            Map czmap3 = new HashMap<>();
            czmap3.put("czsdzds",decf.format(czsdzds));
            czmap3.put("czsdhgds",decf.format(czsdhgds));
            czmap3.put("czsdhgl",czsdzds != 0 ? df.format(czsdhgds / czsdzds * 100) : 0);
            resultList.add(czmap1);
            //resultList.add(czmap2);
            resultList.add(czmap3);
        }

        //平整度：  沥青路面  混凝土路面 沥青桥面 混凝土桥面 隧道沥青路面
        List<Map<String, Object>> pzdlist = jjgZdhPzdJgfcService.lookJdbjg(proname);
        if (pzdlist != null && pzdlist.size() > 0){
            double pzdhntqzds = 0,pzdhntqhgds = 0,pzdhntsdzds = 0,pzdhntsdhgds = 0,pzdhntzds = 0,pzdhnthgds = 0;
            double pzdqzds = 0,pzdqhgds = 0,pzdsdlqzds = 0,pzdsdlqhgds = 0,pzdlqzds = 0,pzdlqhgds = 0;
            for (Map<String, Object> stringObjectMap : pzdlist) {
                String lmlx = stringObjectMap.get("路面类型").toString();
                if (lmlx.contains("混凝土")){
                    if (lmlx.contains("桥")){
                        pzdhntqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                        pzdhntqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                    }else if (lmlx.contains("隧道")){
                        pzdhntsdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                        pzdhntsdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                    }else if (lmlx.contains("路面")){
                        pzdhntzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                        pzdhnthgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                    }
                }else if (lmlx.contains("沥青")){
                    if (lmlx.contains("桥")){
                        pzdqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                        pzdqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                    }else if (lmlx.contains("隧道")){
                        pzdsdlqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                        pzdsdlqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                    }else if (lmlx.contains("路面")){
                        pzdlqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                        pzdlqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                    }
                }

            }
            Map czmap1 = new HashMap<>();
            czmap1.put("pzdhntzds",decf.format(pzdhntzds));
            czmap1.put("pzdhnthgds",decf.format(pzdhnthgds));
            czmap1.put("pzdhnthgl",pzdhntzds != 0 ? df.format(pzdhnthgds / pzdhntzds * 100) : 0);
            Map czmap2 = new HashMap<>();
            czmap2.put("pzdhntsdzds",decf.format(pzdhntsdzds));
            czmap2.put("pzdhntsdhgds",decf.format(pzdhntsdhgds));
            czmap2.put("pzdhntsdhgl",pzdhntsdzds != 0 ? df.format(pzdhntsdhgds / pzdhntsdzds * 100) : 0);
            Map czmap3 = new HashMap<>();
            czmap3.put("pzdhntqzds",decf.format(pzdhntqzds));
            czmap3.put("pzdhntqhgds",decf.format(pzdhntqhgds));
            czmap3.put("pzdhntqhgl",pzdhntqzds != 0 ? df.format(pzdhntqhgds / pzdhntqzds * 100) : 0);
            Map czmap4 = new HashMap<>();
            czmap4.put("pzdlqzds",decf.format(pzdlqzds));
            czmap4.put("pzdlqhgds",decf.format(pzdlqhgds));
            czmap4.put("pzdlqhgl",pzdlqzds != 0 ? df.format(pzdlqhgds / pzdlqzds * 100) : 0);
            Map czmap5 = new HashMap<>();
            czmap5.put("pzdsdlqzds",decf.format(pzdsdlqzds));
            czmap5.put("pzdsdlqhgds",decf.format(pzdsdlqhgds));
            czmap5.put("pzdsdlqhgl",pzdsdlqzds != 0 ? df.format(pzdsdlqhgds / pzdsdlqzds * 100) : 0);
            Map czmap6 = new HashMap<>();
            czmap6.put("pzdqzds",decf.format(pzdqzds));
            czmap6.put("pzdqhgds",decf.format(pzdqhgds));
            czmap6.put("pzdqhgl",pzdqzds != 0 ? df.format(pzdqhgds / pzdqzds * 100) : 0);
            resultList.add(czmap1);
            resultList.add(czmap2);
            resultList.add(czmap3);
            resultList.add(czmap4);
            resultList.add(czmap5);
            resultList.add(czmap6);
        }

        //构造深度
        List<Map<String, Object>> gzsdlist = jjgZdhGzsdJgfcService.lookJdbjg(proname);
        List<Map<String, Object>> gzsdhntlist = jjgFbgcLmgcLmgzsdsgpsfJgfcService.lookJdbjg(proname);
        double gzsdlqzds = 0,gzsdlqhgds = 0,gzsdlqqzds = 0,gzsdlqqhgds = 0,gzsdsdlqzds = 0,gzsdsdlqhgds = 0;
        if (gzsdlist != null && gzsdlist.size() > 0){
            for (Map<String, Object> stringObjectMap : gzsdlist) {
                String lmlx = stringObjectMap.get("路面类型").toString();
                if (lmlx.contains("桥")){
                    gzsdsdlqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    gzsdsdlqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }else if (lmlx.contains("隧道")){
                    gzsdlqqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    gzsdlqqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }else if (lmlx.contains("路面")){
                    gzsdlqzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    gzsdlqhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }
            }
            Map czmap1 = new HashMap<>();
            czmap1.put("gzsdlqzds",decf.format(gzsdlqzds));
            czmap1.put("gzsdlqhgds",decf.format(gzsdlqhgds));
            czmap1.put("gzsdlqhgl",gzsdlqzds != 0 ? df.format(gzsdlqhgds / gzsdlqzds * 100) : 0);
            Map czmap2 = new HashMap<>();
            czmap2.put("gzsdlqqzds",decf.format(gzsdlqqzds));
            czmap2.put("gzsdlqqhgds",decf.format(gzsdlqqhgds));
            czmap2.put("gzsdlqqhgl",gzsdlqqzds != 0 ? df.format(gzsdlqqhgds / gzsdlqqzds * 100) : 0);
            Map czmap3 = new HashMap<>();
            czmap3.put("gzsdsdlqzds",decf.format(gzsdsdlqzds));
            czmap3.put("gzsdsdlqhgds",decf.format(gzsdsdlqhgds));
            czmap3.put("gzsdsdlqhgl",gzsdsdlqzds != 0 ? df.format(gzsdsdlqhgds / gzsdsdlqzds * 100) : 0);
            resultList.add(czmap1);
            resultList.add(czmap2);
            resultList.add(czmap3);

        }
        double gzsdhntzds = 0,gzsdhnthgds = 0;
        if (gzsdhntlist != null && gzsdhntlist.size() > 0){
            for (Map<String, Object> stringObjectMap : gzsdhntlist) {
                gzsdhntzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                gzsdhnthgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
            }
            Map czmap1 = new HashMap<>();
            czmap1.put("gzsdhntzds",decf.format(gzsdhntzds));
            czmap1.put("gzsdhnthgds",decf.format(gzsdhnthgds));
            czmap1.put("gzsdhnthgl",gzsdhntzds != 0 ? df.format(gzsdhnthgds / gzsdhntzds * 100) : 0);
            Map czmap2 = new HashMap<>();
            czmap2.put("gzsdhntqzds",0);
            czmap2.put("gzsdhntqhgds",0);
            czmap2.put("gzsdhntqhgl",0);
            resultList.add(czmap1);
            resultList.add(czmap2);
        }

        //摩擦系数
        List<Map<String, Object>> mcxslist = jjgZdhMcxsJgfcService.lookJdbjg(proname);
        if (mcxslist != null && mcxslist.size() > 0){
            double mcsxlmzds = 0,mcsxlmhgds = 0,mcsxqmzds = 0,mcsxqmhgds = 0,mcsxsmzds = 0,mcsxsmhgds = 0;
            for (Map<String, Object> stringObjectMap : mcxslist) {
                String lmlx = stringObjectMap.get("路面类型").toString();
                if (lmlx.contains("桥")){
                    mcsxqmzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    mcsxqmhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }else if (lmlx.contains("隧道")){
                    mcsxsmzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    mcsxsmhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }else if (lmlx.contains("路面")){
                    mcsxlmzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                    mcsxlmhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                }
            }
            Map czmap1 = new HashMap<>();
            czmap1.put("mcsxlmzds",decf.format(mcsxlmzds));
            czmap1.put("mcsxlmhgds",decf.format(mcsxlmhgds));
            czmap1.put("mcsxlmhgl",mcsxlmzds != 0 ? df.format(mcsxlmhgds / mcsxlmzds * 100) : 0);
            Map czmap2 = new HashMap<>();
            czmap2.put("mcsxsmzds",decf.format(mcsxsmzds));
            czmap2.put("mcsxsmhgds",decf.format(mcsxsmhgds));
            czmap2.put("mcsxsmhgl",mcsxsmzds != 0 ? df.format(mcsxsmhgds / mcsxsmzds * 100) : 0);
            Map czmap3 = new HashMap<>();
            czmap3.put("mcsxqmzds",decf.format(mcsxqmzds));
            czmap3.put("mcsxqmhgds",decf.format(mcsxqmhgds));
            czmap3.put("mcsxqmhgl",mcsxqmzds != 0 ? df.format(mcsxqmhgds / mcsxqmzds * 100) : 0);
            resultList.add(czmap1);
            resultList.add(czmap2);
            resultList.add(czmap3);

        }

        //砼路面相邻板高差
        List<Map<String, Object>> xlbgclist = jjgFbgcLmgcTlmxlbgcJgfcService.lookJdbjg(proname);
        if (xlbgclist != null && xlbgclist.size() > 0){
            double tlmxlbgszds = 0,tlmxlbgshgds = 0;
            for (Map<String, Object> stringObjectMap : xlbgclist) {
                tlmxlbgszds += Double.valueOf(stringObjectMap.get("总点数").toString());
                tlmxlbgshgds += Double.valueOf(stringObjectMap.get("合格点数").toString());

            }
            Map czmap = new HashMap<>();
            czmap.put("tlmxlbgszds",decf.format(tlmxlbgszds));
            czmap.put("tlmxlbgshgds",decf.format(tlmxlbgshgds));
            czmap.put("tlmxlbgshgl",tlmxlbgszds != 0 ? df.format(tlmxlbgshgds / tlmxlbgszds * 100) : 0);
            resultList.add(czmap);
        }
        return resultList;


    }

    @Autowired
    private JjgRyinfoJgfcService jjgRyinfoJgfcService;

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTableryData2(File f, Document xw, String proname) {
        QueryWrapper<JjgRyinfoJgfc> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgRyinfoJgfc> list = jjgRyinfoJgfcService.list(wrapper);

        if (list != null){
            String[] header = {"序号", "姓名", "本项目承担职务", "技术职称", "执业资格证书编号","备注"};
            String s = "{表2.2-1}";
            TextSelection textSelection = findStringInPages(xw,s,12,14);
            // 检查是否找到字符串
            if (textSelection != null) {
                com.spire.doc.Table table=new Table(xw,true);
                table.resetCells(list.size()+1, header.length);
                Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                Body body=paragraph.ownerTextBody();
                int index=body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                float desiredWidth = 74;
                float desiredWidthInTwips = desiredWidth * 20;
                for (int i = 0; i < header.length; i++) {
                    if (i == 3){
                        table.getRows().get(0).getCells().get(i).setWidth(1292);
                    }else if (i == 4 || i == 5){
                        table.getRows().get(0).getCells().get(i).setWidth(1420);
                    } else {
                        table.getRows().get(0).getCells().get(i).setWidth(desiredWidthInTwips);
                    }
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < list.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    JjgRyinfoJgfc rowData = list.get(rowIdx);


                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(String.valueOf(rowIdx+1));
                                break;
                            case 1:
                                p.appendText(rowData.getName());
                                break;
                            case 2:
                                p.appendText(rowData.getZw());
                                break;
                            case 3:
                                p.appendText(rowData.getJszc());
                                break;
                            case 4:
                                p.appendText(rowData.getZgzsbh());
                                break;
                            case 5:
                                p.appendText(rowData.getBz());
                                break;
                            default:
                                break;
                        }

                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index,table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }

        }
    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTableryData(File f, Document xw, String proname) {
        QueryWrapper<JjgRyinfoJgfc> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgRyinfoJgfc> list = jjgRyinfoJgfcService.list(wrapper);
        JjgRyinfoJgfc jgfc1 = new JjgRyinfoJgfc();
        jgfc1.setZw("报告编写人");
        JjgRyinfoJgfc jgfc2 = new JjgRyinfoJgfc();
        jgfc2.setZw("报告审核人");
        JjgRyinfoJgfc jgfc3 = new JjgRyinfoJgfc();
        jgfc3.setZw("报告签发人");
        list.add(jgfc1);
        list.add(jgfc2);
        list.add(jgfc3);

        if (list != null){
            String[] header = {"岗 位", "姓 名", "执业资格证书编号", "职 称", "签  字"};
            String s = "{表1 人员信息}";
            //TextSelection textSelection = xw.findString(s, true, true);
            TextSelection textSelection = findStringInPages(xw,s,1,5);
            // 检查是否找到字符串
            if (textSelection != null) {
                com.spire.doc.Table table=new Table(xw,true);
                table.resetCells(list.size()+1, header.length);
                Paragraph paragraph=textSelection.getAsOneRange().getOwnerParagraph();
                Body body=paragraph.ownerTextBody();
                int index=body.getChildObjects().indexOf(paragraph);
                //将第一行设置为表格标题
                TableRow row = table.getRows().get(0);
                row.isHeader(true);
                row.setHeight(50);
                row.setHeightType(TableRowHeightType.Exactly);
                float desiredWidth = 74;
                float desiredWidthInTwips = desiredWidth * 20;
                for (int i = 0; i < header.length; i++) {
                    //row.getCells().get(i).setWidth(400);

                    if (i == 3){
                        table.getRows().get(0).getCells().get(i).setWidth(1292);
                    }else if (i == 4){
                        table.getRows().get(0).getCells().get(i).setWidth(1420);
                    } else {
                        table.getRows().get(0).getCells().get(i).setWidth(desiredWidthInTwips);
                    }
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(com.spire.doc.documents.HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                    txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < list.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);
                    row1.setHeight(34);
                    JjgRyinfoJgfc rowData = list.get(rowIdx);


                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                String zw = rowData.getZw();
                                //TextRange txtRange3 = p.appendText(rowData.getZw());
                                if ("项目负责人".equals(zw)){
                                    TextRange txtRange3 = p.appendText(rowData.getZw());
                                    txtRange3.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                    txtRange3.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                    txtRange3.getCharacterFormat().setBold(true);
                                }else if ("报告编写人".equals(zw)){
                                    TextRange txtRange3 = p.appendText("报告编写人");
                                    txtRange3.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                    txtRange3.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                    txtRange3.getCharacterFormat().setBold(true);
                                }else if ("报告审核人".equals(zw)){
                                    TextRange txtRange3 = p.appendText("报告审核人");
                                    txtRange3.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                    txtRange3.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                    txtRange3.getCharacterFormat().setBold(true);
                                }else if ("报告签发人".equals(zw)){
                                    TextRange txtRange3 = p.appendText("报告签发人");
                                    txtRange3.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                    txtRange3.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                    txtRange3.getCharacterFormat().setBold(true);
                                } else {
                                    //p.appendText("项目主要参加人员");
                                    TextRange txtRange3 = p.appendText("项目主要参加人员");
                                    txtRange3.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                    txtRange3.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                    txtRange3.getCharacterFormat().setBold(true);
                                }

                                break;
                            case 1:
                                TextRange txtRange = p.appendText(rowData.getName());
                                txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                txtRange.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                txtRange.getCharacterFormat().setBold(true);
                                break;
                            case 2:
                                TextRange txtRange1 = p.appendText(rowData.getZgzsbh());
                                txtRange1.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                txtRange1.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                txtRange1.getCharacterFormat().setBold(true);
                                break;
                            case 3:
                                TextRange txtRange2 = p.appendText(rowData.getJszc());
                                txtRange2.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                                txtRange2.getCharacterFormat().setFontSize(12f); // 设置字号为小四（对应12磅）
                                txtRange2.getCharacterFormat().setBold(true);
                                break;
                            default:
                                break;
                        }

                    }
                }
                body.getChildObjects().remove(paragraph);
                body.getChildObjects().insert(index,table);
                //列宽自动适应内容
                table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

            }

        }
    }

    /**
     *
     * @param xw
     * @param searchString
     * @param startPage
     * @param endPage
     * @return
     */
    public TextSelection findStringInPages(Document xw, String searchString, int startPage, int endPage) {
        for (int currentPage = startPage; currentPage <= endPage; currentPage++) {
            // 在当前页中查找字符串
            TextSelection textSelection = xw.findString(searchString, true, true);
            // 如果找到字符串，返回结果
            if (textSelection != null) {
                return textSelection;
            }
        }
        // 如果在指定范围内未找到字符串，返回null或适当的标识
        return null;
    }

    /**
     * 旧规范
     * @param htd
     * @param data
     * @param proname
     * @throws IOException
     */
    private void writeExceldataOld(String htd, List<Map<String, Object>> data, String proname) throws IOException  {
        //每个合同段中的
        XSSFWorkbook wb = null;
        File f = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        File fdir = new File(jgfilepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        //try {


            /*File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "评定表(旧).xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));*/
            /*InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);*/
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/评定表(旧).xlsx");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            Map<String, List<Map<String,Object>>> groupedData = data.stream()
                    .collect(Collectors.groupingBy(m -> m.get("dwgc").toString()));
            // 对key进行排序
            List<String> sortedKeys = groupedData.keySet().stream()
                    .sorted(new Comparator<String>() {
                        @Override
                        public int compare(String o1, String o2) {
                            // 判断包含关键字的顺序
                            List<String> keywords = Arrays.asList("路面工程", "交安工程","路基工程", "桥", "隧道");
                            for (String keyword : keywords) {
                                if (o1.contains(keyword) && !o2.contains(keyword)) {
                                    return -1; // o1包含关键字，o2不包含，则o1排在前面
                                } else if (!o1.contains(keyword) && o2.contains(keyword)) {
                                    return 1; // o1不包含关键字，o2包含，则o2排在前面
                                }
                            }
                            // 如果没有关键字，按字典序排列
                            return o1.compareTo(o2);
                        }
                    })
                    .collect(Collectors.toList());

            // 按关键字排序后的数据
            Map<String, List<Map<String, Object>>> sortedData = new LinkedHashMap<>();

            // 按排序后的key遍历原始数据
            for (String key : sortedKeys) {
                List<Map<String, Object>> dataList = groupedData.get(key);
                sortedData.put(key, dataList);
            }
            for (String key : sortedData.keySet()) {
                List<Map<String,Object>> list = groupedData.get(key);

                if (key.equals("路基工程")){
                    writeDataOld(wb,list,"分部-路基",proname,htd);
                }else if (key.equals("路面工程")){

                    writeDataOld(wb,list,"分部-路面",proname,htd);
                }else if (key.equals("交安工程")){
                    writeDataOld(wb,list,"分部-交安",proname,htd);
                }else {
                    //桥梁和隧道
                    writeDataOld(wb,list,"分部-"+key,proname,htd);
                }

            }

            //单位工程
            List<Map<String, Object>> dwgclist = new ArrayList<>();
            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.startsWith("分部-")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheetold(sheet);
                    dwgclist.addAll(list);
                }
            }
            writedwgcDataold(wb,dwgclist);

            //合同段
            List<Map<String, Object>> list = processhtdSheetold(proname,wb);
            //查询内页资料扣分
            QueryWrapper<JjgNyzlkf> queryWrapper = new QueryWrapper<>();
            queryWrapper.eq("proname",proname);
            queryWrapper.eq("htd",htd);
            List<JjgNyzlkf> list1 = jjgNyzlkfService.list(queryWrapper);
            String kf = "";
            if (list1!=null && list1.size()>0){
                kf = list1.get(0).getKf();
            }else {
                kf = "0";

            }
            writedhtdDataold(wb,list,kf);

            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();
        /*}catch (Exception e) {
            if(f.exists()){
                f.delete();
            }
            throw new JjgysException(20001, "生成评定表错误，请检查数据的正确性");
        }*/


    }

    private void writedhtdDataold(XSSFWorkbook wb, List<Map<String, Object>> list, String kf) {
        XSSFSheet sheet = wb.getSheet("合同段");
        createdHtdTableold(wb,getNum(list.size()));
        int index = 0;
        int tableNum = 0;
        List<Map<String, Object>> addtzedata = handtzedata(list);
        fillTitleHtdDataold(sheet, tableNum, addtzedata.get(0));
        for (Map<String, Object> datum : addtzedata) {
            if (index > 20){
                tableNum++;
                index = 0;
            }
            fillCommonHtdDataold(sheet,tableNum,index, datum);
            index++;
        }
        calculateHtdSheetold(sheet,kf);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }
    }

    private void calculateHtdSheetold(XSSFSheet sheet, String kf) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合计".equals(row.getCell(0).toString())) {

                rowstart = sheet.getRow(i-20);
                rowend = sheet.getRow(i-1);
                row.getCell(4).setCellFormula("SUM("+rowstart.getCell(4).getReference()+":"+rowend.getCell(4).getReference()+")");
                row.getCell(5).setCellFormula("SUM("+rowstart.getCell(5).getReference()+":"+rowend.getCell(5).getReference()+")");
                //合同段实测得分
                sheet.getRow(i+1).getCell(4).setCellFormula(row.getCell(5).getReference()+"/"+row.getCell(4).getReference());
                //内业资料扣分
                sheet.getRow(i+1).getCell(6).setCellValue(Double.valueOf(kf));

                //合同段鉴定得分
                sheet.getRow(i+2).getCell(4).setCellFormula( sheet.getRow(i+1).getCell(4).getReference()+"-"+sheet.getRow(i+1).getCell(6).getReference());
                //质量等级
                sheet.getRow(i+2).getCell(6).setCellFormula("IF("+sheet.getRow(i+2).getCell(4).getReference()+">=75,\"合格\",\"不合格\")");

            }
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonHtdDataold(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(index + 4).getCell(0).setCellValue(index+1);
        sheet.getRow(index + 4).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(index + 4).getCell(3).setCellValue(Double.valueOf(datum.get("df").toString()));
        sheet.getRow(index + 4).getCell(4).setCellValue(Double.valueOf(datum.get("tze").toString()));

        sheet.getRow(index + 4).getCell(5).setCellFormula(sheet.getRow(index + 4).getCell(3).getReference()+"*"+sheet.getRow(index + 4).getCell(4).getReference());
        sheet.getRow(index + 4).getCell(6).
                setCellFormula("IF("+sheet.getRow(index + 4).getCell(3).getReference()+">75,\"合格\",\"不合格\")");

    }

    /**
     * 给数据加入投资额
     * @param list
     * @return
     */
    private List<Map<String, Object>> handtzedata(List<Map<String, Object>> list) {
        for (Map<String, Object> map : list) {
            String proname = map.get("proname").toString();
            String htd = map.get("htd").toString();
            String dwgc = map.get("dwgc").toString();
            QueryWrapper<JjgDwgctze> queryWrapper =new QueryWrapper<>();
            queryWrapper.eq("proname",proname);
            queryWrapper.eq("htd",htd);
            queryWrapper.eq("name",dwgc);
            List<JjgDwgctze> list1 = jjgDwgctzeService.list(queryWrapper);
            if (list1!=null && list1.size()>0){
                map.put("tze",list1.get(0).getMoney());
            }else {
                map.put("tze","0");
            }
        }
        return list;
    }

    /**
     * 从分部工程的工作簿取数据
     * @param proname
     * @param wb
     * @return
     */
    private List<Map<String,Object>> processhtdSheetold(String proname, XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("单位工程");
        List<Map<String,Object>> list = new ArrayList<>();
        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("单位工程名称：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("proname",proname);
                map.put("dwgc",row.getCell(1).getStringCellValue());
                map.put("jsxm",row.getCell(3).getStringCellValue());
                map.put("htd",sheet.getRow(i+6).getCell(0).getStringCellValue());
                map.put("sgdw",sheet.getRow(i+2).getCell(1).getStringCellValue());
                map.put("jldw",sheet.getRow(i+2).getCell(3).getStringCellValue());
                map.put("df",sheet.getRow(i+25).getCell(1).getNumericCellValue());
                list.add(map);
            }
        }
        return list;

    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     * @param proname
     * @param htd
     */
    private void writeDataOld(XSSFWorkbook wb, List<Map<String,Object>> data, String sheetname, String proname, String htd) {
        /*Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> obj1, Map<String, Object> obj2) {
                String fbgc1 = (String) obj1.get("fbgc");
                String fbgc2 = (String) obj2.get("fbgc");
                return fbgc1.compareTo(fbgc2);
            }
        });*/
        Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> obj1, Map<String, Object> obj2) {
                String fbgc1 = (String) obj1.get("fbgc");
                String fbgc2 = (String) obj2.get("fbgc");
                int result = fbgc1.compareTo(fbgc2);
                if (result != 0) {
                    return result;
                }
                String ccxm1 = (String) obj1.get("ccxm");
                String ccxm2 = (String) obj2.get("ccxm");
                return ccxm1.compareTo(ccxm2);
            }
        });
        copySheet(wb,sheetname);
        XSSFPrintSetup ps = wb.getSheet(sheetname).getPrintSetup();
        ps.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)

        QueryWrapper<JjgJgHtdinfo> wrapperhtd = new QueryWrapper<>();
        wrapperhtd.like("proname", proname);
        wrapperhtd.like("name", htd);
        JjgJgHtdinfo htdlist = jjgJgHtdinfoMapper.selectOne(wrapperhtd);

        XSSFSheet sheet = wb.getSheet(sheetname);
        createTable(wb,gettableNumold(data),sheetname);

        int index = 0;
        int tableNum = 0;
        int startRow = -1, endRow = -1, startCol = -1, endCol = -1, startCols = -1, endCols = -1, startColhgl = -1, endColhgl = -1, startColzl = -1, endColzl = -1, startColjq = -1, endColjq = -1;
        List<Map<String, Object>> rowAndcol = new ArrayList<>();
        List<Map<String, Object>> rowAndcol1 = new ArrayList<>();
        List<Map<String, Object>> rowAndcolhgl = new ArrayList<>();
        List<Map<String, Object>> rowAndcolzl = new ArrayList<>();
        List<Map<String, Object>> rowAndcoljq = new ArrayList<>();
        String ccname = data.get(0).get("ccxm").toString();
        //处理数据，加上外观扣分
        List<Map<String,Object>> hdata = handleData(proname,data);
        Collections.sort(hdata, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> obj1, Map<String, Object> obj2) {
                String fbgc1 = (String) obj1.get("fbgc");
                String fbgc2 = (String) obj2.get("fbgc");
                return fbgc1.compareTo(fbgc2);
            }
        });
        String fbgc = data.get(0).get("fbgc").toString();
        for (Map<String,Object> datum : hdata) {
            if (index % 14 == 0 && index!=0){
                tableNum++;
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                index = 0;
            }
            if (fbgc.equals(datum.get("fbgc").toString())){
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                if (ccname.equals(datum.get("ccxm").toString())){
                    startRow = tableNum * 22 + 6 + index % 16 ;
                    endRow = tableNum * 22 + 6 + index % 16 ;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",ccname);
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 18;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 19;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);

                    startColjq = 20;
                    endColjq = 20;
                    Map<String, Object> map4 = new HashMap<>();
                    map4.put("startRow",startRow);
                    map4.put("endRow",endRow);
                    map4.put("startCol",startColjq);
                    map4.put("endCol",endColjq);
                    map4.put("name",ccname);
                    map4.put("tableNum",tableNum);
                    rowAndcoljq.add(map4);
                }else {
                    ccname = datum.get("ccxm").toString();startRow = tableNum * 22 + 6 + index % 16 ;
                    endRow = tableNum * 22 + 6 + index % 16 ;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",ccname);
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 18;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 19;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);

                    startColjq = 20;
                    endColjq = 20;
                    Map<String, Object> map4 = new HashMap<>();
                    map4.put("startRow",startRow);
                    map4.put("endRow",endRow);
                    map4.put("startCol",startColjq);
                    map4.put("endCol",endColjq);
                    map4.put("name",ccname);
                    map4.put("tableNum",tableNum);
                    rowAndcoljq.add(map4);
                }
                fillCommonDataOld(sheet,tableNum,index,datum);
                index++;
            }else {
                fbgc = datum.get("fbgc").toString();
                tableNum ++;
                index = 0;
                ccname = datum.get("ccxm").toString();
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                if (ccname.equals(datum.get("ccxm").toString())) {
                    startRow = tableNum * 22 + 6 + index % 16;
                    endRow = tableNum * 22 + 6 + index % 16;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow", startRow);
                    map.put("endRow", endRow);
                    map.put("startCol", startCol);
                    map.put("endCol", endCol);
                    map.put("name", ccname);
                    map.put("tableNum", tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 18;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 19;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);

                    startColjq = 20;
                    endColjq = 20;
                    Map<String, Object> map4 = new HashMap<>();
                    map4.put("startRow",startRow);
                    map4.put("endRow",endRow);
                    map4.put("startCol",startColjq);
                    map4.put("endCol",endColjq);
                    map4.put("name",ccname);
                    map4.put("tableNum",tableNum);
                    rowAndcoljq.add(map4);
                }
                fillCommonDataOld(sheet,tableNum,index,datum);
                index += 1;
            }
            ccname = datum.get("ccxm").toString();

        }

        List<Map<String, Object>> maps = mergeCells(rowAndcol);
        List<Map<String, Object>> mapss = mergeCells(rowAndcol1);
        List<Map<String, Object>> maphgl = mergeCellsRow(rowAndcolhgl);
        List<Map<String, Object>> mapzl = mergeCellsRow(rowAndcolzl);
        List<Map<String, Object>> mapjq = mergeCellsRow(rowAndcoljq);

        for (Map<String, Object> map : maps) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : mapss) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : maphgl) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            // 不需要合并单元格的情况
            if (map.get("startRow").equals(map.get("endRow")) && map.get("startCol").equals(map.get("endCol"))) {
                continue;
            } else {
                sheet.addMergedRegion(new CellRangeAddress(
                        Integer.parseInt(map.get("startRow").toString()),
                        Integer.parseInt(map.get("endRow").toString()),
                        Integer.parseInt(map.get("startCol").toString()),
                        Integer.parseInt(map.get("endCol").toString())
                ));
            }
            //sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }

        for (Map<String, Object> map : mapzl) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            // 不需要合并单元格的情况
            if (map.get("startRow").equals(map.get("endRow")) && map.get("startCol").equals(map.get("endCol"))) {
                continue;
            } else {
                sheet.addMergedRegion(new CellRangeAddress(
                        Integer.parseInt(map.get("startRow").toString()),
                        Integer.parseInt(map.get("endRow").toString()),
                        Integer.parseInt(map.get("startCol").toString()),
                        Integer.parseInt(map.get("endCol").toString())
                ));
            }
            //sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : mapjq) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            // 不需要合并单元格的情况
            if (map.get("startRow").equals(map.get("endRow")) && map.get("startCol").equals(map.get("endCol"))) {
                continue;
            } else {
                sheet.addMergedRegion(new CellRangeAddress(
                        Integer.parseInt(map.get("startRow").toString()),
                        Integer.parseInt(map.get("endRow").toString()),
                        Integer.parseInt(map.get("startCol").toString()),
                        Integer.parseInt(map.get("endCol").toString())
                ));
            }
            //sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        //写完当前工作簿的数据后，就要插入公式计算了
        calculateFbgcSheetOLd(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }
    }

    private List<Map<String, Object>> mergeCellsRow(List<Map<String, Object>> rowAndcol) {
        List<Map<String, Object>> result = new ArrayList<>();
        int currentEndRow = -1;
        int currentStartRow = -1;
        int currentStartCol = -1;
        int currentEndCol = -1;
        String currentNameAndCol = null;
        int currentTableNum = -1;
        for (Map<String, Object> row : rowAndcol) {
            int tableNum = (int) row.get("tableNum");
            int startRow = (int) row.get("startRow");
            int endRow = (int) row.get("endRow");
            int startCol = (int) row.get("startCol");
            int endCol = (int) row.get("endCol");
            String name = (String) row.get("name");
            if (currentNameAndCol == null || !currentNameAndCol.equals(name + "-" + startCol + "-" + endCol) || currentTableNum != tableNum) {
                if (currentStartRow != -1) {
                    for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                        Map<String, Object> newRow = new HashMap<>();
                        newRow.put("name", currentNameAndCol.split("-")[0]);
                        newRow.put("startRow", currentStartRow);
                        newRow.put("endRow", currentEndRow);
                        newRow.put("tableNum", currentTableNum);
                        for (Map.Entry<String, Object> entry : result.get(i).entrySet()) {
                            String key = entry.getKey();
                            if (!key.equals("startCol") && !key.equals("endCol")) {
                                newRow.put(key, entry.getValue());
                            }
                        }
                        result.set(i, newRow);
                    }
                }
                currentNameAndCol = name + "-" + startCol + "-" + endCol;
                currentStartCol = startCol;
                currentEndCol = endCol;
                currentTableNum = tableNum;
                currentStartRow = startRow;
                currentEndRow = endRow;
                result.add(row);
            } else {
                Map<String, Object> lastRow = result.get(result.size() - 1);
                if (currentEndCol < endCol) {
                    lastRow.put("endCol", endCol);
                    currentEndCol = endCol;
                }
                lastRow.put("endRow", endRow);
                currentEndRow = endRow;
            }
        }
        if (currentStartRow != -1) {
            for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                Map<String, Object> newRow = new HashMap<>();
                newRow.put("name", currentNameAndCol.split("-")[0]);
                newRow.put("startRow", currentStartRow);
                newRow.put("endRow", currentEndRow);
                newRow.put("tableNum", currentTableNum);
                for (Map.Entry<String, Object> entry : result.get(i).entrySet()) {
                    String key = entry.getKey();
                    if (!key.equals("startCol") && !key.equals("endCol")) {
                        newRow.put(key, entry.getValue());
                    }
                }
                result.set(i, newRow);
            }
        }

        return result;
    }

    /**
     * 合并单元格
     * @param rowAndcol
     * @return
     */
    private List<Map<String, Object>> mergeCells(List<Map<String, Object>> rowAndcol) {
        List<Map<String, Object>> result = new ArrayList<>();
        int currentEndRow = -1;
        int currentStartRow = -1;
        int currentStartCol = -1;
        int currentEndCol = -1;
        String currentName = null;
        int currentTableNum = -1;
        for (Map<String, Object> row : rowAndcol) {
            int tableNum = (int) row.get("tableNum");
            int startRow = (int) row.get("startRow");
            int endRow = (int) row.get("endRow");
            int startCol = (int) row.get("startCol");
            int endCol = (int) row.get("endCol");
            String name = (String) row.get("name");
            if (currentName == null || !currentName.equals(name) || currentStartCol != startCol || currentEndCol != endCol || currentTableNum != tableNum) {
                if (currentStartRow != -1) {
                    for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                        Map<String, Object> newRow = new HashMap<>();
                        newRow.put("name", currentName);
                        newRow.put("startRow", currentStartRow);
                        newRow.put("endRow", currentEndRow);
                        newRow.put("startCol", currentStartCol);
                        newRow.put("endCol", currentEndCol);
                        newRow.put("tableNum", currentTableNum);
                        newRow.putAll(result.get(i));
                        result.set(i, newRow);
                    }
                }
                currentName = name;
                currentStartCol = startCol;
                currentEndCol = endCol;
                currentTableNum = tableNum;
                currentStartRow = startRow;
                currentEndRow = endRow;
                result.add(row);
            } else {
                Map<String, Object> lastRow = result.get(result.size() - 1);
                lastRow.put("endRow", endRow);
                currentEndRow = endRow;
            }
        }
        if (currentStartRow != -1) {
            for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                Map<String, Object> newRow = new HashMap<>();
                newRow.put("name", currentName);
                newRow.put("startRow", currentStartRow);
                newRow.put("endRow", currentEndRow);
                newRow.put("startCol", currentStartCol);
                newRow.put("endCol", currentEndCol);
                newRow.put("tableNum", currentTableNum);
                newRow.putAll(result.get(i));
                result.set(i, newRow);
            }
        }
        return result;
    }

    /**
     *
     * @param proname
     * @param data
     * @return
     */
    private List<Map<String, Object>> handleData(String proname, List<Map<String, Object>> data) {

        for (Map<String, Object> datum : data) {
            QueryWrapper<JjgWgkf> queryWrapper  = new QueryWrapper<>();
            queryWrapper.eq("proname",proname);
            queryWrapper.eq("htd",datum.get("htd").toString());
            queryWrapper.eq("dwgc",datum.get("dwgc").toString());
            queryWrapper.eq("fbgc",datum.get("fbgc").toString());
            List<JjgWgkf> list = jjgWgkfService.list(queryWrapper);
            String fbgckf = "";
            if (list != null && list.size()>0){
                 fbgckf = list.get(0).getFbgckf();
            }else {
                fbgckf = "0";
            }
            datum.put("kf",fbgckf);
        }
        return data;

    }

    /**
     * 计算分部工程的结果
     * @param sheet
     */
    private void calculateFbgcSheetOLd(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合计".equals(row.getCell(1).toString())) {
                rowstart = sheet.getRow(i-14);
                rowend = sheet.getRow(i-1);
                row.getCell(19).setCellFormula("SUM("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+")");
                row.getCell(20).setCellFormula("SUM("+rowstart.getCell(20).getReference()+":"+rowend.getCell(20).getReference()+")");
                sheet.getRow(i+1).getCell(6).setCellFormula(row.getCell(20).getReference()+"/"+row.getCell(19).getReference());
                sheet.getRow(i+1).getCell(17).setCellFormula(sheet.getRow(i+1).getCell(6).getReference()+"-"+sheet.getRow(i+1).getCell(10).getReference());
                sheet.getRow(i+1).getCell(20).setCellFormula("IF("+sheet.getRow(i+1).getCell(17).getReference()+">=75,\"合格\",\"不合格\")");
            }
        }

    }

    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonDataOld(XSSFSheet sheet, int tableNum, int index, Map<String,Object> datum) {
        sheet.getRow(tableNum * 22 + index + 6).getCell(1).setCellValue(1+index);
        sheet.getRow(tableNum * 22 + index + 6).getCell(2).setCellValue(datum.get("ccxm").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(6).setCellValue(String.valueOf(datum.get("gdz")));

        sheet.getRow(tableNum * 22 + index + 6).getCell(7).setCellValue(datum.get("filename").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(17).setCellValue("\\");
        sheet.getRow(tableNum * 22 + index + 6).getCell(18).setCellValue(datum.get("hgl").toString());
        //权值
        sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue(Double.valueOf(datum.get("qz").toString()));
        //加权得分
        sheet.getRow(tableNum * 22 + index + 6).getCell(20).setCellFormula(sheet.getRow(tableNum * 22 + index + 6).getCell(18).getReference()+"*"+sheet.getRow(tableNum * 22 + index + 6).getCell(19).getReference());

        //扣分
        sheet.getRow(tableNum * 22+6+15).getCell(10).setCellValue(Integer.valueOf(datum.get("kf").toString()));

    }



    /**
     *
     * @param htd
     * @param data
     * @param proname
     */
    private void writeExceldata(String htd, List<Map<String,Object>> data, String proname) throws IOException {
        //每个合同段中的
        XSSFWorkbook wb = null;
        File f = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        File fdir = new File(jgfilepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        try {
            /*File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "合同段评定表.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);*/
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/合同段评定表.xlsx");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            Map<String, List<Map<String,Object>>> groupedData = data.stream()
                    .collect(Collectors.groupingBy(m -> m.get("dwgc").toString()));
            // 对key进行排序
            List<String> sortedKeys = groupedData.keySet().stream()
                    .sorted(new Comparator<String>() {
                        @Override
                        public int compare(String o1, String o2) {
                            // 判断包含关键字的顺序
                            List<String> keywords = Arrays.asList("路面工程", "交安工程","路基工程", "桥", "隧道");
                            for (String keyword : keywords) {
                                if (o1.contains(keyword) && !o2.contains(keyword)) {
                                    return -1; // o1包含关键字，o2不包含，则o1排在前面
                                } else if (!o1.contains(keyword) && o2.contains(keyword)) {
                                    return 1; // o1不包含关键字，o2包含，则o2排在前面
                                }
                            }
                            // 如果没有关键字，按字典序排列
                            return o1.compareTo(o2);
                        }
                    })
                    .collect(Collectors.toList());

            // 按关键字排序后的数据
            Map<String, List<Map<String, Object>>> sortedData = new LinkedHashMap<>();

            // 按排序后的key遍历原始数据
            for (String key : sortedKeys) {
                List<Map<String, Object>> dataList = groupedData.get(key);
                sortedData.put(key, dataList);
            }

            for (String key : sortedData.keySet()) {
                List<Map<String,Object>> list = groupedData.get(key);

                if (key.equals("路基工程")){
                    writeData(wb,list,"分部-路基",proname,htd);
                }else if (key.equals("路面工程")){

                    writeData(wb,list,"分部-路面",proname,htd);
                }else if (key.equals("交安工程")){
                    writeData(wb,list,"分部-交安",proname,htd);
                }else {
                    //桥梁和隧道
                    writeData(wb,list,key,proname,htd);
                }

            }

            //单位工程
            List<Map<String, Object>> dwgclist = new ArrayList<>();
            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.startsWith("分部-")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
            writedwgcData(wb,dwgclist);

            //合同段
            List<Map<String, Object>> list = processhtdSheet(wb);
            writedhtdData(wb,list);

            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();
        }catch (Exception e) {
            if(f.exists()){
                f.delete();
            }
            throw new JjgysException(20001, "生成评定表错误，请检查数据的正确性");
        }

    }

    /**
     *
     * @param wb
     * @param list
     */
    private void writedhtdData(XSSFWorkbook wb, List<Map<String, Object>> list) {
        XSSFSheet sheet = wb.getSheet("合同段");
        createdHtdTable(wb,getNum(list.size()));
        int index = 0;
        int tableNum = 0;
        fillTitleHtdData(sheet, tableNum, list.get(0));
        for (Map<String, Object> datum : list) {
            if (index > 20){
                tableNum++;
                index = 0;
            }
            fillCommonHtdData(sheet,tableNum,index, datum);
            index++;
        }
        calculateHtdSheet(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }
    }

    /**
     *
     * @param wb
     * @param gettableNum
     */
    private void createdHtdTable(XSSFWorkbook wb, int gettableNum) {
        int record = 0;
        record = gettableNum;
        for(int i = 1; i < record; i++){
            if(i < record-1){
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 28, (i - 1) * 24 + 29);
            }
            else{
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 26, (i - 1) * 24 + 29);
            }
        }
        XSSFSheet sheet = wb.getSheet("合同段");
        int lastRowNum = sheet.getLastRowNum();
        int numMergedRegions = sheet.getNumMergedRegions();

        for (int i = numMergedRegions - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();

            if (lastRow >= lastRowNum - 2 && lastRow <= lastRowNum) {
                sheet.removeMergedRegion(i);
            }
        }
        RowCopy.copyRows(wb, "source", "合同段", 0, 2,(record) * 24+2);
        wb.setPrintArea(wb.getSheetIndex("合同段"), 0, 7, 0, (record) * 24+4);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getNum(int size) {
        return size%24 ==0 ? size/24 : size/24+1;
    }


    /**
     *
     * @param sheet
     * @param tableNum
     * @param datum
     */
    private void fillTitleHtdData(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 29 +1).getCell(1).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 29 + 1).getCell(5).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(1).setCellValue(datum.get("sgdw").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(5).setCellValue(datum.get("jldw").toString());

    }

    private void fillTitleHtdDataold(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 29 +1).getCell(2).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 29 + 1).getCell(5).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(2).setCellValue(datum.get("sgdw").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(5).setCellValue(datum.get("jldw").toString());

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonHtdData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(index + 5).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(index + 5).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(index + 5).getCell(4).setCellValue(datum.get("zldj").toString());

    }

    /**
     *
     * @param wb
     * @param gettableNum
     */
    private void createdHtdTableold(XSSFWorkbook wb, int gettableNum) {
        int record = 0;
        record = gettableNum;
        for(int i = 1; i < record; i++){
            if(i < record-1){
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 28, (i - 1) * 24 + 29);
            }
            else{
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 26, (i - 1) * 24 + 29);
            }
        }
        XSSFSheet sheet = wb.getSheet("合同段");
        int lastRowNum = sheet.getLastRowNum();
        int numMergedRegions = sheet.getNumMergedRegions();

        for (int i = numMergedRegions - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();

            if (lastRow >= lastRowNum - 3 && lastRow <= lastRowNum) {
                sheet.removeMergedRegion(i);
            }
        }
        RowCopy.copyRows(wb, "source", "合同段", 0, 4,(record) * 23+2);
        wb.setPrintArea(wb.getSheetIndex("合同段"), 0, 7, 0, (record) * 23+5);
    }

    /**
     *
     * @param sheet
     */
    private void calculateHtdSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合同段质量等级".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-21);
                rowend = sheet.getRow(i-1);
                /*row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference() +":"+rowend.getCell(2).getReference()
                        +",\"<>合格\")=0,\"合格\", \"不合格\")");*///=IF(COUNTIF(C64:C81,"<>合格")=0, "合格", "不合格")
                row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(4).getReference()
                        +":"+rowend.getCell(4).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(4).getReference()
                        +":"+rowend.getCell(4).getReference()+"),\"合格\", \"不合格\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")

            }
        }
    }
    /**
     *
     * @param wb
     * @return
     */
    private List<Map<String,Object>> processhtdSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("单位工程");
        List<Map<String,Object>> list = new ArrayList<>();
        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("单位工程名称：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("dwgc",row.getCell(1).getStringCellValue());
                map.put("jsxm",row.getCell(3).getStringCellValue());
                map.put("htd",sheet.getRow(i+6).getCell(0).getStringCellValue());
                map.put("sgdw",sheet.getRow(i+2).getCell(1).getStringCellValue());
                map.put("jldw",sheet.getRow(i+2).getCell(3).getStringCellValue());
                map.put("zldj",sheet.getRow(i+24).getCell(1).getStringCellValue());
                list.add(map);
            }
        }
        return list;

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void writedwgcDataold(XSSFWorkbook wb, List<Map<String, Object>> data) {
        if (data.size()>0&&data!=null){
            XSSFSheet sheet = wb.getSheet("单位工程");
            createdwgcTable(wb,getdwgcNum(data));

            int index = 0;
            int tableNum = 0;
            List<Map<String, Object>> sj = hadnld(data);
            String fbgc = sj.get(0).get("dwgc").toString();
            for (Map<String, Object> datum : sj) {
                if (fbgc.equals(datum.get("dwgc"))){
                    fillTitleDwgcData(sheet,tableNum,datum);
                    fillCommonDwgcDataold(sheet,tableNum,index,datum);
                    index++;
                }else {
                    fbgc = datum.get("dwgc").toString();
                    tableNum ++;
                    index = 0;
                    fillTitleDwgcData(sheet,tableNum,datum);
                    fillCommonDwgcDataold(sheet,tableNum,index,datum);
                    index += 1;
                }
            }
            calculateDwgcSheetold(sheet);
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }
        }


    }

    /**
     *
     * @param data
     * @return
     */
    private List<Map<String, Object>> hadnld(List<Map<String, Object>> data) {
        String proname = data.get(0).get("proname").toString();
        for (Map<String, Object> datum : data) {
            //扣分
            QueryWrapper<JjgWgkf> queryWrapper  = new QueryWrapper<>();
            queryWrapper.eq("proname",proname);
            queryWrapper.eq("htd",datum.get("htd").toString());
            queryWrapper.eq("dwgc",datum.get("dwgc").toString());
            queryWrapper.eq("fbgc",datum.get("fbgc").toString());
            List<JjgWgkf> list = jjgWgkfService.list(queryWrapper);
            String fbgckf ="";
            if (list!=null && list.size()>0){
                fbgckf = list.get(0).getFbgckf();
            }else {
                fbgckf="0";
            }
            datum.put("kf",fbgckf);

            //权值
            String dwgc = datum.get("dwgc").toString();
            String fbgc = datum.get("fbgc").toString();
            if (dwgc.equals("路基工程") && fbgc.equals("路基土石方")){
                datum.put("qz","3");
            }else if(dwgc.equals("路基工程") && fbgc.equals("排水工程")){
                datum.put("qz","1");
            }else if(dwgc.equals("路基工程") && fbgc.equals("小桥")){
                datum.put("qz","2");
            }else if(dwgc.equals("路基工程") && fbgc.equals("涵洞")){
                datum.put("qz","1");
            }else if(dwgc.equals("路基工程") && fbgc.equals("支挡工程")){
                datum.put("qz","2");
            } else if(dwgc.equals("路面工程") && fbgc.equals("路面面层")){
                datum.put("qz","1");
            }else if(fbgc.equals("桥梁下部")){
                datum.put("qz","2");
            }else if(fbgc.equals("桥梁上部")){
                datum.put("qz","3");
            }else if(fbgc.equals("衬砌")){
                datum.put("qz","3");
            }else if(fbgc.equals("总体")){
                datum.put("qz","1");
            }else if(fbgc.equals("标志")){
                datum.put("qz","1");
            }else if(fbgc.equals("标线")){
                datum.put("qz","1");
            }else if(fbgc.equals("桥面系")){
                datum.put("qz","2");
            }else if(fbgc.equals("防护栏")){
                datum.put("qz","2");
            }else{
                datum.put("qz","0");
            }
        }
        return data;


    }

    /**
     *
     * @param wb
     * @param data
     */
    private void writedwgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        if (data.size()>0&&data!=null){
            XSSFSheet sheet = wb.getSheet("单位工程");
            createdwgcTable(wb,getdwgcNum(data));

            int index = 0;
            int tableNum = 0;
            String fbgc = data.get(0).get("dwgc").toString();
            for (Map<String, Object> datum : data) {
                if (fbgc.equals(datum.get("dwgc"))){
                    fillTitleDwgcData(sheet,tableNum,datum);
                    fillCommonDwgcData(sheet,tableNum,index,datum);
                    index++;
                }else {
                    fbgc = datum.get("dwgc").toString();
                    tableNum ++;
                    index = 0;
                    fillTitleDwgcData(sheet,tableNum,datum);
                    fillCommonDwgcData(sheet,tableNum,index,datum);
                    index += 1;
                }

            }
            calculateDwgcSheet(sheet);
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }
        }


    }


    /**
     *
     * @param sheet
     */
    private void calculateDwgcSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("单位工程质量等级".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-18);
                rowend = sheet.getRow(i-1);
                /*row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference() +":"+rowend.getCell(2).getReference()
                        +",\"<>合格\")=0,\"合格\", \"不合格\")");*///=IF(COUNTIF(C64:C81,"<>合格")=0, "合格", "不合格")
                row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference()
                        +":"+rowend.getCell(2).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(2).getReference()
                        +":"+rowend.getCell(2).getReference()+"),\"合格\", \"不合格\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")

            }
        }
    }

    /**
     *
     * @param sheet
     */
    private void calculateDwgcSheetold(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合计".equals(row.getCell(1).toString())) {
                rowstart = sheet.getRow(i-18);
                rowend = sheet.getRow(i-1);
                row.getCell(3).setCellFormula("SUM("+rowstart.getCell(3).getReference()+":"+rowend.getCell(3).getReference()+")");
                row.getCell(4).setCellFormula("SUM("+rowstart.getCell(4).getReference()+":"+rowend.getCell(4).getReference()+")");
                sheet.getRow(i+1).getCell(1).setCellFormula(row.getCell(4).getReference()+"/"+row.getCell(3).getReference());
                sheet.getRow(i+1).getCell(5).setCellFormula("IF("+sheet.getRow(i+1).getCell(1).getReference()+">=75,\"合格\",\"不合格\")");

            }
        }
    }


    /**
     *
     * @param sheet
     * @param tableNum
     * @param datum
     */
    private void fillTitleDwgcData(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 +1).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(tableNum * 28 +1).getCell(3).setCellValue(datum.get("proname").toString());
        sheet.getRow(tableNum * 28 +2).getCell(1).setCellValue(datum.get("proname").toString());
        sheet.getRow(tableNum * 28 +2).getCell(3).setCellValue(datum.get("gcbw").toString());
        sheet.getRow(tableNum * 28 +3).getCell(1).setCellValue(datum.get("sgbw").toString());
        sheet.getRow(tableNum * 28 +3).getCell(3).setCellValue(datum.get("jldw").toString());


    }
    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonDwgcDataold(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 + index + 7).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(1).setCellValue(datum.get("fbgc").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(2).setCellValue(Double.valueOf(datum.get("scfs").toString()));
        sheet.getRow(tableNum * 28 + index + 7).getCell(3).setCellValue(Double.valueOf(datum.get("qz").toString()));
        sheet.getRow(tableNum * 28 + index + 7).getCell(4).
                setCellFormula(sheet.getRow(tableNum * 28 + index + 7).getCell(2).getReference()+"*"+sheet.getRow(tableNum * 28 + index + 7).getCell(3).getReference());
    }


    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonDwgcData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 + index + 7).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(1).setCellValue(datum.get("fbgc").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(2).setCellValue(datum.get("sfhg").toString());

    }

    /**
     *
     * @param data
     * @return
     */
    private int getdwgcNum(List<Map<String, Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("dwgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%18==0){
                num += value/18;
            }else {
                num += value/18+1;
            }
        }
        return num;

    }

    /**
     * 单位工程
     * @param wb
     * @param gettableNum
     */
    private void createdwgcTable(XSSFWorkbook wb, int gettableNum) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, "单位工程", "单位工程", 0, 27, i*28);
        }
        if(gettableNum > 1){
            wb.setPrintArea(wb.getSheetIndex("单位工程"), 0, 5, 0, gettableNum*28-1);
        }

    }

    /**
     *
     * @param sheet
     * @return
     */
    private List<Map<String,Object>> processSheetold(Sheet sheet) {
        List<Map<String,Object>> list = new ArrayList<>();

        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("合同段：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("htd",row.getCell(2).getStringCellValue());
                map.put("fbgc",row.getCell(8).getStringCellValue());
                map.put("proname",row.getCell(15).getStringCellValue());

                map.put("gcbw",sheet.getRow(i+1).getCell(2).getStringCellValue());
                map.put("sgbw",sheet.getRow(i+1).getCell(8).getStringCellValue());
                map.put("jldw",sheet.getRow(i+1).getCell(15).getStringCellValue());

                map.put("scfs",sheet.getRow(i+20).getCell(6).getNumericCellValue());
                String s = sheet.getSheetName().substring(sheet.getSheetName().indexOf("-")+1);
                if (s.equals("路面") || s.equals("路基") || s.equals("交安"))
                {
                    map.put("dwgc",s+"工程");
                }else {
                    map.put("dwgc",s);
                }
                list.add(map);
            }
        }
        return list;
    }


    /**
     *
     * @param sheet
     * @return
     */
    private List<Map<String,Object>> processSheet(Sheet sheet) {
        List<Map<String,Object>> list = new ArrayList<>();

        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("合同段：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("htd",row.getCell(2).getStringCellValue());
                map.put("fbgc",row.getCell(8).getStringCellValue());
                map.put("jsxm",row.getCell(15).getStringCellValue());
                map.put("proname",row.getCell(15).getStringCellValue());
                map.put("gcbw",sheet.getRow(i+1).getCell(2).getStringCellValue());
                map.put("sgbw",sheet.getRow(i+1).getCell(8).getStringCellValue());
                map.put("jldw",sheet.getRow(i+1).getCell(15).getStringCellValue());
                if (sheet.getRow(i+20).getCell(16).getStringCellValue().equals("√")){
                    map.put("sfhg","合格");
                }else {
                    map.put("sfhg","不合格");
                }

                map.put("dwgc",sheet.getSheetName().substring(sheet.getSheetName().indexOf("-")+1));
                list.add(map);
            }
        }
        return list;
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     * @param proname
     * @param htd
     */
    private void writeData(XSSFWorkbook wb, List<Map<String,Object>> data, String sheetname, String proname, String htd) {
        /*Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> obj1, Map<String, Object> obj2) {
                String fbgc1 = (String) obj1.get("fbgc");
                String fbgc2 = (String) obj2.get("fbgc");
                return fbgc1.compareTo(fbgc2);
            }
        });*/
        Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> obj1, Map<String, Object> obj2) {
                String fbgc1 = (String) obj1.get("fbgc");
                String fbgc2 = (String) obj2.get("fbgc");
                int result = fbgc1.compareTo(fbgc2);
                if (result != 0) {
                    return result;
                }
                String ccxm1 = (String) obj1.get("ccxm");
                String ccxm2 = (String) obj2.get("ccxm");
                return ccxm1.compareTo(ccxm2);
            }
        });
        copySheet(wb,sheetname);
        XSSFPrintSetup ps = wb.getSheet(sheetname).getPrintSetup();
        ps.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)

        QueryWrapper<JjgJgHtdinfo> wrapperhtd = new QueryWrapper<>();
        wrapperhtd.like("proname", proname);
        wrapperhtd.like("name", htd);
        JjgJgHtdinfo htdlist = jjgJgHtdinfoMapper.selectOne(wrapperhtd);

        XSSFSheet sheet = wb.getSheet(sheetname);
        createTable(wb,gettableNum(data),sheetname);

        int index = 0;
        int tableNum = 0;
        int startRow = -1, endRow = -1, startCol = -1, endCol = -1, startCols = -1, endCols = -1, startColhgl = -1, endColhgl = -1, startColzl = -1, endColzl = -1;
        List<Map<String, Object>> rowAndcol = new ArrayList<>();
        List<Map<String, Object>> rowAndcol1 = new ArrayList<>();
        List<Map<String, Object>> rowAndcolhgl = new ArrayList<>();
        List<Map<String, Object>> rowAndcolzl = new ArrayList<>();
        String ccname = data.get(0).get("ccxm").toString();
        String fbgc = data.get(0).get("fbgc").toString();
        for (Map<String,Object> datum : data) {
            if (index % 15 == 0 && index!=0){
                tableNum++;
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                index = 0;
            }
            if (fbgc.equals(datum.get("fbgc").toString())){
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                if (ccname.equals(datum.get("ccxm").toString())){
                    startRow = tableNum * 22 + 6 + index % 16 ;
                    endRow = tableNum * 22 + 6 + index % 16 ;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",ccname);
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 17;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 20;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);
                }else {
                    ccname = datum.get("ccxm").toString();startRow = tableNum * 22 + 6 + index % 16 ;
                    endRow = tableNum * 22 + 6 + index % 16 ;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",ccname);
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 17;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 20;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);
                }

                fillCommonData(sheet,tableNum,index,datum);
                index++;
            }else {
                fbgc = datum.get("fbgc").toString();
                tableNum ++;
                index = 0;
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                ccname = datum.get("ccxm").toString();
                if (ccname.equals(datum.get("ccxm").toString())) {
                    startRow = tableNum * 22 + 6 + index % 16;
                    endRow = tableNum * 22 + 6 + index % 16;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow", startRow);
                    map.put("endRow", endRow);
                    map.put("startCol", startCol);
                    map.put("endCol", endCol);
                    map.put("name", ccname);
                    map.put("tableNum", tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 17;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 20;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);
                }
                fillCommonData(sheet,tableNum,index,datum);
                index += 1;
            }
            ccname = datum.get("ccxm").toString();

        }
        List<Map<String, Object>> maps = mergeCells(rowAndcol);
        List<Map<String, Object>> mapss = mergeCells(rowAndcol1);
        List<Map<String, Object>> maphgl = mergeCells(rowAndcolhgl);
        List<Map<String, Object>> mapzl = mergeCells(rowAndcolzl);

        for (Map<String, Object> map : maps) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : mapss) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : maphgl) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }

        for (Map<String, Object> map : mapzl) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        //写完当前工作簿的数据后，就要插入公式计算了
        calculateFbgcSheet(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }

    }

    /**
     * 计算分部工程的结果
     * @param sheet
     */
    private void calculateFbgcSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("实测项目是否全部合格".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-15);
                rowend = sheet.getRow(i-1);
                row.getCell(8).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"\",\"√\")");//=IF(COUNTIF(T29:U43,"不合格")>0,"","√")
                row.getCell(10).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"×\",\"\")");//=IF(COUNTIF(T29:U43,"不合格")>0,"","√")

                //row.getCell(8).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"√\", \"\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")
                //row.getCell(10).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"\", \"×\")");//=IF(COUNTIF(T7:T21, "不合格") = COUNTA(T7:T21), "", "×")

                row.getCell(16).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"\",\"√\")");//=IF(COUNTIF(T29:U43,"不合格")>0,"","√")
                row.getCell(19).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"×\",\"\")");

                //row.getCell(16).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"√\", \"\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")
                //row.getCell(19).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"\", \"×\")");//=IF(COUNTIF(T7:T21, "不合格") = COUNTA(T7:T21), "", "×")
            }
        }

    }

    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonData(XSSFSheet sheet, int tableNum, int index, Map<String,Object> datum) {
        sheet.getRow(tableNum * 22 + index + 6).getCell(1).setCellValue(1+index);
        sheet.getRow(tableNum * 22 + index + 6).getCell(2).setCellValue(datum.get("ccxm").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(6).setCellValue(String.valueOf(datum.get("gdz")));

        sheet.getRow(tableNum * 22 + index + 6).getCell(7).setCellValue(datum.get("filename").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(17).setCellValue(datum.get("hgl").toString());
        if (datum.get("ccxm").toString().contains("*")){
            Double value = Double.valueOf(datum.get("hgl").toString());
            if (value == 100){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }else if (datum.get("ccxm").toString().contains("△")){
            Double value = Double.valueOf(datum.get("hgl").toString());
            if (value >= 95){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }else {
            Double value = Double.valueOf(datum.get("hgl").toString());
            if (value >= 80){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }
    }

    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param htdlist
     * @param fbgc
     */
    private void fillTitleData(XSSFSheet sheet, int tableNum, String proname, String htd,JjgJgHtdinfo htdlist,String fbgc){
        sheet.getRow(tableNum * 22 +1).getCell(2).setCellValue(htd);
        sheet.getRow(tableNum * 22 +1).getCell(8).setCellValue(fbgc);
        sheet.getRow(tableNum * 22 +2).getCell(2).setCellValue(getgcbw(htdlist.getZhq(),htdlist.getZhz()));
        sheet.getRow(tableNum * 22 +2).getCell(8).setCellValue(htdlist.getSgdw());
        sheet.getRow(tableNum * 22 +2).getCell(15).setCellValue(htdlist.getJldw());
        sheet.getRow(tableNum * 22 +1).getCell(15).setCellValue(proname);
    }

    /**
     *分部
     * @param zhq
     * @param zhz
     * @return
     */
    private String getgcbw(String zhq, String zhz) {
        DecimalFormat decimalFormat = new DecimalFormat("#.###");
        double a = Double.valueOf(zhq);
        double b = Double.valueOf(zhz);

        int aa = (int) a / 1000;
        int bb = (int) b / 1000;
        double cc = a % 1000;
        double dd = b % 1000;
        String result = "K"+aa+"+"+decimalFormat.format(cc)+"--"+"K"+bb+"+"+decimalFormat.format(dd);

        return result;
    }

    /**
     * 分部
     * @param wb
     * @param gettableNum
     * @param sheetname
     */
    private void createTable(XSSFWorkbook wb,int gettableNum,String sheetname) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 21, i*22);
        }
        if(gettableNum > 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 20, 0, gettableNum*22-1);
        }
    }



    /**
     * 分部
     * @param data
     * @return
     */
    private int gettableNum(List<Map<String,Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String,Object> map : data) {
            String name = map.get("fbgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%15==0){
                num += value/15;
            }else {
                num += value/15+1;
            }
        }
        return num;
    }

    /**
     * 分部
     * @param data
     * @return
     */
    private int gettableNumold(List<Map<String,Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String,Object> map : data) {
            String name = map.get("fbgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%14==0){
                num += value/14;
            }else {
                num += value/14+1;
            }
        }
        return num;
    }

    /**
     *
     * @param wb
     * @param sheetname
     */
    private void copySheet(XSSFWorkbook wb,String sheetname) {
        String sourceSheetName = "模板";
        String targetSheetName = sheetname;
        XSSFSheet sourceSheet = wb.getSheet(sourceSheetName);

        // 创建新的工作表，并将源工作表中的内容复制到新工作表
        XSSFSheet targetSheet = wb.createSheet(targetSheetName);
        copySheetInfo(wb,sourceSheet, targetSheet);


    }

    /**
     *
     * @param wb
     * @param sourceSheet
     * @param targetSheet
     */
    private static void copySheetInfo(XSSFWorkbook wb, Sheet sourceSheet, Sheet targetSheet) {
        int maxColumnNum = 0;
        for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row newRow = targetSheet.createRow(i);
            if (sourceRow != null) {
                copyRow(sourceRow, newRow, wb);
                if (sourceRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = sourceRow.getLastCellNum();
                }
            }
        }

        // 复制列宽
        for (int j = 0; j < maxColumnNum; j++) {
            targetSheet.setColumnWidth(j, sourceSheet.getColumnWidth(j));
        }

        // 复制合并单元格
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            targetSheet.addMergedRegion(mergedRegion);
        }

        // 复制打印区域
        PrintSetup sourcePrintSetup = sourceSheet.getPrintSetup();
        PrintSetup targetPrintSetup = targetSheet.getPrintSetup();
        targetPrintSetup.setLandscape(sourcePrintSetup.getLandscape());
        targetPrintSetup.setPaperSize(sourcePrintSetup.getPaperSize());
        targetPrintSetup.setFitWidth(sourcePrintSetup.getFitWidth());
        targetPrintSetup.setFitHeight(sourcePrintSetup.getFitHeight());
        targetPrintSetup.setScale(sourcePrintSetup.getScale());

        // 复制页眉
        Header sourceHeader = sourceSheet.getHeader();
        Header targetHeader = targetSheet.getHeader();
        targetHeader.setCenter(sourceHeader.getCenter());
        targetHeader.setLeft(sourceHeader.getLeft());
        targetHeader.setRight(sourceHeader.getRight());

        // 复制页脚
        Footer sourceFooter = sourceSheet.getFooter();
        Footer targetFooter = targetSheet.getFooter();
        targetFooter.setCenter(sourceFooter.getCenter());
        targetFooter.setLeft(sourceFooter.getLeft());
        targetFooter.setRight(sourceFooter.getRight());


    }

    /**
     *
     * @param sourceRow
     * @param newRow
     * @param wb
     */
    private static void copyRow(Row sourceRow, Row newRow, Workbook wb) {
        newRow.setHeight(sourceRow.getHeight());
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);
            if (sourceCell != null) {
                CellStyle sourceCellStyle = sourceCell.getCellStyle();
                CellStyle newCellStyle = wb.createCellStyle();
                newCellStyle.cloneStyleFrom(sourceCellStyle);
                newCell.setCellStyle(newCellStyle);

                CellType cellType = sourceCell.getCellTypeEnum();

                switch (cellType) {
                    case BOOLEAN:
                        newCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case STRING:
                        newCell.setCellValue(sourceCell.getRichStringCellValue());
                        break;
                    case NUMERIC:
                        newCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case FORMULA:
                        newCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    case BLANK:
                        // do nothing
                        break;
                    default:
                        // do nothing
                        break;
                }
            }
        }
    }

    /**
     *
     * @param jjgJgjcsjs
     * @return
     */
    private List<Map<String, Object>> getdata(List<JjgJgjcsj> jjgJgjcsjs) {
        DecimalFormat df = new DecimalFormat("0.00");
        List<Map<String, Object>> result = new ArrayList<>();
        for (JjgJgjcsj data : jjgJgjcsjs) {
            String ccxm = data.getCcxm();
   
            String hgds = data.getHgds();
            String zds = data.getZds();
            String fbgc = data.getFbgc();

            Map<String, Object> map = new HashMap<>();
            map.put("htd",data.getHtd());
            map.put("dwgc",data.getDwgc());
            map.put("fbgc",fbgc);
            map.put("ccxm",ccxm);
            map.put("gdz",data.getGdz());

            if (fbgc.equals("路基土石方") && ccxm.equals("△沉降")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《路基压实度沉降质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路基土石方") && ccxm.equals("△压实度（沙砾）")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《路基压实度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路基土石方") && ccxm.equals("△压实度（灰土）")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《路基压实度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路基土石方") && ccxm.equals("△弯沉")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《路基弯沉质量鉴定表》检测"+zds+"个评定单元,合格"+hgds+"个评定单元");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路基土石方") && ccxm.equals("边坡")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《路基边坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("排水工程") && ccxm.equals("断面尺寸")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《结构（断面）尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("排水工程") && ccxm.equals("铺砌厚度")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《排水铺砌厚度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("小桥") && ccxm.equals("*混凝土强度")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《小桥砼强度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("小桥") && ccxm.equals("△断面尺寸")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《小桥结构尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("涵洞") && ccxm.equals("*混凝土强度")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《涵洞砼强度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("涵洞") && ccxm.equals("结构尺寸")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《涵洞结构尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("支挡工程") && ccxm.equals("*混凝土强度")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《支挡工程砼强度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("支挡工程") && ccxm.equals("△断面尺寸")){
                map.put("sheetname","分部-路基");
                map.put("filename","详见《支挡工程结构尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }
            //桥梁和隧道
            else if (fbgc.equals("桥梁上部") && ccxm.equals("*混凝土强度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁上部墩台砼强度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥梁上部") && ccxm.equals("主要结构尺寸")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁上部主要结构尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥梁上部") && ccxm.equals("钢筋保护层厚度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁上部钢筋保护层厚度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥梁下部") && ccxm.equals("*混凝土强度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁下部墩台砼强度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥梁下部") && ccxm.equals("主要结构尺寸")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁下部主要结构尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥梁下部") && ccxm.equals("钢筋保护层厚度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁下部钢筋保护层厚度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥梁下部") && ccxm.equals("*竖直度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥梁下部墩台垂直度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }

            else if (fbgc.equals("桥面系") && ccxm.equals("沥青桥平整度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥面系平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥面系") && ccxm.equals("混凝土桥平整度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥面系平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥面系") && ccxm.equals("沥青桥面构造深度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥面系构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("桥面系") && ccxm.equals("混凝土桥面构造深度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土桥面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥面系") && ccxm.equals("摩擦系数")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《桥面系摩擦系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥面系") && ccxm.equals("沥青桥面横坡")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《沥青桥面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("桥面系") && ccxm.equals("混凝土桥面横坡")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }

            else if (fbgc.equals("隧道路面") && ccxm.equals("沥青路面压实度（隧道路面）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("沥青路面弯沉")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","0");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("沥青路面车辙（隧道路面）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《隧道路面车辙质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("沥青路面渗水系数（隧道路面）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《路面渗水系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("混凝土路面强度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土路面弯拉强度鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("混凝土路面相邻板高差")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土路面相邻板高差质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("平整度（隧道路面沥青）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("平整度（隧道路面混凝土）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("构造深度（隧道路面）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《隧道路面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("摩擦系数（隧道路面）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《隧道路面摩擦系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("厚度（隧道路面混凝土）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("厚度（隧道路面沥青）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("厚度（隧道路面雷达）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《隧道路面厚度质量鉴定表（雷达法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("横坡（隧道路面沥青）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《沥青路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("隧道路面") && ccxm.equals("横坡（隧道路面混凝土）")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("总体") && ccxm.equals("△净空")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《净空质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("总体") && ccxm.equals("宽度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《隧道宽度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("衬砌") && ccxm.equals("*衬砌强度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土强度质量鉴定表（回弹法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("衬砌") && ccxm.equals("△衬砌厚度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《隧道衬砌厚度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("衬砌") && ccxm.equals("大面平整度")){
                map.put("sheetname","分部-"+data.getDwgc());
                map.put("filename","详见《混凝土大面平整度试验检测鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }
            else if (fbgc.equals("路面面层") && ccxm.equals("△沥青路面压实度（路面面层）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△沥青路面压实度（隧道路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△沥青路面弯沉")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("沥青路面车辙（路面面层）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面车辙质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("沥青路面车辙（隧道路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《隧道路面车辙质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("沥青路面渗水系数（路面面层）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("沥青路面渗水系数（隧道路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("*混凝土路面强度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面强度鉴定结果汇总表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("混凝土路面相邻板高差")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面相邻板高差质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("平整度（沥青路面面层）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("平整度（混凝土路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("路面面层") && ccxm.equals("平整度（桥面系混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("平整度（桥面系沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("平整度（隧道路面混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("平整度（隧道路面沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《平整度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("路面面层") && ccxm.equals("构造深度（混凝土路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("构造深度（沥青路面面层）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("路面面层") && ccxm.equals("构造深度（桥面系混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("构造深度（桥面系沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《桥面系构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("构造深度（隧道路面混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("构造深度（隧道路面沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《隧道路面构造深度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("路面面层") && ccxm.equals("摩擦系数（沥青路面面层）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面摩擦系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("摩擦系数（桥面系）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《桥面系摩擦系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("摩擦系数（隧道路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《隧道路面摩擦系数质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("路面面层") && ccxm.equals("△厚度（混凝土路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△厚度（沥青路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△厚度（路面雷达法）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《路面厚度质量鉴定表（雷达法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△厚度（隧道路面混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△厚度（隧道路面沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("△厚度（隧道路面雷达法）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《隧道路面厚度质量鉴定表（雷达法）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","3");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("路面面层") && ccxm.equals("横坡（混凝土路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("横坡（沥青路面）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("横坡（桥面系混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("横坡（桥面系沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("横坡（隧道路面混凝土）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《混凝土路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("路面面层") && ccxm.equals("横坡（隧道路面沥青）")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《沥青路面横坡质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }

            else if (fbgc.equals("标志") && ccxm.equals("立柱竖直度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《交通标志板安装质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("标志") && ccxm.equals("标志板净空")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《交通标志板安装质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("标志") && ccxm.equals("标志板厚度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《交通标志板安装质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","1");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("标志") && ccxm.equals("标志面反光膜等级及逆射光系数")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《交通标志板安装质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            } else if (fbgc.equals("标线") && ccxm.equals("△反光标线逆反射系数")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《道路交通标线施工质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("标线") && ccxm.equals("△标线厚度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《道路交通标线施工质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("防护栏") && ccxm.equals("波形梁板基地金属厚度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《道路防护栏施工质量鉴定表（波形梁钢护栏）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("防护栏") && ccxm.equals("*波形梁钢护栏立柱壁厚")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《道路防护栏施工质量鉴定表（波形梁钢护栏）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("防护栏") && ccxm.equals("*波形梁钢护栏立柱埋入深度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《道路防护栏施工质量鉴定表（波形梁钢护栏）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("防护栏") && ccxm.equals("*波形梁钢护栏横梁中心高度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《道路防护栏施工质量鉴定表（波形梁钢护栏）》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("防护栏") && ccxm.equals("*混凝土护栏强度")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《交安工程砼护栏强度质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }else if (fbgc.equals("防护栏") && ccxm.equals("△砼护栏断面尺寸")){
                map.put("sheetname","分部-路面");
                map.put("filename","详见《交安砼护栏断面尺寸质量鉴定表》检测"+zds+"点,合格"+hgds+"点");
                map.put("qz","2");
                map.put("hgl",Double.valueOf(zds)!=0?df.format(Double.valueOf(hgds)/Double.valueOf(zds)*100):0);
                result.add(map);
            }
        }
        return result;
    }

    /**
     *
     * @param response
     * @param proname
     * @param result
     */
    private void exportExcel(HttpServletResponse response, String proname, List<JCSJVo> result) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = proname+"检测数据录入表";
            String sheetName1 = "检测数据";
            ExcelUtil.writeExcelWithSheets(response, result, fileName, sheetName1, new JCSJVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (Exception e) {
            throw new JjgysException(20001,"导出失败");
        }finally {
            File t = new File(courseFile+File.separator+proname+"检测数据录入表.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @return
     */
    private List<JCSJVo> getSddata(String proname, String htd) {
        String[] ccxm1 = {"*衬砌强度", "△衬砌厚度", "大面平整度"};
        String[] ccxm2 = {"宽度", "△净空"};
        String[] ccxm3 = {"沥青路面压实度（隧道路面）", "沥青路面弯沉","沥青路面车辙（隧道路面）","沥青路面渗水系数（隧道路面）","混凝土路面强度","混凝土路面相邻板高差","平整度（隧道路面沥青）",
                "平整度（隧道路面混凝土）","构造深度（隧道路面）","摩擦系数（隧道路面）","厚度（隧道路面混凝土）","厚度（隧道路面沥青）","厚度（隧道路面雷达）","横坡（隧道路面沥青）","横坡（隧道路面混凝土）"};
        List<JCSJVo> result = new ArrayList<>();
        QueryWrapper<JjgLqsJgSd> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        List<JjgLqsJgSd> htdinfo = jjgLqsJgSdMapper.selectList(wrapper);
        if (htdinfo!=null && !htdinfo.isEmpty()){
            for (JjgLqsJgSd jjgLqsJgSd : htdinfo) {
                String sdname = jjgLqsJgSd.getSdname();
                for (String ccxm : ccxm1) {
                    JCSJVo jgjcsj = new JCSJVo();
                    jgjcsj.setHtd(htd);
                    jgjcsj.setDwgc(sdname);
                    jgjcsj.setFbgc("衬砌");
                    jgjcsj.setCcxm(ccxm);
                    /*Map<String,Object> map = new HashMap<>();
                    map.put("htd", htd);
                    map.put("dwgc", sdname);
                    map.put("fbgc", "衬砌");
                    map.put("ccxm", ccxm);*/
                    result.add(jgjcsj);
                }
                for (String ccxm : ccxm2) {
                    /*Map<String,Object> map = new HashMap<>();
                    map.put("htd", htd);
                    map.put("dwgc", sdname);
                    map.put("fbgc", "总体");
                    map.put("ccxm", ccxm);*/
                    JCSJVo jgjcsj = new JCSJVo();
                    jgjcsj.setHtd(htd);
                    jgjcsj.setDwgc(sdname);
                    jgjcsj.setFbgc("总体");
                    jgjcsj.setCcxm(ccxm);
                    result.add(jgjcsj);
                }
                for (String ccxm : ccxm3) {
                    /*Map<String,Object> map = new HashMap<>();
                    map.put("htd", htd);
                    map.put("dwgc", sdname);
                    map.put("fbgc", "隧道路面");
                    map.put("ccxm", ccxm);*/
                    JCSJVo jgjcsj = new JCSJVo();
                    jgjcsj.setHtd(htd);
                    jgjcsj.setDwgc(sdname);
                    jgjcsj.setFbgc("隧道路面");
                    jgjcsj.setCcxm(ccxm);

                    result.add(jgjcsj);
                }

            }
        }
        return result;
    }

    /**
     *
     * @param proname
     * @param htd
     * @return
     */
    private List<JCSJVo> getQldata(String proname, String htd) {
        String[] ccxm1 = {"*混凝土强度", "主要结构尺寸", "钢筋保护层厚度"};
        String[] ccxm2 = {"*混凝土强度", "主要结构尺寸","钢筋保护层厚度","*竖直度"};
        String[] ccxm3 = {"沥青桥平整度", "混凝土桥平整度","沥青桥面构造深度","混凝土桥面构造深度","摩擦系数","沥青桥面横坡","混凝土桥面横坡"};
        List<JCSJVo> result = new ArrayList<>();
        QueryWrapper<JjgLqsJgQl> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        List<JjgLqsJgQl> htdinfo = jjgLqsJgQlMapper.selectList(wrapper);
        if (htdinfo!=null && !htdinfo.isEmpty()){
            for (JjgLqsJgQl jjgLqsJgQl : htdinfo) {
                String qlname = jjgLqsJgQl.getQlname();
                for (String ccxm : ccxm1) {
                   /* Map<String,Object> map = new HashMap<>();
                    map.put("htd", htd);
                    map.put("dwgc", qlname);
                    map.put("fbgc", "上部");
                    map.put("ccxm", ccxm);*/
                    JCSJVo jgjcsj = new JCSJVo();
                    jgjcsj.setHtd(htd);
                    jgjcsj.setDwgc(qlname);
                    jgjcsj.setFbgc("桥梁上部");
                    jgjcsj.setCcxm(ccxm);
                    result.add(jgjcsj);
                }
                for (String ccxm : ccxm2) {
                    /*Map<String,Object> map = new HashMap<>();
                    map.put("htd", htd);
                    map.put("dwgc", qlname);
                    map.put("fbgc", "下部");
                    map.put("ccxm", ccxm);*/
                    JCSJVo jgjcsj = new JCSJVo();
                    jgjcsj.setHtd(htd);
                    jgjcsj.setDwgc(qlname);
                    jgjcsj.setFbgc("桥梁下部");
                    jgjcsj.setCcxm(ccxm);
                    result.add(jgjcsj);
                }
                for (String ccxm : ccxm3) {
                    /*Map<String,Object> map = new HashMap<>();
                    map.put("htd", htd);
                    map.put("dwgc", qlname);
                    map.put("fbgc", "桥面系");
                    map.put("ccxm", ccxm);*/
                    JCSJVo jgjcsj = new JCSJVo();
                    jgjcsj.setHtd(htd);
                    jgjcsj.setDwgc(qlname);
                    jgjcsj.setFbgc("桥面系");
                    jgjcsj.setCcxm(ccxm);
                    result.add(jgjcsj);
                }
            }
        }
        return result;
    }

    /**
     *
     * @param htd
     * @return
     */
    private List<JCSJVo> getJadata(String htd) {
        String[] ccxm1 = {"立柱竖直度", "标志板净空", "标志板厚度", "标志面反光膜等级及逆射光系数"};
        String[] ccxm2 = {"△反光标线逆反射系数", "△标线厚度"};
        String[] ccxm3 = {"*波形梁板基地金属厚度", "*波形梁钢护栏立柱壁厚","*波形梁钢护栏立柱埋入深度","*波形梁钢护栏横梁中心高度","*混凝土护栏强度","△砼护栏断面尺寸"};
        List<JCSJVo> result = new ArrayList<>();
        for (String ccxm : ccxm1) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "交通安全设施");
            map.put("fbgc", "标志");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("交通安全设施");
            jgjcsj.setFbgc("标志");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        for (String ccxm : ccxm2) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "交通安全设施");
            map.put("fbgc", "标线");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("交通安全设施");
            jgjcsj.setFbgc("标线");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        for (String ccxm : ccxm3) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "交通安全设施");
            map.put("fbgc", "防护栏");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("交通安全设施");
            jgjcsj.setFbgc("防护栏");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        return result;

    }

    /**
     * @param htd
     * @param htd
     * @return
     */
    private List<JCSJVo> getLmdata(String proname, String htd) throws IOException {
        DecimalFormat df = new DecimalFormat("0");
        CommonInfoVo c = new CommonInfoVo();
        c.setProname(proname);
        c.setHtd(htd);

        //弯沉
        List<Map<String, Object>> maps1 = jjgFbgcLmgcLmwcJgfcService.lookJdbjg(c);
        List<Map<String, Object>> maps2 = jjgFbgcLmgcLmwcLcfJgfcService.lookJdbjg(c);
        double wczds = 0,wchgds = 0;
        String wcgdz = "";
        if (maps1 != null && maps1.size()>0){
            for (Map<String, Object> map : maps1) {
                wczds += Double.valueOf(map.get("检测单元数").toString());
                wchgds += Double.valueOf(map.get("合格单元数").toString());
                wcgdz = map.get("规定值").toString();

            }
        }
        if (maps2 != null && maps2.size()>0){
            for (Map<String, Object> map : maps2) {
                wczds += Double.valueOf(map.get("检测单元数").toString());
                wchgds += Double.valueOf(map.get("合格单元数").toString());
                wcgdz = map.get("规定值").toString();

            }
        }

        //车辙
        List<Map<String, Object>> maps3 = jjgZdhCzJgfcService.lookJdbjg(proname);
        double czlmzds = 0,czlmhgds = 0,czsdzds = 0,czsdhgds = 0,czqzds = 0,czqhgds = 0;
        String czlmgdz = "",czsdgdz = "",czqgdz = "";
        if (maps3 != null && maps3.size()>0){
            for (Map<String, Object> map : maps3) {
                String htd1 = map.get("htd").toString();
                if (htd1.equals(htd)){
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("路面")){
                        czlmgdz = map.get("设计值").toString();
                        czlmzds += Double.valueOf(map.get("总点数").toString());
                        czlmhgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lmlx.contains("隧道")){
                        czsdgdz = map.get("设计值").toString();
                        czsdzds += Double.valueOf(map.get("总点数").toString());
                        czsdhgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lmlx.contains("桥")){
                        czqgdz = map.get("设计值").toString();
                        czqzds += Double.valueOf(map.get("总点数").toString());
                        czqhgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }
            }

        }
        //混凝土路面相邻板高差
        double xlbgczds = 0,xlbgchgds = 0;
        String xlbgcgdz = "";
        List<Map<String, Object>> maps4 = jjgFbgcLmgcTlmxlbgcJgfcService.lookJdbjg(proname);
        if (maps4 != null && maps4.size()>0){
            for (Map<String, Object> map : maps4) {
                String htd1 = map.get("htd").toString();
                if (htd1.equals(htd)){
                    xlbgcgdz = map.get("规定值").toString();
                    xlbgczds += Double.valueOf(map.get("总点数").toString());
                    xlbgchgds += Double.valueOf(map.get("合格点数").toString());
                }
            }

        }

        //平整度
        double pzdlmzds = 0,pzdlmhgds = 0,pzdsdzds = 0,pzdsdhgds = 0,pzdqzds = 0,pzdqhgds = 0,pzdhntzds = 0,pzdhnthgds = 0;
        String pzdlmgdz = "",pzdsdgdz = "",pzdqgdz = "",pzdhntgdz = "";
        List<Map<String, Object>> maps5 = jjgZdhPzdJgfcService.lookJdbjg(proname);
        if (maps5 != null && maps5.size()>0){
            for (Map<String, Object> map : maps5) {
                String htd1 = map.get("htd").toString();
                if (htd1.equals(htd)){
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.equals("沥青路面") || lmlx.equals("沥青匝道")){
                        pzdlmgdz = map.get("设计值").toString();
                        pzdlmzds += Double.valueOf(map.get("总点数").toString());
                        pzdlmhgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lmlx.contains("隧道")){
                        pzdsdgdz = map.get("设计值").toString();
                        pzdsdzds += Double.valueOf(map.get("总点数").toString());
                        pzdsdhgds += Double.valueOf(map.get("合格点数").toString());

                    }else if (lmlx.contains("桥")){
                        pzdqgdz = map.get("设计值").toString();
                        pzdqzds += Double.valueOf(map.get("总点数").toString());
                        pzdqhgds += Double.valueOf(map.get("合格点数").toString());

                    }else if (lmlx.contains("混凝土")){
                        pzdhntgdz = map.get("设计值").toString();
                        pzdhntzds += Double.valueOf(map.get("总点数").toString());
                        pzdhnthgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }
            }

        }

        double gzsdlmzds = 0,gzsdlmhgds = 0,gzsdsdzds = 0,gzsdsdhgds = 0,gzsdqzds = 0,gzsdqhgds = 0;
        String gzsdlmgdz = "",gzsdsdgdz = "",gzsdqgdz = "";
        List<Map<String, Object>> maps6 = jjgZdhGzsdJgfcService.lookJdbjg(proname);
        if (maps6 != null && maps6.size()>0){
            for (Map<String, Object> map : maps6) {
                String htd1 = map.get("htd").toString();
                if (htd1.equals(htd)){
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("路面")){
                        gzsdlmgdz = map.get("设计值").toString();
                        gzsdlmzds += Double.valueOf(map.get("总点数").toString());
                        gzsdlmhgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lmlx.contains("隧道")){
                        gzsdsdgdz = map.get("设计值").toString();
                        gzsdsdzds += Double.valueOf(map.get("总点数").toString());
                        gzsdsdhgds += Double.valueOf(map.get("合格点数").toString());

                    }else if (lmlx.contains("桥")){
                        gzsdqgdz = map.get("设计值").toString();
                        gzsdqzds += Double.valueOf(map.get("总点数").toString());
                        gzsdqhgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }

            }

        }

        double mcxslmzds = 0,mcxslmhgds = 0,mcxssdzds = 0,mcxssdhgds = 0,mcxsqzds = 0,mcxsqhgds = 0;
        String mcxslmgdz = "",mcxssdgdz = "",mcxsqgdz = "";
        List<Map<String, Object>> maps7 = jjgZdhMcxsJgfcService.lookJdbjg(proname);
        if (maps7 != null && maps7.size()>0){
            for (Map<String, Object> map : maps7) {
                String htd1 = map.get("htd").toString();
                if (htd1.equals(htd)){
                    String lmlx = map.get("路面类型").toString();
                    if (lmlx.contains("路面")){
                        mcxslmgdz = map.get("设计值").toString();
                        mcxslmzds += Double.valueOf(map.get("总点数").toString());
                        mcxslmhgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lmlx.contains("隧道")){
                        mcxssdgdz = map.get("设计值").toString();
                        mcxssdzds += Double.valueOf(map.get("总点数").toString());
                        mcxssdhgds += Double.valueOf(map.get("合格点数").toString());

                    }else if (lmlx.contains("桥")){
                        mcxsqgdz = map.get("设计值").toString();
                        mcxsqzds += Double.valueOf(map.get("总点数").toString());
                        mcxsqhgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }

            }
        }


        List<JCSJVo> result = new ArrayList<>();
        String[] ccxm = {"△沥青路面压实度（路面面层）", "△沥青路面压实度（隧道路面）", "△沥青路面弯沉", "沥青路面车辙（路面面层）", "沥青路面车辙（桥面）","沥青路面车辙（隧道路面）",
                "沥青路面渗水系数（路面面层）","沥青路面渗水系数（隧道路面）","*混凝土路面强度","混凝土路面相邻板高差",
                "平整度（沥青路面面层）", "平整度（桥面系沥青）","平整度（隧道路面沥青）",
                "平整度（混凝土路面）","平整度（桥面系混凝土）","平整度（隧道路面混凝土）",
                "构造深度（沥青路面面层）","构造深度（桥面系沥青）","构造深度（隧道路面沥青）",
                "构造深度（混凝土路面）","构造深度（桥面系混凝土）", "构造深度（隧道路面混凝土）",
                "摩擦系数（沥青路面面层）","摩擦系数（桥面系）","摩擦系数（隧道路面）",
                "△厚度（混凝土路面）","△厚度（沥青路面）","△厚度（路面雷达法）"
                ,"△厚度（隧道路面混凝土）","△厚度（隧道路面沥青）","△厚度（隧道路面雷达法）","横坡（混凝土路面）","横坡（沥青路面）","横坡（桥面系混凝土）","横坡（桥面系沥青）","横坡（隧道路面混凝土）","横坡（隧道路面沥青）"};

        for (String s : ccxm) {
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("路面工程");
            jgjcsj.setFbgc("路面面层");
            jgjcsj.setCcxm(s);
            if ("摩擦系数（沥青路面面层）".equals(s)){
                jgjcsj.setGdz(mcxslmgdz);
                jgjcsj.setZds(String.valueOf(df.format(mcxslmzds)));
                jgjcsj.setHgds(String.valueOf(df.format(mcxslmhgds)));
            }
            if ("摩擦系数（桥面系）".equals(s)){
                jgjcsj.setGdz(mcxsqgdz);
                jgjcsj.setZds(String.valueOf(df.format(mcxsqzds)));
                jgjcsj.setHgds(String.valueOf(df.format(mcxsqhgds)));
            }
            if ("摩擦系数（隧道路面）".equals(s)){
                jgjcsj.setGdz(mcxssdgdz);
                jgjcsj.setZds(String.valueOf(df.format(mcxssdzds)));
                jgjcsj.setHgds(String.valueOf(df.format(mcxssdhgds)));
            }

            if ("构造深度（沥青路面面层）".equals(s)){
                jgjcsj.setGdz(gzsdlmgdz);
                jgjcsj.setZds(String.valueOf(df.format(gzsdlmzds)));
                jgjcsj.setHgds(String.valueOf(df.format(gzsdlmhgds)));
            }
            if ("构造深度（桥面系沥青）".equals(s)){
                jgjcsj.setGdz(gzsdqgdz);
                jgjcsj.setZds(String.valueOf(df.format(gzsdqzds)));
                jgjcsj.setHgds(String.valueOf(df.format(gzsdqhgds)));
            }
            if ("构造深度（隧道路面沥青）".equals(s)){
                jgjcsj.setGdz(gzsdsdgdz);
                jgjcsj.setZds(String.valueOf(df.format(gzsdsdzds)));
                jgjcsj.setHgds(String.valueOf(df.format(gzsdsdhgds)));
            }

            if ("平整度（沥青路面面层）".equals(s)){
                jgjcsj.setGdz(pzdlmgdz);
                jgjcsj.setZds(String.valueOf(df.format(pzdlmzds)));
                jgjcsj.setHgds(String.valueOf(df.format(pzdlmhgds)));
            }
            if ("平整度（桥面系沥青）".equals(s)){
                jgjcsj.setGdz(pzdqgdz);
                jgjcsj.setZds(String.valueOf(df.format(pzdqzds)));
                jgjcsj.setHgds(String.valueOf(df.format(pzdqhgds)));
            }
            if ("平整度（隧道路面沥青）".equals(s)){
                jgjcsj.setGdz(pzdsdgdz);
                jgjcsj.setZds(String.valueOf(df.format(pzdsdzds)));
                jgjcsj.setHgds(String.valueOf(df.format(pzdsdhgds)));
            }
            if ("平整度（混凝土路面）".equals(s)){
                jgjcsj.setGdz(pzdhntgdz);
                jgjcsj.setZds(String.valueOf(df.format(pzdhntzds)));
                jgjcsj.setHgds(String.valueOf(df.format(pzdhnthgds)));
            }
            if ("△沥青路面弯沉".equals(s)){
                jgjcsj.setGdz(StringUtils.substringBetween(wcgdz,"[","]"));
                jgjcsj.setZds(String.valueOf(df.format(wczds)));
                jgjcsj.setHgds(String.valueOf(df.format(wchgds)));
            }
            if ("沥青路面车辙（路面面层）".equals(s)){
                jgjcsj.setGdz(czlmgdz);
                jgjcsj.setZds(String.valueOf(df.format(czlmzds)));
                jgjcsj.setHgds(String.valueOf(df.format(czlmhgds)));
            }
            if ("沥青路面车辙（桥面）".equals(s)){
                jgjcsj.setGdz(czqgdz);
                jgjcsj.setZds(String.valueOf(df.format(czqzds)));
                jgjcsj.setHgds(String.valueOf(df.format(czqhgds)));
            }
            if ("沥青路面车辙（隧道路面）".equals(s)){
                jgjcsj.setGdz(czsdgdz);
                jgjcsj.setZds(String.valueOf(df.format(czsdzds)));
                jgjcsj.setHgds(String.valueOf(df.format(czsdhgds)));
            }
            if ("混凝土路面相邻板高差".equals(s)){
                jgjcsj.setGdz(xlbgcgdz);
                jgjcsj.setZds(String.valueOf(df.format(xlbgczds)));
                jgjcsj.setHgds(String.valueOf(df.format(xlbgchgds)));
            }
            result.add(jgjcsj);
        }
        return result;
    }

    /**
     *
     * @param htd
     * @return
     */
    private List<JCSJVo> getLjdata(String htd){
        String[] ccxm1 = {"△沉降", "△压实度（沙砾）", "△压实度（灰土）", "△弯沉", "边坡"};
        String[] ccxm2 = {"断面尺寸", "铺砌厚度"};
        String[] ccxm3 = {"*混凝土强度", "△断面尺寸"};
        String[] ccxm4 = {"*混凝土强度", "结构尺寸"};
        List<JCSJVo> result = new ArrayList<>();
        for (String ccxm : ccxm1) {
            //Map<String,Object> map = new HashMap<>();
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("路基工程");
            jgjcsj.setFbgc("路基土石方");
            jgjcsj.setCcxm(ccxm);
            /*map.put("htd", htd);
            map.put("dwgc", "路基工程");
            map.put("fbgc", "路基土石方");
            map.put("ccxm", ccxm);*/
            result.add(jgjcsj);
        }
        for (String ccxm : ccxm2) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "路基工程");
            map.put("fbgc", "排水工程");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("路基工程");
            jgjcsj.setFbgc("排水工程");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        for (String ccxm : ccxm3) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "路基工程");
            map.put("fbgc", "小桥");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("路基工程");
            jgjcsj.setFbgc("小桥");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        for (String ccxm : ccxm4) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "路基工程");
            map.put("fbgc", "涵洞");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("路基工程");
            jgjcsj.setFbgc("涵洞");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        for (String ccxm : ccxm3) {
            /*Map<String,Object> map = new HashMap<>();
            map.put("htd", htd);
            map.put("dwgc", "路基工程");
            map.put("fbgc", "支挡工程");
            map.put("ccxm", ccxm);*/
            JCSJVo jgjcsj = new JCSJVo();
            jgjcsj.setHtd(htd);
            jgjcsj.setDwgc("路基工程");
            jgjcsj.setFbgc("支挡工程");
            jgjcsj.setCcxm(ccxm);
            result.add(jgjcsj);
        }
        return result;

    }

}
