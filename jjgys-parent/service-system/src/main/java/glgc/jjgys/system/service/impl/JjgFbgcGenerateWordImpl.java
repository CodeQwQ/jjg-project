package glgc.jjgys.system.service.impl;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.conditions.update.UpdateWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;
import com.spire.xls.CellRange;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.WorksheetsCollection;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.Xmtj;
import glgc.jjgys.system.mapper.JjgFbgcGenerateWordMapper;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;


@Slf4j
@Service
public class JjgFbgcGenerateWordImpl extends ServiceImpl<JjgFbgcGenerateWordMapper, Object> implements JjgFbgcGenerateWordService {

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @Autowired
    private ProjectService projectService;

    @Autowired
    private JjgXmtjService jjgXmtjService;

    @Autowired
    private JjgHtdService jjgHtdService;

    @Autowired
    private JjgLqsQlService jjgLqsQlService;

    @Autowired
    private JjgLqsSdService jjgLqsSdService;

    @Autowired
    private JjgLqsHntlmzdService jjgLqsHntlmzdService;

    @Autowired
    private JjgLqsLjxService jjgLqsLjxService;

    @Autowired
    private JjgLqsSfzService jjgLqsSfzService;

    @Autowired
    private JjgNyzlkfJgService jjgNyzlkfJgService;

    @Autowired
    private JjgWgjcService jjgWgjcService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;

    @Autowired
    private JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;

    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Autowired
    private JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;

    @Autowired
    private JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;

    @Autowired
    private JjgFbgcLjgcXqgqdService jjgFbgcLjgcXqgqdService;

    @Autowired
    private JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;

    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;

    @Autowired
    private JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;

    @Autowired
    private JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;

    @Autowired
    private JjgFbgcQlgcXbJgccService jjgFbgcQlgcXbJgccService;

    @Autowired
    private JjgFbgcQlgcXbBhchdService jjgFbgcQlgcXbBhchdService;

    @Autowired
    private JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Autowired
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;

    @Autowired
    private JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;

    @Autowired
    private JjgFbgcSdgcCqtqdService jjgFbgcSdgcCqtqdService;

    @Autowired
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;

    @Autowired
    private JjgFbgcSdgcDmpzdService jjgFbgcSdgcDmpzdService;

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private JjgFbgcSdgcJkService jjgFbgcSdgcJkService;

    @Autowired
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Autowired
    private JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;

    @Autowired
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Autowired
    private JjgFbgcLmgcLmwcService jjgFbgcLmgcLmwcService;

    @Autowired
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;

    @Autowired
    private JjgFbgcLmgcLmssxsService jjgFbgcLmgcLmssxsService;

    @Autowired
    private JjgZdhPzdService jjgZdhPzdService;

    @Autowired
    private JjgZdhMcxsService jjgZdhMcxsService;

    @Autowired
    private JjgZdhGzsdService jjgZdhGzsdService;

    @Autowired
    private JjgZdhCzService jjgZdhCzService;

    @Autowired
    private JjgZdhLdhdService jjgZdhLdhdService;

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Autowired
    private JjgFbgcLmgcHntlmqdService jjgFbgcLmgcHntlmqdService;

    @Autowired
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfService jjgFbgcLmgcLmgzsdsgpsfServicel;

    @Autowired
    private JjgFbgcLmgcTlmxlbgcService jjgFbgcLmgcTlmxlbgcService;

    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;


    @Override
    public boolean generateword(String proname) throws IOException {
        String excelFilePath = filespath+File.separator+proname+File.separator+"报告中表格.xlsx";
        File excelf = new File(excelFilePath);
        File f = new File(filespath + File.separator + proname + File.separator + "高速公路交工验收质量检测报告.docx");
        File fdir = new File(filespath + File.separator + proname);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }

        //替换项目全称
        QueryWrapper<Project> wrapper = new QueryWrapper<>();
        wrapper.eq("proname", proname);
        Project pro = projectService.getOne(wrapper);

        InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/交工报告.docx");
        Document xw = new Document(inputStream);

        xw.replace("{项目全称}", pro.getXmqc(), false, true);
        xw.replace("${项目名称}", pro.getProName(), false, true);
        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

        CreateTableryqzbData(f, xw, proname);

        //表1.3-1 合同段划分及施工单位、监理单位一览表
        CreateTableData(f, xw, proname);
        //表4.1-2 仪器信息
        CreateTableyqData(f, xw, proname);
        //表2.2-1 人员信息
        CreateTableryData(f, xw, proname);
        //附表1桥梁清单
        CreateTableqlqdData(f, xw, proname);
        //附表2隧道清单
        CreateTablesdqdData(f, xw, proname);
        //外观检查结果（文字描述）
        //CreateTablewgjclwzms(f, xw, proname);
        //读取Excel文件
        if (excelf.exists()) {
            Workbook workbook = new Workbook();
            workbook.loadFromFile(excelFilePath);
            WorksheetsCollection sheetnum = workbook.getWorksheets();
            for (int j = 0; j < sheetnum.getCount() - 1; j++) {
                String sheetname = sheetnum.get(j).getName();
                Worksheet sheet = workbook.getWorksheets().get(j);
                boolean ishidden = JjgFbgcCommonUtils.ishidden(excelFilePath, sheetname);
                if (!ishidden) {
                    CellRange allocatedRange = sheet.getAllocatedRange();
                    String s = allocatedRange.get(1, 1).getDisplayedText();
                    log.info("表名：{}", s);
                    //TextSelection[] selection = xw.findAllString(s, true, true);
                    TextSelection[] selection = findStringInPagess(xw, s, 10, 60);
                    if (selection != null) {
                        for (int k = 0; k < selection.length; k++) {
                            TextRange range = selection[k].getAsOneRange();
                            //删除没用数据的行
                            CellRange dataRange = RemoveEmptyRows(sheet, allocatedRange);
                            //复制到Word文档
                            log.info("开始复制{}中的数据到word中", sheetname);
                            log.info("此处共{}处匹配，目前第{}处", selection.length, k + 1);
                            copyToWord(dataRange, xw, f.getPath(), range);
                            log.info("{}中的数据复制完成", sheetname);
                        }
                    }
                }
            }
        }
        //工程质量评定 路基工程
        CreateTableljzlpd(f, xw, proname);

        //工程质量评定 路面工程
        CreateTablelmzlpd(f, xw, proname);

        //工程质量评定 桥梁工程
        CreateTableqlzlpd(f, xw, proname);

        //工程质量评定 隧道工程
        CreateTablesdzlpd(f, xw, proname);

        //工程质量评定 交安工程
        CreateTablejazlpd(f, xw, proname);

        //工程质量评定 合同段评定
        CreateTablehtdzlpd(f, xw, proname);

        //建设项目质量评定
        CreateTablejsxmpd(f, xw, proname);

        //外观检查
        CreateTablewgjcl(f, xw, proname);

        //内页资料
        CreateTablenyzl(f, xw, proname);

        //交工验收工程质量评定表  建设项目质量评定表
        CreateTablejsxmzlpdb(f, xw, proname);

        CreateTablehtdpdb(f, xw, proname);

        //主要指标检测结果汇总表
        CreateTablezyzbjcjghz(f, xw, proname);




        //更新状态
        UpdateWrapper<Xmtj> updateWrapper = new UpdateWrapper<>();
        updateWrapper.set("type", 2);
        updateWrapper.set("jgnf", String.valueOf(LocalDate.now()));
        updateWrapper.eq("proname", proname);
        jjgXmtjService.update(null, updateWrapper);


        return true;

    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablezyzbjcjghz(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> list = getzyzbjcjghzData(proname);
        if (list != null && list.size()>0){
            for (Map<String, Object> map : list) {
                for (Map.Entry<String, Object> entry : map.entrySet()) {
                    String s = "${"+entry.getKey()+"}";
                    String value = entry.getValue().toString();
                    TextSelection[] selection = findStringInPagess(xw, s, 50, 80);
                    if (selection != null && selection.length > 0) {
                        for (int k = 0; k < selection.length; k++) {
                            TextRange range = selection[k].getAsOneRange();
                            String originalText = range.getText(); // 假设有一个方法来获取TextRange的文本

                            // 检查originalText中是否包含与key相同的字符串（去掉${}）
                            if (originalText.contains(s)) {
                                // 替换文本（这里假设您有一个方法来设置TextRange的文本）
                                // 注意：这里只是一个示例，具体实现取决于您的TextSelection和TextRange API
                                range.setText(value); // 假设存在这样一个方法来设置TextRange的文本
                            }
                        }
                    }
                }
            }
            xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
        }

    }

    /**
     *
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getzyzbjcjghzData(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        List<Map<String,Object>> resultList = new ArrayList<>();
        if (list != null && list.size() > 0){
            List<String> ljhtdlist = new ArrayList<>();
            List<String> lmhtdlist = new ArrayList<>();
            List<String> qlhtdlist = new ArrayList<>();
            List<String> sdhtdlist = new ArrayList<>();
            List<String> jahtdlist = new ArrayList<>();

            for (JjgHtd jjgHtd : list) {
                String lx = jjgHtd.getLx();
                String name = jjgHtd.getName();
                if (lx.contains("路基工程")){
                    ljhtdlist.add(name);
                }
                if (lx.contains("路面工程")){
                    lmhtdlist.add(name);
                }
                if (lx.contains("桥梁工程")){
                    qlhtdlist.add(name);
                }
                if (lx.contains("隧道工程")){
                    sdhtdlist.add(name);
                }
                if (lx.contains("交安工程")){
                    jahtdlist.add(name);
                }
            }

            if (ljhtdlist != null && ljhtdlist.size() > 0){
                double ysdzds = 0,ysdhgds = 0;
                double wczds = 0,wchgds = 0;
                double bpzds = 0,bphgds = 0;
                double psdmcczds = 0,psdmcchgds = 0;
                double pspqhdzds = 0,pspqhdhgds = 0;
                double xqtqdzds = 0,xqtqdhgds = 0;
                double xqjgcczds = 0,xqjgcchgds = 0;
                double hdtqdzds = 0,hdtqdhgds = 0;
                double hdjgcczds = 0,hdjgcchgds = 0;
                double zdgqdzds = 0,zdgqdhgds = 0;
                double zddmcczds = 0,zddmcchgds = 0;

                for (String htd : ljhtdlist) {
                    CommonInfoVo commonInfoVo = new CommonInfoVo();
                    commonInfoVo.setProname(proname);
                    commonInfoVo.setHtd(htd);
                    List<Map<String, Object>> ysdlist = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
                    if (ysdlist != null && ysdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : ysdlist) {
                            ysdzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                            ysdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }
                    List<Map<String, Object>> cjlist = jjgFbgcLjgcLjcjService.lookJdbjg(commonInfoVo);
                    if (cjlist != null && cjlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : cjlist) {
                            ysdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            ysdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> wclist = jjgFbgcLjgcLjwcService.lookJdbjg(commonInfoVo);
                    if (wclist != null && wclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : wclist) {
                            wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                            wchgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());

                        }
                    }
                    List<Map<String, Object>> wclcflist = jjgFbgcLjgcLjwcLcfService.lookJdbjg(commonInfoVo);
                    if (wclcflist != null && wclcflist.size() > 0){
                        for (Map<String, Object> stringObjectMap : wclcflist) {
                            wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                            wchgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
                        }

                    }
                    List<Map<String, Object>> ljbplist = jjgFbgcLjgcLjbpService.lookJdbjg(commonInfoVo);
                    if (ljbplist != null && ljbplist.size() > 0){
                        for (Map<String, Object> stringObjectMap : ljbplist) {
                            bpzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            bphgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> psdmcclist = jjgFbgcLjgcPsdmccService.lookJdbjg(commonInfoVo);
                    if (psdmcclist != null && psdmcclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : psdmcclist) {
                            psdmcczds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            psdmcchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> pspqhdlist = jjgFbgcLjgcPspqhdService.lookJdbjg(commonInfoVo);
                    if (pspqhdlist != null && pspqhdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : pspqhdlist) {
                            pspqhdzds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            pspqhdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> xqtqdlist = jjgFbgcLjgcXqgqdService.lookJdbjg(commonInfoVo);
                    if (xqtqdlist != null && xqtqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : xqtqdlist) {
                            xqtqdzds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            xqtqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> xqjgcclist = jjgFbgcLjgcXqjgccService.lookJdbjg(commonInfoVo);
                    if (xqjgcclist != null && xqjgcclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : xqjgcclist) {
                            xqjgcczds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            xqjgcchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> hdtqdlist = jjgFbgcLjgcHdgqdService.lookJdbjg(commonInfoVo);
                    if (hdtqdlist != null && hdtqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : hdtqdlist) {
                            hdtqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            hdtqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> hdjgcclist = jjgFbgcLjgcHdjgccService.lookJdbjg(commonInfoVo);
                    if (hdjgcclist != null && hdjgcclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : hdjgcclist) {
                            hdjgcczds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            hdjgcchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> zdgqdlist = jjgFbgcLjgcZdgqdService.lookJdbjg(commonInfoVo);
                    if (zdgqdlist != null && zdgqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : zdgqdlist) {
                            zdgqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            zdgqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> zddmcclist = jjgFbgcLjgcZddmccService.lookJdbjg(commonInfoVo);
                    if (zddmcclist != null && zddmcclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : zddmcclist) {
                            zddmcczds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            zddmcchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }
                }

                Map<String,Object> ysdmap  =new HashMap<>();
                ysdmap.put("ysdzds",decf.format(ysdzds));
                ysdmap.put("ysdhgds",decf.format(ysdhgds));
                ysdmap.put("ysdhgl",ysdzds != 0 ? df.format(ysdhgds / ysdzds * 100) : 0);
                resultList.add(ysdmap);

                Map<String,Object> wcmap  =new HashMap<>();
                wcmap.put("wczds",decf.format(wczds));
                wcmap.put("wchgds",decf.format(wchgds));
                wcmap.put("wchgl",wczds != 0 ? df.format(wchgds / wczds * 100) : 0);
                resultList.add(wcmap);

                Map<String,Object> bpmap  =new HashMap<>();
                bpmap.put("bpzds",decf.format(bpzds));
                bpmap.put("bphgds",decf.format(bphgds));
                bpmap.put("bphgl",bpzds != 0 ? df.format(bphgds / bpzds * 100) : 0);
                resultList.add(bpmap);

                Map<String,Object> psdmccmap  =new HashMap<>();
                psdmccmap.put("psdmcczds",decf.format(psdmcczds));
                psdmccmap.put("psdmcchgds",decf.format(psdmcchgds));
                psdmccmap.put("psdmcchgl",psdmcczds != 0 ? df.format(psdmcchgds / psdmcczds * 100) : 0);
                resultList.add(psdmccmap);

                Map<String,Object> pspqhdmap  =new HashMap<>();
                pspqhdmap.put("pspqhdzds",decf.format(pspqhdzds));
                pspqhdmap.put("pspqhdhgds",decf.format(pspqhdhgds));
                pspqhdmap.put("pspqhdhgl",pspqhdzds != 0 ? df.format(pspqhdhgds / pspqhdzds * 100) : 0);
                resultList.add(pspqhdmap);

                Map<String,Object> xqtqdmap  =new HashMap<>();
                xqtqdmap.put("xqtqdzds",decf.format(xqtqdzds));
                xqtqdmap.put("xqtqdhgds",decf.format(xqtqdhgds));
                xqtqdmap.put("xqtqdhgl",xqtqdzds != 0 ? df.format(xqtqdhgds / xqtqdzds * 100) : 0);
                resultList.add(xqtqdmap);

                Map<String,Object> xqjgccmap  =new HashMap<>();
                xqjgccmap.put("xqjgcczds",decf.format(xqjgcczds));
                xqjgccmap.put("xqjgcchgds",decf.format(xqjgcchgds));
                xqjgccmap.put("xqjgcchgl",xqjgcczds != 0 ? df.format(xqjgcchgds / xqjgcczds * 100) : 0);
                resultList.add(xqjgccmap);

                Map<String,Object> hdtqdmap  =new HashMap<>();
                hdtqdmap.put("hdtqdzds",decf.format(hdtqdzds));
                hdtqdmap.put("hdtqdhgds",decf.format(hdtqdhgds));
                hdtqdmap.put("hdtqdhgl",hdtqdzds != 0 ? df.format(hdtqdhgds / hdtqdzds * 100) : 0);
                resultList.add(hdtqdmap);

                Map<String,Object> hdjgccmap  =new HashMap<>();
                hdjgccmap.put("hdjgcczds",decf.format(hdjgcczds));
                hdjgccmap.put("hdjgcchgds",decf.format(hdjgcchgds));
                hdjgccmap.put("hdjgcchgl",hdjgcczds != 0 ? df.format(hdjgcchgds / hdjgcczds * 100) : 0);
                resultList.add(hdjgccmap);

                Map<String,Object> zdgqdmap  =new HashMap<>();
                zdgqdmap.put("zdgqdzds",decf.format(zdgqdzds));
                zdgqdmap.put("zdgqdhgds",decf.format(zdgqdhgds));
                zdgqdmap.put("zdgqdhgl",zdgqdzds != 0 ? df.format(zdgqdhgds / zdgqdzds * 100) : 0);
                resultList.add(zdgqdmap);

                Map<String,Object> zddmccmap  =new HashMap<>();
                zddmccmap.put("zddmcczds",decf.format(zddmcczds));
                zddmccmap.put("zddmcchgds",decf.format(zddmcchgds));
                zddmccmap.put("zddmcchgl",zddmcczds != 0 ? df.format(zddmcchgds / zddmcczds * 100) : 0);
                resultList.add(zddmccmap);
            }

            if (lmhtdlist != null && lmhtdlist.size() > 0){
                double ysdzds = 0,ysdhgds = 0;
                double wczds = 0,wchgds = 0;
                double ssxszds = 0,ssxshgds = 0;
                double pzdzds = 0,pzdhgds = 0;
                double mcxszds = 0,mcxshgds = 0;
                double gzsdzds = 0,gzsdhgds = 0;
                double czzds = 0,czhgds = 0;
                double ldhdzds = 0,ldhdhgds = 0;
                double lqlmhpzds = 0,lqlmhphgds = 0;
                double hntlmhpzds = 0,hntlmhphgds = 0;
                double hntlmqdzds = 0,hntlmqdhgds = 0;
                double hntlmhdzds = 0,hntlmhdhgds = 0;
                double gzsdsgpsfzds = 0,gzsdsgpsfhgds = 0;
                double tlmxlbgczds = 0,tlmxlbgchgds = 0;

                for (String htd : lmhtdlist) {
                    CommonInfoVo commonInfoVo = new CommonInfoVo();
                    commonInfoVo.setProname(proname);
                    commonInfoVo.setHtd(htd);

                    List<Map<String, Object>> ysdlist = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
                    if (ysdlist != null && ysdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : ysdlist) {
                            ysdzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                            ysdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> wclist = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo,1);
                    List<Map<String, Object>> wclcflist = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo,1);
                    if (wclist != null && wclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : wclist) {
                            wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                            wchgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
                        }
                    }
                    if (wclcflist != null && wclcflist.size() > 0){
                        for (Map<String, Object> stringObjectMap : wclcflist) {
                            wczds += Double.valueOf(stringObjectMap.get("检测单元数").toString());
                            wchgds += Double.valueOf(stringObjectMap.get("合格单元数").toString());
                        }
                    }

                    List<Map<String, Object>> ssxslist = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo,1);
                    if (ssxslist != null && ssxslist.size() > 0){
                        for (Map<String, Object> stringObjectMap : ssxslist) {
                            ssxszds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                            ssxshgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> pzdlist = jjgZdhPzdService.lookJdbjg(commonInfoVo, 1);
                    if (pzdlist != null && pzdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : pzdlist) {
                            pzdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            pzdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> mcxslist = jjgZdhMcxsService.lookJdbjg(commonInfoVo,1);
                    if (mcxslist != null && mcxslist.size() > 0){
                        for (Map<String, Object> stringObjectMap : mcxslist) {
                            mcxszds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            mcxshgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> gzsdlist = jjgZdhGzsdService.lookJdbjg(commonInfoVo, 1);
                    if (gzsdlist != null && gzsdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : gzsdlist) {
                            gzsdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            gzsdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }
                    List<Map<String, Object>> czlist = jjgZdhCzService.lookJdbjg(commonInfoVo, 1);
                    if (czlist != null && czlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : czlist) {
                            czzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            czhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> ldhdlist = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> lqhdlist = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
                    if (ldhdlist != null && ldhdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : ldhdlist) {
                            ldhdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            ldhdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }
                    if (lqhdlist != null && lqhdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : lqhdlist) {
                            ldhdzds += Double.valueOf(stringObjectMap.get("上面层厚度检测点数").toString());
                            ldhdhgds += Double.valueOf(stringObjectMap.get("上面层厚度合格点数").toString());

                            ldhdzds += Double.valueOf(stringObjectMap.get("总厚度检测点数").toString());
                            ldhdhgds += Double.valueOf(stringObjectMap.get("总厚度合格点数").toString());

                        }
                    }

                    List<Map<String, String>> lmhplist = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
                    if (lmhplist != null && lmhplist.size() > 0){
                        for (Map<String, String> stringObjectMap : lmhplist) {
                            String lmlx = stringObjectMap.get("路面类型").toString();
                            if (lmlx.contains("沥青")){
                                lqlmhpzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                                lqlmhphgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }
                            if (lmlx.contains("混凝土")){
                                hntlmhpzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                                hntlmhphgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }
                        }
                    }

                    List<Map<String, Object>> hntlmqdlist = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
                    if (hntlmqdlist != null && hntlmqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : hntlmqdlist) {
                            hntlmqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            hntlmqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> hntlmhdlist = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
                    if (hntlmhdlist != null && hntlmhdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : hntlmhdlist) {
                            hntlmhdzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                            hntlmhdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> gzsdsgpsflist = jjgFbgcLmgcLmgzsdsgpsfServicel.lookJdbjg(commonInfoVo);
                    if (gzsdsgpsflist != null && gzsdsgpsflist.size() > 0){
                        for (Map<String, Object> stringObjectMap : gzsdsgpsflist) {
                            gzsdsgpsfzds += Double.valueOf(stringObjectMap.get("检测点数").toString());
                            gzsdsgpsfhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> tlmxlbgclist = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
                    if (tlmxlbgclist != null && tlmxlbgclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : tlmxlbgclist) {
                            tlmxlbgczds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            tlmxlbgchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }


                }

                Map<String,Object> ysdmap = new HashMap<>();
                ysdmap.put("ysdzds",decf.format(ysdzds));
                ysdmap.put("ysdhgds",decf.format(ysdhgds));
                ysdmap.put("ysdhgl",ysdzds != 0 ? df.format(ysdhgds / ysdzds * 100) : 0);
                resultList.add(ysdmap);

                Map<String,Object> wcmap = new HashMap<>();
                wcmap.put("wczds",decf.format(wczds));
                wcmap.put("wchgds",decf.format(wchgds));
                wcmap.put("wchgl",wczds != 0 ? df.format(wchgds / wczds * 100) : 0);
                resultList.add(wcmap);

                Map<String,Object> ssxsmap = new HashMap<>();
                ssxsmap.put("ssxszds",decf.format(ssxszds));
                ssxsmap.put("ssxshgds",decf.format(ssxshgds));
                ssxsmap.put("ssxshgl",ssxszds != 0 ? df.format(ssxshgds / ssxszds * 100) : 0);
                resultList.add(ssxsmap);

                Map<String,Object> pzdmap = new HashMap<>();
                pzdmap.put("pzdzds",decf.format(pzdzds));
                pzdmap.put("pzdhgds",decf.format(pzdhgds));
                pzdmap.put("pzdhgl",pzdzds != 0 ? df.format(pzdhgds / pzdzds * 100) : 0);
                resultList.add(pzdmap);

                Map<String,Object> mcxsmap = new HashMap<>();
                mcxsmap.put("mcxszds",decf.format(mcxszds));
                mcxsmap.put("mcxshgds",decf.format(mcxshgds));
                mcxsmap.put("mcxshgl",mcxszds != 0 ? df.format(mcxshgds / mcxszds * 100) : 0);
                resultList.add(mcxsmap);

                Map<String,Object> gzsdmap = new HashMap<>();
                gzsdmap.put("gzsdzds",decf.format(gzsdzds));
                gzsdmap.put("gzsdhgds",decf.format(gzsdhgds));
                gzsdmap.put("gzsdhgl",gzsdzds != 0 ? df.format(gzsdhgds / gzsdzds * 100) : 0);
                resultList.add(gzsdmap);

                Map<String,Object> czmap = new HashMap<>();
                czmap.put("czzds",decf.format(czzds));
                czmap.put("czhgds",decf.format(czhgds));
                czmap.put("czhgl",czzds != 0 ? df.format(czhgds / czzds * 100) : 0);
                resultList.add(czmap);

                Map<String,Object> ldhdmap = new HashMap<>();
                ldhdmap.put("ldhdzds",decf.format(ldhdzds));
                ldhdmap.put("ldhdhgds",decf.format(ldhdhgds));
                ldhdmap.put("ldhdhgl",ldhdzds != 0 ? df.format(ldhdhgds / ldhdzds * 100) : 0);
                resultList.add(ldhdmap);

                Map<String,Object> lqhpmap = new HashMap<>();
                lqhpmap.put("lqlmhpzds",decf.format(lqlmhpzds));
                lqhpmap.put("lqlmhphgds",decf.format(lqlmhphgds));
                lqhpmap.put("lqlmhphgl",lqlmhpzds != 0 ? df.format(lqlmhphgds / lqlmhpzds * 100) : 0);
                resultList.add(lqhpmap);

                Map<String,Object> hntlmqdmap = new HashMap<>();
                hntlmqdmap.put("hntlmqdzds",decf.format(hntlmqdzds));
                hntlmqdmap.put("hntlmqdhgds",decf.format(hntlmqdhgds));
                hntlmqdmap.put("hntlmqdhgl",hntlmqdzds != 0 ? df.format(hntlmqdhgds / hntlmqdzds * 100) : 0);
                resultList.add(hntlmqdmap);

                Map<String,Object> hntlmhdmap = new HashMap<>();
                hntlmhdmap.put("hntlmhdzds",decf.format(hntlmhdzds));
                hntlmhdmap.put("hntlmhdhgds",decf.format(hntlmhdhgds));
                hntlmhdmap.put("hntlmhdhgl",hntlmhdzds != 0 ? df.format(hntlmhdhgds / hntlmhdzds * 100) : 0);
                resultList.add(hntlmhdmap);

                Map<String,Object> gzsdsgpsfmap = new HashMap<>();
                gzsdsgpsfmap.put("gzsdsgpsfzds",decf.format(gzsdsgpsfzds));
                gzsdsgpsfmap.put("gzsdsgpsfhgds",decf.format(gzsdsgpsfhgds));
                gzsdsgpsfmap.put("gzsdsgpsfhgl",gzsdsgpsfzds != 0 ? df.format(gzsdsgpsfhgds / gzsdsgpsfzds * 100) : 0);
                resultList.add(gzsdsgpsfmap);

                Map<String,Object> tlmxlbgcmap = new HashMap<>();
                tlmxlbgcmap.put("tlmxlbgczds",decf.format(tlmxlbgczds));
                tlmxlbgcmap.put("tlmxlbgchgds",decf.format(tlmxlbgchgds));
                tlmxlbgcmap.put("tlmxlbgchgl",tlmxlbgczds != 0 ? df.format(tlmxlbgchgds / tlmxlbgczds * 100) : 0);
                resultList.add(tlmxlbgcmap);

                Map<String,Object> hnthpmap = new HashMap<>();
                hnthpmap.put("hntlmhpzds",decf.format(hntlmhpzds));
                hnthpmap.put("hntlmhphgds",decf.format(hntlmhphgds));
                hnthpmap.put("hntlmhphgl",hntlmhpzds != 0 ? df.format(hntlmhphgds / hntlmhpzds * 100) : 0);
                resultList.add(hnthpmap);




            }

            if (qlhtdlist != null && qlhtdlist.size() > 0){
                double qlxbtqdzds = 0,qlxbtqdhgds = 0;
                double qlxbjgcczds = 0,qlxbjgcchgds = 0;
                double qlxbbhchdzds = 0,qlxbbhchdhgds = 0;
                double qlxbszdzds = 0,qlxbszdhgds = 0;
                double qlsbtqdzds = 0,qlsbtqdhgds = 0;
                double qlsbjgcczds = 0,qlsbjgcchgds = 0;
                double qlsbbhchdzds = 0,qlsbbhchdhgds = 0;

                for (String htd : qlhtdlist) {
                    CommonInfoVo commonInfoVo = new CommonInfoVo();
                    commonInfoVo.setProname(proname);
                    commonInfoVo.setHtd(htd);

                    List<Map<String, Object>> qlxbtqdlist = jjgFbgcQlgcXbTqdService.lookJdbjg(commonInfoVo);
                    if (qlxbtqdlist != null && qlxbtqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : qlxbtqdlist) {
                            qlxbtqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlxbtqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> qlxbjgcclist = jjgFbgcQlgcXbJgccService.lookJdbjg(commonInfoVo);
                    if (qlxbjgcclist != null && qlxbjgcclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : qlxbjgcclist) {
                            qlxbjgcczds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlxbjgcchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }
                    List<Map<String, Object>> xbbhchdlist = jjgFbgcQlgcXbBhchdService.lookJdbjg(commonInfoVo);
                    if (xbbhchdlist != null && xbbhchdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : xbbhchdlist) {
                            qlxbbhchdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlxbbhchdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> qlxbszdlist = jjgFbgcQlgcXbSzdService.lookJdbjg(commonInfoVo);
                    if (qlxbszdlist != null && qlxbszdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : qlxbszdlist) {
                            qlxbszdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlxbszdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> qlsbtqdlist = jjgFbgcQlgcSbTqdService.lookJdbjg(commonInfoVo);
                    if (qlsbtqdlist != null && qlsbtqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : qlsbtqdlist) {
                            qlsbtqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlsbtqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> qlsbjgcclist = jjgFbgcQlgcSbJgccService.lookJdbjg(commonInfoVo);
                    if (qlsbjgcclist != null && qlsbjgcclist.size() > 0){
                        for (Map<String, Object> stringObjectMap : qlsbjgcclist) {
                            qlsbjgcczds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlsbjgcchgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> qlsbbhchdlist = jjgFbgcQlgcSbBhchdService.lookJdbjg(commonInfoVo);
                    if (qlsbbhchdlist != null && qlsbbhchdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : qlsbbhchdlist) {
                            qlsbbhchdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            qlsbbhchdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }



                }

                Map<String,Object> qlxbtqdmap = new HashMap<>();
                qlxbtqdmap.put("qlxbtqdzds",decf.format(qlxbtqdzds));
                qlxbtqdmap.put("qlxbtqdhgds",decf.format(qlxbtqdhgds));
                qlxbtqdmap.put("qlxbtqdhgl",qlxbtqdzds != 0 ? df.format(qlxbtqdhgds / qlxbtqdzds * 100) : 0);
                resultList.add(qlxbtqdmap);

                Map<String,Object> qlxbjgccmap = new HashMap<>();
                qlxbjgccmap.put("qlxbjgcczds",decf.format(qlxbjgcczds));
                qlxbjgccmap.put("qlxbjgcchgds",decf.format(qlxbjgcchgds));
                qlxbjgccmap.put("qlxbjgcchgl",qlxbjgcczds != 0 ? df.format(qlxbjgcchgds / qlxbjgcczds * 100) : 0);
                resultList.add(qlxbjgccmap);

                Map<String,Object> qlxbbhchdmap = new HashMap<>();
                qlxbbhchdmap.put("qlxbbhchdzds",decf.format(qlxbbhchdzds));
                qlxbbhchdmap.put("qlxbbhchdhgds",decf.format(qlxbbhchdhgds));
                qlxbbhchdmap.put("qlxbbhchdhgl",qlxbbhchdzds != 0 ? df.format(qlxbbhchdhgds / qlxbbhchdzds * 100) : 0);
                resultList.add(qlxbbhchdmap);

                Map<String,Object> qlxbszdmap = new HashMap<>();
                qlxbszdmap.put("qlxbszdzds",decf.format(qlxbszdzds));
                qlxbszdmap.put("qlxbszdhgds",decf.format(qlxbszdhgds));
                qlxbszdmap.put("qlxbszdhgl",qlxbszdzds != 0 ? df.format(qlxbszdhgds / qlxbszdzds * 100) : 0);
                resultList.add(qlxbszdmap);

                Map<String,Object> qlsbtqdmap = new HashMap<>();
                qlsbtqdmap.put("qlsbtqdzds",decf.format(qlsbtqdzds));
                qlsbtqdmap.put("qlsbtqdhgds",decf.format(qlsbtqdhgds));
                qlsbtqdmap.put("qlsbtqdhgl",qlsbtqdzds != 0 ? df.format(qlsbtqdhgds / qlsbtqdzds * 100) : 0);
                resultList.add(qlsbtqdmap);

                Map<String,Object> qlsbjgccmap = new HashMap<>();
                qlsbjgccmap.put("qlsbjgcczds",decf.format(qlsbjgcczds));
                qlsbjgccmap.put("qlsbjgcchgds",decf.format(qlsbjgcchgds));
                qlsbjgccmap.put("qlsbjgcchgl",qlsbjgcczds != 0 ? df.format(qlsbjgcchgds / qlsbjgcczds * 100) : 0);
                resultList.add(qlsbjgccmap);

                Map<String,Object> qlsbbhchdmap = new HashMap<>();
                qlsbbhchdmap.put("qlsbbhchdzds",decf.format(qlsbbhchdzds));
                qlsbbhchdmap.put("qlsbbhchdhgds",decf.format(qlsbbhchdhgds));
                qlsbbhchdmap.put("qlsbbhchdhgl",qlsbbhchdzds != 0 ? df.format(qlsbbhchdhgds / qlsbbhchdzds * 100) : 0);
                resultList.add(qlsbbhchdmap);

            }


            if (sdhtdlist != null && sdhtdlist.size() > 0){
                double sdgccqtqdzds = 0,sdgccqtqdhgds = 0;
                double sdgccqhdzds = 0,sdgccqhdhgds = 0;
                double sdgcdmpzdzds = 0,sdgcdmpzdhgds = 0;
                double sdgcztkdzds = 0,sdgcztkdhgds = 0;
                double sdgcjkzds = 0,sdgcjkhgds = 0;

                for (String htd : sdhtdlist) {
                    CommonInfoVo commonInfoVo = new CommonInfoVo();
                    commonInfoVo.setProname(proname);
                    commonInfoVo.setHtd(htd);

                    List<Map<String, Object>> sdgccqtqdlist = jjgFbgcSdgcCqtqdService.lookJdbjg(commonInfoVo);
                    if (sdgccqtqdlist != null && sdgccqtqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : sdgccqtqdlist) {
                            sdgccqtqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            sdgccqtqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> sdgccqhdlist = jjgFbgcSdgcCqhdService.lookJdbjg(commonInfoVo);
                    if (sdgccqhdlist != null && sdgccqhdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : sdgccqhdlist) {
                            sdgccqhdzds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            sdgccqhdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> sddmpzdlist = jjgFbgcSdgcDmpzdService.lookJdbjg(commonInfoVo);
                    if (sddmpzdlist != null && sddmpzdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : sddmpzdlist) {
                            sdgcdmpzdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            sdgcdmpzdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> sdgcztkdlist = jjgFbgcSdgcZtkdService.lookJdbjg(commonInfoVo);
                    if (sdgcztkdlist != null && sdgcztkdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : sdgcztkdlist) {
                            sdgcztkdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            sdgcztkdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> sdgcjklist = jjgFbgcSdgcJkService.lookJdbjg(commonInfoVo);
                    if (sdgcjklist != null && sdgcjklist.size() > 0){
                        for (Map<String, Object> stringObjectMap : sdgcjklist) {
                            sdgcjkzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            sdgcjkhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }


                }

                Map<String,Object> sdgccqtqdmap = new HashMap<>();
                sdgccqtqdmap.put("sdgccqtqdzds",decf.format(sdgccqtqdzds));
                sdgccqtqdmap.put("sdgccqtqdhgds",decf.format(sdgccqtqdhgds));
                sdgccqtqdmap.put("sdgccqtqdhgl",sdgccqtqdzds != 0 ? df.format(sdgccqtqdhgds / sdgccqtqdzds * 100) : 0);
                resultList.add(sdgccqtqdmap);

                Map<String,Object> sdgccqhdmap = new HashMap<>();
                sdgccqhdmap.put("sdgccqhdzds",decf.format(sdgccqhdzds));
                sdgccqhdmap.put("sdgccqhdhgds",decf.format(sdgccqhdhgds));
                sdgccqhdmap.put("sdgccqhdhgl",sdgccqhdzds != 0 ? df.format(sdgccqhdhgds / sdgccqhdzds * 100) : 0);
                resultList.add(sdgccqhdmap);

                Map<String,Object> sddmpzdmap = new HashMap<>();
                sddmpzdmap.put("sdgcdmpzdzds",decf.format(sdgcdmpzdzds));
                sddmpzdmap.put("sdgcdmpzdhgds",decf.format(sdgcdmpzdhgds));
                sddmpzdmap.put("sdgcdmpzdhgl",sdgcdmpzdzds != 0 ? df.format(sdgcdmpzdhgds / sdgcdmpzdzds * 100) : 0);
                resultList.add(sddmpzdmap);

                Map<String,Object> sdgcztkdmap = new HashMap<>();
                sdgcztkdmap.put("sdgcztkdzds",decf.format(sdgcztkdzds));
                sdgcztkdmap.put("sdgcztkdhgds",decf.format(sdgcztkdhgds));
                sdgcztkdmap.put("sdgcztkdhgl",sdgcztkdzds != 0 ? df.format(sdgcztkdhgds / sdgcztkdzds * 100) : 0);
                resultList.add(sdgcztkdmap);

                Map<String,Object> sdgcjkmap = new HashMap<>();
                sdgcjkmap.put("sdgcjkzds",decf.format(sdgcjkzds));
                sdgcjkmap.put("sdgcjkhgds",decf.format(sdgcjkhgds));
                sdgcjkmap.put("sdgcjkhgl",sdgcjkzds != 0 ? df.format(sdgcjkhgds / sdgcjkzds * 100) : 0);
                resultList.add(sdgcjkmap);

            }


            if (jahtdlist != null && jahtdlist.size() > 0){
                double lzszdzds = 0,lzszdhgds = 0;
                double bzbjkzds = 0,bzbjkhgds = 0;
                double bzbhdzds = 0,bzbhdhgds = 0;
                double fsxszds = 0,fsxshgds = 0;
                double jabxhdzds = 0,jabxhdhgds = 0;
                double jafsxszds = 0,jafsxshgds = 0;

                double jshdzds = 0,jshdhgds = 0;
                double lzbhzds = 0,lzbhhgds = 0;
                double mrsdzds = 0,mrsdhgds = 0;
                double zxgdzds = 0,zxgdhgds = 0;

                double jathlzds = 0,jathlhgds = 0;
                double thlqdzds = 0,thlqdhgds = 0;

                for (String htd : jahtdlist) {
                    CommonInfoVo commonInfoVo = new CommonInfoVo();
                    commonInfoVo.setProname(proname);
                    commonInfoVo.setHtd(htd);
                    List<Map<String, Object>> jabzlist = jjgFbgcJtaqssJabzService.lookJdbjg(commonInfoVo);
                    if (jabzlist != null && jabzlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : jabzlist) {
                            String xm = stringObjectMap.get("项目").toString();

                            if (xm.contains("立柱竖直度")){
                                lzszdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                lzszdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (xm.contains("标志板净空")){
                                bzbjkzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                bzbjkhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (xm.contains("标志板厚度")){
                                bzbhdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                bzbhdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (xm.contains("标志面反光膜逆反射系数")){
                                fsxszds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                fsxshgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }
                        }

                    }
                    List<Map<String, Object>> jabxlist = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
                    if (jabxlist != null && jabxlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : jabxlist) {
                            String jcxm = stringObjectMap.get("检测项目").toString();
                            if (jcxm.contains("交安标线厚度")){
                                jabxhdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                jabxhdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (jcxm.contains("交安标线逆反射系数")){
                                jafsxszds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                jafsxshgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                        }
                        
                    }

                    List<Map<String, Object>> jafhllist = jjgFbgcJtaqssJabxfhlService.lookJdbjg(commonInfoVo);
                    /**
                     * [{不合格点数=11, 合格点数=11, 总点数=15, 检测项目=波形梁板基底金属厚度, 合格率=73.33, 规定值或允许偏差=≥4.0≥3.0},
                     * {不合格点数=6, 合格点数=6, 总点数=15, 检测项目=波形梁钢护栏立柱壁厚, 合格率=40.00, 规定值或允许偏差=方柱≥6.0圆柱≥4.5},
                     * {不合格点数=3, 合格点数=12, 总点数=15, 检测项目=波形梁钢护栏横梁中心高度, 合格率=80.00, 规定值或允许偏差=三波板697±20},
                     * {不合格点数=3, 合格点数=3, 总点数=3, 检测项目=波形梁钢护栏立柱埋入深度, 合格率=100.00, 规定值或允许偏差=方柱≥1400.0圆柱≥1400.0}]
                     */
                    if (jafhllist != null && jafhllist.size() > 0){
                        for (Map<String, Object> stringObjectMap : jafhllist) {
                            String jcxm = stringObjectMap.get("检测项目").toString();
                            if (jcxm.contains("波形梁板基底金属厚度")){
                                jshdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                jshdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (jcxm.contains("波形梁钢护栏立柱壁厚")){
                                lzbhzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                lzbhhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (jcxm.contains("波形梁钢护栏横梁中心高度")){
                                mrsdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                mrsdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                            if (jcxm.contains("波形梁钢护栏立柱埋入深度")){
                                zxgdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                                zxgdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                            }

                        }

                    }

                    List<Map<String, Object>> jathllist = jjgFbgcJtaqssJathldmccService.lookJdbjg(commonInfoVo);
                    if (jathllist != null && jathllist.size() > 0){
                        for (Map<String, Object> stringObjectMap : jathllist) {
                            jathlzds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                            jathlhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }

                    List<Map<String, Object>> thlqdlist = jjgFbgcJtaqssJathlqdService.lookJdbjg(commonInfoVo);
                    if (thlqdlist != null && thlqdlist.size() > 0){
                        for (Map<String, Object> stringObjectMap : thlqdlist) {
                            thlqdzds += Double.valueOf(stringObjectMap.get("总点数").toString());
                            thlqdhgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
                        }
                    }


                }

                Map<String,Object> lzszdmap = new HashMap<>();
                lzszdmap.put("lzszdzds",decf.format(lzszdzds));
                lzszdmap.put("lzszdhgds",decf.format(lzszdhgds));
                lzszdmap.put("lzszdhgl",lzszdzds != 0 ? df.format(lzszdhgds / lzszdzds * 100) : 0);
                resultList.add(lzszdmap);
                Map<String,Object> bzbjkmap = new HashMap<>();
                bzbjkmap.put("bzbjkzds",decf.format(bzbjkzds));
                bzbjkmap.put("bzbjkhgds",decf.format(bzbjkhgds));
                bzbjkmap.put("bzbjkhgl",bzbjkzds != 0 ? df.format(bzbjkhgds / bzbjkzds * 100) : 0);
                resultList.add(bzbjkmap);
                Map<String,Object> bzbhdmap = new HashMap<>();
                bzbhdmap.put("bzbhdzds",decf.format(bzbhdzds));
                bzbhdmap.put("bzbhdhgds",decf.format(bzbhdhgds));
                bzbhdmap.put("bzbhdhgl",bzbhdzds != 0 ? df.format(bzbhdhgds / bzbhdzds * 100) : 0);
                resultList.add(bzbhdmap);
                Map<String,Object> fsxsmap = new HashMap<>();
                fsxsmap.put("fsxszds",decf.format(fsxszds));
                fsxsmap.put("fsxshgds",decf.format(fsxshgds));
                fsxsmap.put("fsxshgl",fsxszds != 0 ? df.format(fsxshgds / fsxszds * 100) : 0);
                resultList.add(fsxsmap);
                Map<String,Object> bxhdmap = new HashMap<>();
                bxhdmap.put("jabxhdzds",decf.format(jabxhdzds));
                bxhdmap.put("jabxhdhgds",decf.format(jabxhdhgds));
                bxhdmap.put("jabxhdhgl",jabxhdzds != 0 ? df.format(jabxhdhgds / jabxhdzds * 100) : 0);
                resultList.add(bxhdmap);
                Map<String,Object> jafsxsmap = new HashMap<>();
                jafsxsmap.put("jafsxszds",decf.format(jafsxszds));
                jafsxsmap.put("jafsxshgds",decf.format(jafsxshgds));
                jafsxsmap.put("jafsxshgl",jafsxszds != 0 ? df.format(jafsxshgds / jafsxszds * 100) : 0);
                resultList.add(jafsxsmap);


                Map<String,Object> jshdmap = new HashMap<>();
                jshdmap.put("jshdzds",decf.format(jshdzds));
                jshdmap.put("jshdhgds",decf.format(jshdhgds));
                jshdmap.put("jshdhgl",jshdzds != 0 ? df.format(jshdhgds / jshdzds * 100) : 0);
                resultList.add(jshdmap);

                Map<String,Object> lzbhmap = new HashMap<>();
                lzbhmap.put("lzbhzds",decf.format(lzbhzds));
                lzbhmap.put("lzbhhgds",decf.format(lzbhhgds));
                lzbhmap.put("lzbhhgl",lzbhzds != 0 ? df.format(lzbhhgds / lzbhzds * 100) : 0);
                resultList.add(lzbhmap);

                Map<String,Object> mrsdmap = new HashMap<>();
                mrsdmap.put("mrsdzds",decf.format(mrsdzds));
                mrsdmap.put("mrsdhgds",decf.format(mrsdhgds));
                mrsdmap.put("mrsdhgl",mrsdzds != 0 ? df.format(mrsdhgds / mrsdzds * 100) : 0);
                resultList.add(mrsdmap);

                Map<String,Object> zxgdmap = new HashMap<>();
                zxgdmap.put("zxgdzds",decf.format(zxgdzds));
                zxgdmap.put("zxgdhgds",decf.format(zxgdhgds));
                zxgdmap.put("zxgdhgl",zxgdzds != 0 ? df.format(zxgdhgds / zxgdzds * 100) : 0);
                resultList.add(zxgdmap);

                Map<String,Object> jathlmap = new HashMap<>();
                jathlmap.put("jathlzds",decf.format(jathlzds));
                jathlmap.put("jathlhgds",decf.format(jathlhgds));
                jathlmap.put("jathlhgl",jathlzds != 0 ? df.format(jathlhgds / jathlzds * 100) : 0);
                resultList.add(jathlmap);

                Map<String,Object> thlqdmap = new HashMap<>();
                thlqdmap.put("thlqdzds",decf.format(thlqdzds));
                thlqdmap.put("thlqdhgds",decf.format(thlqdhgds));
                thlqdmap.put("thlqdhgl",thlqdzds != 0 ? df.format(thlqdhgds / thlqdzds * 100) : 0);
                resultList.add(thlqdmap);

            }

        }
        return resultList;

    }

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablehtdpdb(File f, Document xw, String proname) throws IOException {
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjgJgHtdinfo : list) {
                String htd = jjgJgHtdinfo.getName();
                File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                if (ff.exists()){
                    Workbook workbook = new Workbook();
                    workbook.loadFromFile(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    WorksheetsCollection sheetnum = workbook.getWorksheets();
                    for (int j = 0; j < sheetnum.getCount() ; j++) {
                        String sheetname = sheetnum.get(j).getName();
                        Worksheet sheet = workbook.getWorksheets().get(j);
                        boolean ishidden = JjgFbgcCommonUtils.ishidden(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx", sheetname);
                        if (!ishidden) {
                            CellRange allocatedRange = sheet.getAllocatedRange();
                            //TextSelection[] selection = xw.findAllString("${附表21评定表}", true, true);
                            TextSelection[] selection = findStringInPagess(xw, "${附表21评定表}", 60, 90);
                            if (selection != null) {
                                for (int k = 0; k < selection.length; k++) {
                                    TextRange range = selection[k].getAsOneRange();
                                    //删除没用数据的行
                                    CellRange dataRange = RemoveEmptyRows(sheet, allocatedRange);
                                    //复制到Word文档
                                    copyToWord1(dataRange, xw, f.getPath(), range);
                                }
                            }
                        }
                    }
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
    private void CreateTablejsxmzlpdb(File f, Document xw, String proname) throws IOException {
        File ff = new File(filespath + File.separator + proname + File.separator + "建设项目质量评定表.xlsx");
        if (ff.exists()) {

            Workbook workbook = new Workbook();
            workbook.loadFromFile(filespath + File.separator + proname + File.separator + "建设项目质量评定表.xlsx");
            Worksheet sheet = workbook.getWorksheets().get("建设项目");

            /*XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
            XSSFSheet slSheet = xwb.getSheet("建设项目");*/
            TextSelection[] selection = xw.findAllString("${附表21建设项目质量评定表}", true, true);

            CellRange allocatedRange = sheet.getAllocatedRange();

            if (selection != null) {
                for (int k = 0; k < selection.length; k++) {
                    TextRange range = selection[k].getAsOneRange();
                    //删除没用数据的行
                    CellRange dataRange = RemoveEmptyRows(sheet, allocatedRange);
                    //复制到Word文档
                    copyToWord(dataRange, xw, f.getPath(), range);
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
    private void CreateTablejsxmpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getjsxmpdData(proname);
        if (data !=null && data.size()>0) {
            String[] header = {"合同段", "实得分", "投资额（万元）", "实得分×投资额", "质量等级", "备注"};
            String s = "${表8.3-1}";
            TextSelection textSelection = findStringInPages(xw, s, 35, 65);
            if (textSelection != null) {
                com.spire.doc.Table table = new Table(xw, true);
                table.resetCells(data.size() + 1, 6);
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
                    txtRange.getCharacterFormat().setBold(true);

                }

                //添加数据
                for (int rowIdx = 0; rowIdx < data.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    Map<String, Object> rowData = data.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row2.getCells().get(colIdx).getCellFormat().setVerticalAlignment(com.spire.doc.documents.VerticalAlignment.Middle);
                        Paragraph p = row2.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);


                        // 根据表头的不同，设置相应的数据 2 7 12
                        switch (colIdx) {
                            case 0:
                                p.appendText(rowData.get("htd").toString() + "合同段").getCharacterFormat().setFontSize(9f);
                                break;
                            case 1:
                                p.appendText(rowData.get("sdf").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.get("tze").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.get("sdftze").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.get("zldj").toString()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.get("bz").toString()).getCharacterFormat().setFontSize(9f);
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
    private List<Map<String, Object>> getjsxmpdData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        File ff = new File(filespath + File.separator + proname + File.separator + "建设项目质量评定表.xlsx");
        if (!ff.exists()) {
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
            XSSFSheet slSheet = xwb.getSheet("建设项目");
            for (int i = 0; i < slSheet.getPhysicalNumberOfRows(); i++){
                slSheet.getRow(i).getCell(0).setCellType(CellType.STRING);
                String value1 = slSheet.getRow(i).getCell(0).getStringCellValue();
                if (!value1.equals("建设项目质量评定表") && !value1.equals("项目名称：") && !value1.equals("起讫桩号：") && !value1.equals("合同段") && !value1.equals("建设项目质量等级") && !value1.equals("")){
                    slSheet.getRow(i).getCell(0).setCellType(CellType.STRING);
                    slSheet.getRow(i).getCell(3).setCellType(CellType.STRING);
                    slSheet.getRow(i).getCell(6).setCellType(CellType.STRING);

                    String value = slSheet.getRow(i).getCell(0).getStringCellValue();
                    String value2 = slSheet.getRow(i).getCell(3).getStringCellValue();
                    String value3 = slSheet.getRow(i).getCell(6).getStringCellValue();

                    Map<String,Object> map = new HashMap<>();

                    map.put("htd",value);
                    map.put("sdf","\\");
                    map.put("tze","\\");
                    map.put("sdftze","\\");
                    map.put("zldj",value2);
                    map.put("bz",value3);
                    resultList.add(map);

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
    private void CreateTablehtdzlpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = gethtdzlpdData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "", "", "", "", "单位工程", "", "", ""};
            String[] header2 = {"名称", "评分", "资料扣分", "质量得分", "质量等级", "名称", "得分", "投资额（万元）", "质量等级"};
            String s = "${表8.2-1}";
            TextSelection textSelection = findStringInPages(xw, s, 35, 65);
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
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjgJgHtdinfo : list) {
                String htd = jjgJgHtdinfo.getName();
                File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                if (!ff.exists()) {
                    break;
                }else {
                    XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                    XSSFSheet slSheet1 = xwb.getSheet("合同段");
                    for (int i = 0; i < slSheet1.getPhysicalNumberOfRows(); i++){
                        CellType cellTypeEnum = slSheet1.getRow(i).getCell(0).getCellTypeEnum();


                        slSheet1.getRow(i).getCell(1).setCellType(CellType.STRING);//合同段
                        slSheet1.getRow(i).getCell(5).setCellType(CellType.STRING);//项目名称

                        String bzhtd = slSheet1.getRow(i).getCell(1).getStringCellValue();
                        String bzxmmc = slSheet1.getRow(i).getCell(5).getStringCellValue();

                        if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                            Map<String,Object> map = new HashMap<>();
                            slSheet1.getRow(i+25).getCell(1).setCellType(CellType.STRING);
                            slSheet1.getRow(i+25).getCell(6).setCellType(CellType.STRING);
                            slSheet1.getRow(i+26).getCell(4).setCellType(CellType.STRING);
                            slSheet1.getRow(i+26).getCell(6).setCellType(CellType.STRING);

                            map.put("htdpf","\\");
                            map.put("zlkf","\\");
                            map.put("zldf","\\");
                            map.put("htdzldj",slSheet1.getRow(i+25).getCell(1).getStringCellValue());
                            resultList1.add(map);
                        }
                        slSheet1.getRow(i).getCell(0).getStringCellValue();
                        String string = slSheet1.getRow(i).getCell(0).toString();
                        if (string.equals(htd)){
                            String value1 = slSheet1.getRow(i).getCell(1).getStringCellValue();

                            String value4 = slSheet1.getRow(i).getCell(4).getStringCellValue();
                            Map<String,Object> map = new HashMap<>();
                            map.put("htd",htd);
                            map.put("dwgc",value1);
                            map.put("df","\\");
                            map.put("tze","\\");
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
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablejazlpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getjapdbData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "单位工程", "", "", "分部工程", "", "", "", "", ""};
            String[] header2 = {"", "名称", "得分", "质量等级", "名称", "权值", "得分", "实测得分", "外观扣分", "质量等级"};
            String s = "${表8.1.5-1}";
            TextSelection textSelection = findStringInPages(xw, s, 35, 60);
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
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("交安工程")){
                    File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
                    if (!ff.exists()) {
                        break;
                    }else {
                        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(ff));
                        XSSFSheet slSheet1 = xwb.getSheet("分部-交安");
                        for (int i = 0; i < slSheet1.getPhysicalNumberOfRows()-20; i++){
                            slSheet1.getRow(i).getCell(2).setCellType(CellType.STRING);//合同段
                            slSheet1.getRow(i).getCell(15).setCellType(CellType.STRING);//项目名称
                            slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);//分部工程名称

                            String bzhtd = slSheet1.getRow(i).getCell(2).getStringCellValue();
                            String bzxmmc = slSheet1.getRow(i).getCell(15).getStringCellValue();

                            if (bzhtd.equals(htd) && bzxmmc.equals(proname)){
                                slSheet1.getRow(i+20).getCell(6).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(10).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(16).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                String zldj = "";
                                String value1 = slSheet1.getRow(i + 20).getCell(16).getStringCellValue();
                                if (value1.equals("√")){
                                    zldj = "合格";
                                }else {
                                    zldj = "不合格";
                                }

                                String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","交通安全设施");
                                map.put("fbgc",fbgcname);
                                map.put("scdf","\\");
                                map.put("wgkf","\\");
                                map.put("df","\\");
                                map.put("zldj",zldj);
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
                            String dwvalue = slSheet2.getRow(i).getCell(0).getStringCellValue();
                            if (dwvalue.equals("单位工程名称：") && value.equals("交安")){
                                slSheet2.getRow(i + 24).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 24).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 24).getCell(1).getStringCellValue();
                                //String value2 = slSheet2.getRow(i + 24).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","交通安全设施");
                                map.put("dwgcdf","\\");
                                map.put("dwgczldj",value1);
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
    private void CreateTablesdzlpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getsdpdbData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "单位工程", "", "", "分部工程", "", "", "", "", ""};
            String[] header2 = {"", "名称", "得分", "质量等级", "名称", "权值", "得分", "实测得分", "外观扣分", "质量等级"};
            String s = "${表8.1.4-1}";
            TextSelection textSelection = findStringInPages(xw,s,35,60);
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
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjgJgHtdinfo : list) {
                String lx = jjgJgHtdinfo.getLx();
                String htd = jjgJgHtdinfo.getName();
                if (lx.contains("隧道工程")){
                    File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
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
                                        slSheet1.getRow(i+20).getCell(16).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                        String zldj = "";
                                        String value1 = slSheet1.getRow(i + 20).getCell(16).getStringCellValue();
                                        if (value1.equals("√")){
                                            zldj = "合格";
                                        }else {
                                            zldj = "不合格";
                                        }

                                        String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                        if (!fbgcname.equals("隧道路面")){
                                            Map<String,Object> map = new HashMap<>();
                                            map.put("htd",htd);
                                            map.put("dwgc",sheetName.substring(3));
                                            map.put("fbgc",fbgcname);
                                            map.put("scdf","\\");
                                            map.put("wgkf","\\");
                                            map.put("df","\\");
                                            map.put("zldj",zldj);
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
                                slSheet2.getRow(i + 24).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 24).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 24).getCell(1).getStringCellValue();
                                //String value2 = slSheet2.getRow(i + 24).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc",value);
                                map.put("dwgcdf","\\");
                                map.put("dwgczldj",value1);
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
    private void CreateTableqlzlpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getqlpdbData(proname);
        if (data !=null && data.size()>0) {
            String[] header1 = {"合同段", "单位工程", "", "", "分部工程", "", "", "", "", ""};
            String[] header2 = {"", "名称", "得分", "质量等级", "名称", "权值", "得分", "实测得分", "外观扣分", "质量等级"};
            String s = "${表8.1.3-1}";
            TextSelection textSelection = findStringInPages(xw,s,35,60);
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
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjghtd : list) {
                String lx = jjghtd.getLx();
                String htd = jjghtd.getName();
                if (lx.contains("桥梁工程")){
                    File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
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
                                        slSheet1.getRow(i+20).getCell(16).setCellType(CellType.STRING);
                                        slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                        String zldj = "";
                                        String value1 = slSheet1.getRow(i + 20).getCell(16).getStringCellValue();
                                        if (value1.equals("√")){
                                            zldj = "合格";
                                        }else {
                                            zldj = "不合格";
                                        }
                                        String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                        if (!fbgcname.equals("桥面系")){
                                            Map<String,Object> map = new HashMap<>();
                                            map.put("htd",htd);
                                            map.put("dwgc",sheetName.substring(3));
                                            map.put("fbgc",fbgcname);
                                            map.put("scdf","\\");
                                            map.put("wgkf","\\");
                                            map.put("df","\\");
                                            map.put("zldj",zldj);
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
                                slSheet2.getRow(i + 24).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 24).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 24).getCell(1).getStringCellValue();
                                //String value2 = slSheet2.getRow(i + 24).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc",value);
                                map.put("dwgcdf","\\");
                                map.put("dwgczldj",value1);
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
    private void CreateTablelmzlpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getlmpdbData(proname);
        if (data !=null && data.size()>0){
            String[] header1 = {"合同段", "单位工程","","","分部工程","","","","",""};
            String[] header2 = {"","名称", "得分","质量等级","名称","权值","得分","实测得分","外观扣分","质量等级"};
            String s = "${表8.1.2-1}";
            TextSelection textSelection = findStringInPages(xw,s,35,60);
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
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjghtd : list) {
                String lx = jjghtd.getLx();
                String htd = jjghtd.getName();
                if (lx.contains("路面工程")){
                    File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
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
                                slSheet1.getRow(i+20).getCell(16).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);

                                String zldj = "";
                                String value1 = slSheet1.getRow(i + 20).getCell(16).getStringCellValue();
                                if (value1.equals("√")){
                                    zldj = "合格";
                                }else {
                                    zldj = "不合格";
                                }

                                String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路面工程");
                                map.put("fbgc",fbgcname);
                                map.put("scdf","\\");
                                map.put("wgkf","\\");
                                map.put("df","\\");
                                map.put("zldj",zldj);
                                map.put("qz",1);
                                resultList.add(map);
                            }
                        }
                        XSSFSheet slSheet2 = xwb.getSheet("单位工程");
                        for (int i = 0; i < slSheet2.getPhysicalNumberOfRows()-26; i++){
                            slSheet2.getRow(i).getCell(1).setCellType(CellType.STRING);

                            String valuemc = slSheet2.getRow(i).getCell(0).getStringCellValue();
                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (valuemc.equals("单位工程名称：") && value.equals("路面")){
                                slSheet2.getRow(i + 24).getCell(1).setCellType(CellType.STRING);
                                slSheet2.getRow(i + 24).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 24).getCell(1).getStringCellValue();
                                //String value2 = slSheet2.getRow(i + 25).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路面工程");
                                map.put("dwgcdf","\\");
                                map.put("dwgczldj",value1);
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
    private void CreateTableljzlpd(File f, Document xw, String proname) throws IOException {
        List<Map<String, Object>> data = getljpdbData(proname);
        if (data !=null && data.size()>0){
            String[] header1 = {"合同段", "单位工程","","","分部工程","","","","",""};
            String[] header2 = {"","名称", "得分","质量等级","名称","权值","得分","实测得分","外观扣分","质量等级"};
            String s = "${表8.1.1-1}";
            TextSelection textSelection = findStringInPages(xw,s,35,60);
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

    private List<Map<String, Object>> getljpdbData(String proname) throws IOException {
        List<Map<String,Object>> resultList = new ArrayList<>();
        List<Map<String,Object>> resultList1 = new ArrayList<>();
        //先查合同段
        QueryWrapper<JjgHtd> htdinfo = new QueryWrapper<>();
        htdinfo.eq("proname", proname);
        List<JjgHtd> list = jjgHtdService.list(htdinfo);
        if (list != null && list.size() > 0){
            for (JjgHtd jjghtd : list) {
                String lx = jjghtd.getLx();
                String htd = jjghtd.getName();
                if (lx.contains("路基工程")){
                    File ff = new File(filespath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
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


                                slSheet1.getRow(i+20).getCell(16).setCellType(CellType.STRING);
                                slSheet1.getRow(i+20).getCell(20).setCellType(CellType.STRING);
                                String zldj = "";
                                String value1 = slSheet1.getRow(i + 20).getCell(16).getStringCellValue();
                                if (value1.equals("√")){
                                    zldj = "合格";
                                }else {
                                    zldj = "不合格";
                                }

                                String fbgcname = slSheet1.getRow(i).getCell(8).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路基工程");
                                map.put("fbgc",fbgcname);
                                map.put("scdf","\\");
                                map.put("wgkf","\\");
                                map.put("df","\\");
                                map.put("zldj",zldj);
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

                            String valuemc = slSheet2.getRow(i).getCell(0).getStringCellValue();
                            String value = slSheet2.getRow(i).getCell(1).getStringCellValue();
                            if (valuemc.equals("单位工程名称：") && value.contains("路基")){
                                slSheet2.getRow(i + 24).getCell(1).setCellType(CellType.STRING);
                                //slSheet2.getRow(i + 24).getCell(5).setCellType(CellType.STRING);

                                String value1 = slSheet2.getRow(i + 24).getCell(1).getStringCellValue();
                                //String value2 = slSheet2.getRow(i + 24).getCell(5).getStringCellValue();
                                Map<String,Object> map = new HashMap<>();
                                map.put("htd",htd);
                                map.put("dwgc","路基工程");
                                map.put("dwgcdf","\\");
                                map.put("dwgczldj",value1);
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
    private void CreateTableryqzbData(File f, Document xw, String proname) {
        QueryWrapper<JjgRyinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgRyinfo> list = jjgRyinfoService.list(wrapper);
        JjgRyinfo jgfc1 = new JjgRyinfo();
        jgfc1.setZw("报告编写人");
        JjgRyinfo jgfc2 = new JjgRyinfo();
        jgfc2.setZw("报告审核人");
        JjgRyinfo jgfc3 = new JjgRyinfo();
        jgfc3.setZw("报告签发人");
        list.add(jgfc1);
        list.add(jgfc2);
        list.add(jgfc3);

        if (list != null){
            String[] header = {"岗 位", "姓 名", "执业资格证书编号", "职 称", "签  字"};
            String s = "{签字表}";
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
                    JjgRyinfo rowData = list.get(rowIdx);


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
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablesdqdData(File f, Document xw, String proname) {
        QueryWrapper<JjgLqsSd> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd","name");
        List<JjgLqsSd> list = jjgLqsSdService.list(wrapper);
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

                    JjgLqsSd rowData = list.get(rowIdx);

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


    /**
     * 桥梁清单
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTableqlqdData(File f, Document xw, String proname) {
        QueryWrapper<JjgLqsQl> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd","qlname");
        List<JjgLqsQl> list = jjgLqsQlService.list(wrapper);
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

                    JjgLqsQl rowData = list.get(rowIdx);

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

    @Autowired
    private JjgRyinfoService jjgRyinfoService;

    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTableryData(File f, Document xw, String proname) {
        QueryWrapper<JjgRyinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgRyinfo> list = jjgRyinfoService.list(wrapper);
        if (list != null){
            String[] header = {"序号", "姓 名", "本项目承担职务", "技术职称", "执业资格证书编号","备注"};
            String s = "{表2.2-1 人员信息}";
            //TextSelection textSelection = xw.findString(s, true, true);
            TextSelection textSelection = findStringInPages(xw,s,5,10);
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

                    JjgRyinfo rowData = list.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
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

    @Autowired
    private JjgYqinfoService jjgYqinfoService;

    /**
     * 仪器信息
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTableyqData(File f, Document xw, String proname) {
        QueryWrapper<JjgYqinfo> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgYqinfo> list = jjgYqinfoService.list(wrapper);
        if (list != null){
            String[] header = {"序号", "仪器名称", "数量", "管理编号", "检测项目"};

            String s = "{表4.1-2 仪器信息}";
            //TextSelection textSelection = xw.findString(s, true, true);
            TextSelection textSelection = findStringInPages(xw,s,10,20);

            // 检查是否找到字符串
            if (textSelection != null)
            {
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
                /*float[] floatArray = {1.4f, 4.3f, 1.9f,3.7f,4.4f};
                table.setColumnWidth(floatArray);*/
                for (int i = 0; i < header.length; i++) {
                    //row.getCells().get(i).setWidth(400);
                    row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row.getCells().get(i).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    TextRange txtRange = p.appendText(header[i]);
                    txtRange.getCharacterFormat().setFontName("宋体"); // 设置字体为宋体
                    txtRange.getCharacterFormat().setFontSize(9f); // 设置字号为小四（对应12磅）
                    //txtRange.getCharacterFormat().setBold(true);
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < list.size(); rowIdx++) {
                    TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row1.setHeightType(TableRowHeightType.Exactly);

                    JjgYqinfo rowData = list.get(rowIdx);

                    for (int colIdx = 0; colIdx < header.length; colIdx++) {
                        row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                        Paragraph p = row1.getCells().get(colIdx).addParagraph();
                        p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                        // 根据表头的不同，设置相应的数据
                        switch (colIdx) {
                            case 0:
                                p.appendText(String.valueOf(rowIdx + 1)).getCharacterFormat().setFontSize(9f);
                                break;
                            case 1:
                                p.appendText(rowData.getName()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.getNum()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.getGlbh()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.getJcxm()).getCharacterFormat().setFontSize(9f);
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
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablewgjclwzms(File f, Document xw, String proname) {
        QueryWrapper<JjgWgjc> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd");
        List<JjgWgjc> list = jjgWgjcService.list(wrapper);
        if (list != null){
            List<JjgWgjc> ljgc = new ArrayList<>();
            List<JjgWgjc> sdgc = new ArrayList<>();
            List<JjgWgjc> qlgc = new ArrayList<>();
            List<JjgWgjc> lmgc = new ArrayList<>();
            List<JjgWgjc> ja = new ArrayList<>();

            for (JjgWgjc jjgWgjc : list) {
                String dwgc = jjgWgjc.getDwgc();
                if (dwgc.contains("路基工程")){
                    ljgc.add(jjgWgjc);
                }else if (dwgc.contains("隧道工程")){
                    sdgc.add(jjgWgjc);
                }else if (dwgc.contains("桥梁工程")){
                    qlgc.add(jjgWgjc);
                }else if (dwgc.contains("路面工程")){
                    lmgc.add(jjgWgjc);
                }else if (dwgc.contains("交通安全设施")){
                    ja.add(jjgWgjc);
                }
            }
            //按分部工程分类
            Map<String, List<JjgWgjc>> groupedData = ljgc.stream()
                    .collect(Collectors.groupingBy(JjgWgjc::getFbgc));
            if (groupedData != null){
                groupedData.forEach((fbgc, ljlist) -> {
                    System.out.println("Key: " + fbgc);//路基土石方
                    if (fbgc.contains("路基土石方")){
                        String tsf = extracted(list, fbgc, ljlist);
                        xw.replace("${路基土石方}", tsf, false, true);
                        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

                    }else if (fbgc.contains("排水工程")){
                        String ps = extracted(list, fbgc, ljlist);
                        xw.replace("${排水工程}", ps, false, true);
                        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

                    }else if (fbgc.contains("小桥")){
                        String xq = extracted(list, fbgc, ljlist);
                        xw.replace("${小桥}", xq, false, true);
                        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

                    }else if (fbgc.contains("涵洞")){
                        String hd = extracted(list, fbgc, ljlist);
                        xw.replace("${涵洞}", hd, false, true);
                        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

                    }else if (fbgc.contains("支挡工程")){
                        String zd = extracted(list, fbgc, ljlist);
                        xw.replace("${支挡工程}", zd, false, true);
                        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);

                    }
                });
            }

            Map<String, List<JjgWgjc>> groupedDataql = qlgc.stream()
                    .collect(Collectors.groupingBy(j -> {
                        String gjmc2 = j.getGjmc2();
                        return gjmc2 != null ? gjmc2 : j.getGjmc();
                    }));
            if (groupedDataql!=null){
                groupedDataql.forEach((gjmc, qllist) -> {
                    System.out.println("Key: " + gjmc);//路基土石方
                    int i = 1;
                    String ql = extractedql(gjmc, qllist,i);
                    xw.replace("${桥梁工程}", ql, false, true);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                    i++;

                });
            }
            Map<String, List<JjgWgjc>> groupedDatasd = sdgc.stream()
                    .collect(Collectors.groupingBy(j -> {
                        String gjmc2 = j.getGjmc2();
                        String gjmc = j.getGjmc();
                        return gjmc != null ? gjmc : gjmc2;
                    }));
            if (groupedDatasd!=null){
                groupedDatasd.forEach((gjmc, qllist) -> {
                    System.out.println("Key: " + gjmc);//路基土石方
                    int i = 1;
                    String ql = extractedql(gjmc, qllist,i);
                    xw.replace("${隧道工程}", ql, false, true);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                    i++;

                });
            }

            Map<String, List<JjgWgjc>> groupedDataja = ljgc.stream()
                    .collect(Collectors.groupingBy(JjgWgjc::getFbgc));
            if (groupedDataja!=null){
                groupedDataja.forEach((gjmc, qllist) -> {
                    System.out.println("Key: " + gjmc);//路基土石方
                    int i = 1;
                    String jass = extractedql(gjmc, qllist,i);
                    xw.replace("${交通安全设施}", jass, false, true);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                    i++;

                });
            }

            Map<String, List<JjgWgjc>> groupedDatalm = sdgc.stream()
                    .collect(Collectors.groupingBy(j -> {
                        String gjmc2 = j.getGjmc2();
                        String gjmc = j.getGjmc();
                        return gjmc != null ? gjmc : gjmc2;
                    }));
            if (groupedDatalm!=null){
                groupedDatalm.forEach((gjmc, qllist) -> {
                    System.out.println("Key: " + gjmc);//路基土石方
                    int i = 1;
                    String lm = extractedql(gjmc, qllist,i);
                    xw.replace("${路面工程}", lm, false, true);
                    xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                    i++;

                });
            }
        }
    }

    /**
     *
     * @param gjmc
     * @param qllist
     * @param i
     * @return
     */
    private String extractedql(String gjmc, List<JjgWgjc> qllist, int i) {
        StringBuilder ljtsfBuilder = new StringBuilder();
        if (qllist!=null){
            ljtsfBuilder.append("("+i+")个别").append(gjmc).append("存在");
            String bhlxms = "";
            for (JjgWgjc jjgWgjc : qllist) {
                bhlxms += jjgWgjc.getBhlx()+"、";
            }
            ljtsfBuilder.append(bhlxms).append("等");
            Set<String> htd = new HashSet<>();
            for (JjgWgjc jjgWgjc : qllist) {
                htd.add(jjgWgjc.getHtd());
            }
            ljtsfBuilder.append(",共计").append(htd.size()).append("个合同段");
            ljtsfBuilder.append("。如").append(qllist.get(0).getHtd()).append("标").append(qllist.get(0).getGjbh()).
                    append(qllist.get(0).getBhlx()).append(qllist.get(0).getBhsl()).append("处，").append(qllist.get(0).getBhdl()).append("。");

        }
        String ql = ljtsfBuilder.toString();
        return ql;
    }

    /**
     *
     * @param list
     * @param fbgc
     * @param ljlist
     * @return
     */
    private String extracted(List<JjgWgjc> list, String fbgc, List<JjgWgjc> ljlist) {
        StringBuilder ljtsfBuilder = new StringBuilder();
        //还要根据构件名称分一下
        Map<String, List<JjgWgjc>> groupedljlist = ljlist.stream()
                .collect(Collectors.groupingBy(JjgWgjc::getGjmc));
        if (groupedljlist!=null){
            groupedljlist.forEach((gjmc, list1) -> {
                int i = 1;
                ljtsfBuilder.append(i+")个别").append(gjmc);
                String bhlxms = "";
                for (JjgWgjc jjgWgjc : list1) {
                    bhlxms += jjgWgjc.getBhlx()+"、";
                }
                ljtsfBuilder.append("有").append(bhlxms).append("现象");
                //查询一下一共有几个合同段
                Set<String> htd = new HashSet<>();
                for (JjgWgjc jjgWgjc : list) {
                    String fbgc1 = jjgWgjc.getFbgc();
                    String gjmc1 = jjgWgjc.getGjmc();
                    if (fbgc1.equals(fbgc) && gjmc1.equals(gjmc)){
                        htd.add(jjgWgjc.getHtd());
                    }
                }
                ljtsfBuilder.append(",共计").append(htd.size()).append("个合同段");

                ljtsfBuilder.append("。如").append(list1.get(0).getHtd()).append("标").append(list1.get(0).getGjbh()).
                        append(list1.get(0).getBhlx()).append(list1.get(0).getBhsl()).append("处，").append(list1.get(0).getBhdl()).append("。");
                i++;
            });
        }
        String ljtsf = ljtsfBuilder.toString();
        return ljtsf;
    }


    /**
     *
     * @param f
     * @param xw
     * @param proname
     */
    private void CreateTablewgjcl(File f, Document xw, String proname) {
        QueryWrapper<JjgWgjc> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd");
        List<JjgWgjc> list = jjgWgjcService.list(wrapper);

        //按单位工程分类
        if (list != null && list.size()>0){
            List<JjgWgjc> ljgc = new ArrayList<>();
            List<JjgWgjc> lmgc = new ArrayList<>();
            List<JjgWgjc> sdgc = new ArrayList<>();
            List<JjgWgjc> qlgc = new ArrayList<>();
            List<JjgWgjc> jagc = new ArrayList<>();
            for (JjgWgjc jjgWgjc : list) {
                String dwgc = jjgWgjc.getDwgc();
                if (dwgc.contains("路基工程")){
                    ljgc.add(jjgWgjc);

                }else if (dwgc.contains("隧道工程")){
                    sdgc.add(jjgWgjc);

                }else if (dwgc.contains("路面工程")){
                    lmgc.add(jjgWgjc);

                }else if (dwgc.contains("桥梁工程")){
                    qlgc.add(jjgWgjc);

                }else if (dwgc.contains("交通安全设施工程")){
                    jagc.add(jjgWgjc);

                }
            }

           /* String lm = "${路面外观检查}";
            String ql = "${桥梁工程外观检查}";
            String sd = "${隧道工程外观检查}";
            String ja = "${交通安全设施工程外观检查}";
            //TextSelection textSelectionlj = xw.findString(lj, true, true);
            TextSelection textSelectionlm = xw.findString(lm, true, true);
            TextSelection textSelectionql = xw.findString(ql, true, true);
            TextSelection textSelectionsd = xw.findString(sd, true, true);
            TextSelection textSelectionja = xw.findString(ja, true, true);*/

            extractedlj(f, xw, ljgc);
            extractedlm(f, xw, lmgc);
            extractedqlgc(f, xw, qlgc);
            extractedsd(f, xw, sdgc);
            extractedja(f, xw, jagc);



            /*List<TextSelection> selection = new ArrayList<>();
            selection.add(textSelectionlj);
            selection.add(textSelectionlm);
            selection.add(textSelectionql);
            selection.add(textSelectionsd);
            selection.add(textSelectionja);

            String[] header = {"合同段", "单位工程", "分部工程","外观检查(病害描述)","备注"};
            if (selection != null){
                for (TextSelection textSelection : selection) {
                    if (textSelection != null){
                        Table table = new Table(xw, true);

                        table.resetCells(ljgc.size() + 1, header.length);
                        Paragraph paragraph = textSelection.getAsOneRange().getOwnerParagraph();
                        Body body = paragraph.ownerTextBody();
                        int index = body.getChildObjects().indexOf(paragraph);

                        //将第一行设置为表格标题
                        TableRow row = table.getRows().get(0);
                        row.isHeader(true);
                        row.setHeight(40);
                        row.setHeightType(TableRowHeightType.Exactly);
                        for (int i = 0; i < header.length; i++) {
                            row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                            Paragraph p = row.getCells().get(i).addParagraph();
                            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                            TextRange txtRange = p.appendText(header[i]);
                            txtRange.getCharacterFormat().setBold(true);
                        }
                        //添加数据
                        if (selection.equals("textSelectionlj")){
                            wgff(ljgc, header, table);
                        }else if (selection.equals("textSelectionlm")){
                            wgff(lmgc, header, table);
                        }else if (selection.equals("textSelectionql")){
                            wgff(qlgc, header, table);
                        }else if (selection.equals("textSelectionsd")){
                            wgff(sdgc, header, table);
                        }else if (selection.equals("textSelectionja")){
                            wgff(jagc, header, table);
                        }

                        body.getChildObjects().remove(paragraph);
                        body.getChildObjects().insert(index, table);
                        //列宽自动适应内容
                        table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
                        xw.saveToFile(f.getPath(), FileFormat.Docx_2013);
                    }
                }


            }*/

        }


    }

    /**
     *
     * @param f
     * @param xw
     * @param jagc
     */
    private void extractedja(File f, Document xw, List<JjgWgjc> jagc) {
        if (jagc !=null && jagc.size()>0) {
            String[] headerlj1 = {"合同段", "分部工程", "外观检查结果","","","","","工程量","分部工程扣分"};
            String[] headerlj2 = {"", "", "检查部位","缺陷描述","扣分单位","缺陷数量","扣分","",""};
            String ja = "${交通安全设施工程外观检查}";
            TextSelection textSelectionja = xw.findString(ja, true, true);
            if (textSelectionja != null){
                Table table = new Table(xw, true);
                table.resetCells(jagc.size() + 2, 9);
                Paragraph paragraph = textSelectionja.getAsOneRange().getOwnerParagraph();
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
                for (int rowIdx = 0; rowIdx < jagc.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    JjgWgjc rowData = jagc.get(rowIdx);

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
                                p.appendText(rowData.getBhms()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.getBhsl()).getCharacterFormat().setFontSize(9f);
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

    /**
     *
     * @param f
     * @param xw
     * @param sdgc
     */
    private void extractedsd(File f, Document xw, List<JjgWgjc> sdgc) {
        if (sdgc !=null && sdgc.size()>0) {
            String[] headerlj1 = {"合同段", "序号", "单位工程","分部工程","外观检查结果","","分部工程扣分"};
            String[] headerlj2 = {"", "", "","","缺陷描述","扣分",""};
            String sd = "${隧道工程外观检查}";
            TextSelection textSelectionsd = xw.findString(sd, true, true);
            if (textSelectionsd != null){
                Table table = new Table(xw, true);
                table.resetCells(sdgc.size() + 2, 7);
                Paragraph paragraph = textSelectionsd.getAsOneRange().getOwnerParagraph();
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
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 5);
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
                    if (i == 0 || i == 1 || i == 2 || i == 3 || i == 6) {
                        table.applyVerticalMerge(i, 0, 1);
                    }
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < sdgc.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    JjgWgjc rowData = sdgc.get(rowIdx);

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
                                p.appendText(String.valueOf(rowIdx+1)).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.getDwgc()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.getFbgc()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.getBhms()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
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

    /**
     *
     * @param f
     * @param xw
     * @param qlgc
     */
    private void extractedqlgc(File f, Document xw, List<JjgWgjc> qlgc) {
        if (qlgc !=null && qlgc.size()>0) {
            String[] headerlj1 = {"合同段", "序号", "单位工程","分部工程","外观检查","","分部工程扣分"};
            String[] headerlj2 = {"", "", "","","缺陷描述","扣分",""};
            String ql = "${桥梁工程外观检查}";
            TextSelection textSelectionql = xw.findString(ql, true, true);
            if (textSelectionql != null){
                Table table = new Table(xw, true);
                table.resetCells(qlgc.size() + 2, 7);
                Paragraph paragraph = textSelectionql.getAsOneRange().getOwnerParagraph();
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
                    if (i == 4) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 4, 5);
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
                    if (i == 0 || i == 1 || i == 2 || i == 3 || i == 6) {
                        table.applyVerticalMerge(i, 0, 1);
                    }
                }
                //添加数据
                for (int rowIdx = 0; rowIdx < qlgc.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    JjgWgjc rowData = qlgc.get(rowIdx);

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
                                p.appendText(String.valueOf(rowIdx+1)).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText(rowData.getDwgc()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.getFbgc()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText(rowData.getBhms()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 6:
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

    /**
     *
     * @param f
     * @param xw
     * @param lmgc
     */
    private void extractedlm(File f, Document xw, List<JjgWgjc> lmgc) {
        if (lmgc !=null && lmgc.size()>0) {
            String[] headerlj1 = {"合同段", "单位工程", "部位","外观检查结果","","分部工程扣分"};
            String[] headerlj2 = {"", "", "","存在问题","扣分",""};
            String lm = "${路面外观检查}";
            TextSelection textSelectionlm = xw.findString(lm, true, true);
            if (textSelectionlm != null){
                Table table = new Table(xw, true);
                table.resetCells(lmgc.size() + 2, 6);
                Paragraph paragraph = textSelectionlm.getAsOneRange().getOwnerParagraph();
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
                    if (i == 3) {
                        //合并向右的三个单元格
                        table.applyHorizontalMerge(0, 3, 4);
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
                    if (i == 0 || i == 1 || i == 2 || i == 5) {
                        table.applyVerticalMerge(i, 0, 1);
                    }
                }

                //添加数据
                for (int rowIdx = 0; rowIdx < lmgc.size(); rowIdx++) {
                    TableRow row2 = table.getRows().get(rowIdx + 2); // 第一行已经是表头，所以从第二行开始添加数据
                    //row1.setHeight(80);
                    row2.setHeightType(TableRowHeightType.Exactly);
                    row2.setHeight(34);
                    JjgWgjc rowData = lmgc.get(rowIdx);

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
                                p.appendText(rowData.getDwgc()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 2:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 3:
                                p.appendText(rowData.getBhms()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
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

    /**
     *
     * @param f
     * @param xw
     * @param ljgc
     */
    private static void extractedlj(File f, Document xw, List<JjgWgjc> ljgc) {
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
                    JjgWgjc rowData = ljgc.get(rowIdx);

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
                                p.appendText(rowData.getBhms()).getCharacterFormat().setFontSize(9f);
                                break;
                            case 4:
                                p.appendText("\\").getCharacterFormat().setFontSize(9f);
                                break;
                            case 5:
                                p.appendText(rowData.getBhsl()).getCharacterFormat().setFontSize(9f);
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

    /**
     *
     * @param ljgc
     * @param header
     * @param table
     */
    private void wgff(List<JjgWgjc> ljgc, String[] header, Table table) {
        for (int rowIdx = 0; rowIdx < ljgc.size(); rowIdx++) {
            TableRow row1 = table.getRows().get(rowIdx + 1); // 第一行已经是表头，所以从第二行开始添加数据
            row1.setHeightType(TableRowHeightType.Exactly);

            JjgWgjc rowData = ljgc.get(rowIdx);

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
                        p.appendText(rowData.getDwgc());
                        break;
                    case 2:
                        p.appendText(rowData.getFbgc());
                        break;
                    case 3:
                        p.appendText(rowData.getGjbh()+"等"+rowData.getGjmc()+"有"+rowData.getBhlx()+"现象");
                        break;
                    case 4:
                        p.appendText(rowData.getBz());
                        break;
                    default:
                        break;
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
    private void CreateTablenyzl(File f, Document xw, String proname) {
        QueryWrapper<JjgNyzlkfJg> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.orderByAsc("htd");
        List<JjgNyzlkfJg> list = jjgNyzlkfJgService.list(wrapper);

        String[] header = {"合同段", "存在问题", "扣分"};

        String s = "${附表19内业资料}";
        TextSelection textSelection = xw.findString(s, true, true);
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
                row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
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
                            p.appendText(rowData.getHtd()+"合同段");
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
    private void CreateTableData(File f, Document xw, String proname) {
        //获取数据
        QueryWrapper<JjgHtd> wrapper = new QueryWrapper<>();
        wrapper.eq("proname",proname);
        List<JjgHtd> list = jjgHtdService.list(wrapper);
        List<Map<String,Object>> htdlist = new ArrayList<>();
        if (list != null){
            for (int i = 0; i < list.size(); i++) {
                Map<String,Object> map = new HashMap<>();
                map.put("xh",i+1);
                map.put("htd",list.get(i).getName());
                map.put("zh",list.get(i).getZhq()+"~"+list.get(i).getZhz());
                map.put("tze",list.get(i).getTze());
                map.put("sgdw",list.get(i).getSgdw());
                map.put("jldw",list.get(i).getJldw());
                String lx = list.get(i).getLx();
                String gcnr = "";
                if (lx.contains("桥梁工程")){
                    QueryWrapper<JjgLqsQl> wrapperql = new QueryWrapper<>();
                    wrapperql.eq("proname",proname);
                    wrapperql.eq("htd",list.get(i).getName());
                    List<JjgLqsQl> list1 = jjgLqsQlService.list(wrapperql);
                    if (list1 != null && list1.size()>0){
                        gcnr = "桥梁:";
                        Set<String> set  = new HashSet<>();
                        for (JjgLqsQl jjgLqsQl : list1) {
                            String qlname = jjgLqsQl.getQlname();
                            set.add(qlname);
                        }
                        if (set !=null && set.size()>0){
                            for (String s : set) {
                                gcnr+=s+"、";
                            }
                        }
                        gcnr += "\r\n";
                    }

                }
                if (lx.contains("隧道工程")){
                    QueryWrapper<JjgLqsSd> wrappersd = new QueryWrapper<>();
                    wrappersd.eq("proname",proname);
                    wrappersd.eq("htd",list.get(i).getName());
                    List<JjgLqsSd> list1 = jjgLqsSdService.list(wrappersd);
                    if (list1 != null && list1.size()>0){
                        gcnr += "隧道：";
                        Set<String> set  = new HashSet<>();
                        for (JjgLqsSd jjgLqsSd : list1) {
                            String sdname = jjgLqsSd.getSdname();
                            set.add(sdname);
                        }
                        if (set !=null && set.size()>0){
                            for (String s : set) {
                                gcnr+=s+"、";
                            }
                        }
                        gcnr += "\r\n";
                    }

                }
                if (lx.contains("路面工程")){
                    gcnr += "路面工程面层";
                    gcnr += "\r\n";
                }
                if (lx.contains("交安工程")){
                    gcnr += "标志、标线、护栏";
                    gcnr += "\r\n";
                }
                //匝道
                QueryWrapper<JjgLqsHntlmzd> queryWrapper = new QueryWrapper<>();
                queryWrapper.eq("proname",proname);
                queryWrapper.eq("htd",list.get(i).getName());
                List<JjgLqsHntlmzd> list1 = jjgLqsHntlmzdService.list(queryWrapper);
                if (list1!=null && list1.size()>0){
                    gcnr += "匝道：";
                    Set<String> set  = new HashSet<>();
                    for (JjgLqsHntlmzd jjgLqsHntlmzd : list1) {
                        String hntlmname = jjgLqsHntlmzd.getWz();
                        set.add(hntlmname);
                        //gcnr += hntlmname+"、";
                    }
                    if (set !=null && set.size()>0){
                        for (String s : set) {
                            gcnr+=s+"、";
                        }
                    }
                    gcnr += "\r\n";
                }
                //连接线
                QueryWrapper<JjgLjx> ljxWrapper = new QueryWrapper<>();
                ljxWrapper.eq("proname",proname);
                ljxWrapper.eq("htd",list.get(i).getName());
                List<JjgLjx> list2 = jjgLqsLjxService.list(ljxWrapper);
                if (list2!=null && list2.size()>0){
                    gcnr += "连接线：";
                    Set<String> set  = new HashSet<>();
                    for (JjgLjx ljx : list2) {
                        String name = ljx.getLjxname();
                        set.add(name);
                        //gcnr += name+"、";
                    }
                    if (set !=null && set.size()>0){
                        for (String s : set) {
                            gcnr+=s+"、";
                        }
                    }
                    gcnr += "\r\n";
                }
                //收费站
                QueryWrapper<JjgSfz> sfzWrapper = new QueryWrapper<>();
                sfzWrapper.eq("proname",proname);
                sfzWrapper.eq("htd",list.get(i).getName());
                List<JjgSfz> list3 = jjgLqsSfzService.list(sfzWrapper);
                if (list3!=null && list3.size()>0){
                    gcnr += "收费站：";
                    Set<String> set  = new HashSet<>();
                    for (JjgSfz jjgSfz : list3) {
                        String name = jjgSfz.getZdsfzname();
                        set.add(name);
                        //cnr += name+"、";
                    }
                    if (set !=null && set.size()>0){
                        for (String s : set) {
                            gcnr+=s+"、";
                        }
                    }
                }

                map.put("gcnr",gcnr);
                htdlist.add(map);
            }
        }
        String[] header = {"序号", "施工合同段", "起讫（中心）桩号", "合同造价（万元）", "施工单位","监理单位","主要工程内容"};

        String s = "${表1.3-1一览表}";
        //TextSelection textSelection = xw.findString(s, true, true);
        TextSelection textSelection = findStringInPages(xw,s,3,6);
        // 检查是否找到字符串
        if (textSelection != null)
        {
            Table table=new Table(xw,true);

            table.resetCells(htdlist.size()+1, header.length);
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
            for (int rowIdx = 0; rowIdx < htdlist.size(); rowIdx++) {
                TableRow row1 = table.getRows().get(rowIdx+1); // 第一行已经是表头，所以从第二行开始添加数据
                //row1.setHeight(80);
                row1.setHeightType(TableRowHeightType.Exactly);

                Map<String, Object> rowData = htdlist.get(rowIdx);

                for (int colIdx = 0; colIdx < header.length; colIdx++) {
                    /*if (colIdx == 0){
                        row1.getCells().get(colIdx).setWidth(200);
                    }else if (colIdx == 1 || colIdx == 2 || colIdx == 3){
                        row1.getCells().get(colIdx).setWidth(300);
                    }else if (colIdx == 4 || colIdx == 5 ){
                        row1.getCells().get(colIdx).setWidth(400);
                    }else if (colIdx == 6 ){
                        row1.getCells().get(colIdx).setWidth(600);
                    }*/

                    row1.getCells().get(colIdx).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                    Paragraph p = row1.getCells().get(colIdx).addParagraph();
                    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

                    // 根据表头的不同，设置相应的数据
                    switch (colIdx) {
                        case 0:
                            p.appendText(rowData.get("xh").toString());
                            break;
                        case 1:
                            p.appendText(rowData.get("htd").toString());
                            break;
                        case 2:
                            p.appendText(rowData.get("zh").toString());
                            break;
                        case 3:
                            p.appendText(rowData.get("tze").toString());
                            break;
                        case 4:
                            p.appendText(rowData.get("sgdw").toString());
                            break;
                        case 5:
                            p.appendText(rowData.get("jldw").toString());
                            break;
                        case 6:
                            p.appendText(rowData.get("gcnr").toString());
                            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
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

    /**
     *
     * @param sheet
     * @param allocatedRange
     * @return
     */
    private CellRange RemoveEmptyRows(Worksheet sheet, CellRange allocatedRange) {
        CellRange dataRange = null;
        for (int i = allocatedRange.getLastRow(); i >= 1; i--) {
            boolean isRowEmpty = true;
            for (int col = 1; col <= allocatedRange.getLastColumn(); col++) {
                if (!sheet.getCellRange(i, col).getText().isEmpty()) {
                    isRowEmpty = false;
                    break;
                }
            }
            if (isRowEmpty) {
                sheet.deleteRow(i);
            } else {
                dataRange = sheet.getCellRange(1, 1, i, allocatedRange.getLastColumn());
                break;
            }
        }
        return dataRange;
    }

    /**
     *
     * @param cell
     * @param doc
     * @param path
     * @param range
     */
    private static void copyToWord1(CellRange cell, Document doc, String path,TextRange range) {
        //添加表格
        Table table=new Table(doc,true);
        table.resetCells(cell.getRowCount(), cell.getColumnCount());
        Paragraph paragraph=range.getOwnerParagraph();
        Body body=paragraph.ownerTextBody();
        int index=body.getChildObjects().indexOf(paragraph);

        for (int r = 0; r < cell.getRowCount(); r++) {
            for (int c = 0; c < cell.getLastColumn(); c++) {
                CellRange xCell = cell.get(r+1, c+1);
                CellRange mergeArea = xCell.getMergeArea();
                //合并单元格
                if (mergeArea != null && mergeArea.getRow() == r+1 && mergeArea.getColumn() == c+1) {
                    int rowIndex = mergeArea.getRow()-1;
                    int columnIndex = mergeArea.getColumn()-1;
                    int rowCount = mergeArea.getRowCount();
                    int columnCount = mergeArea.getColumnCount();
                    for (int m = 0; m < rowCount; m++) {
                        table.applyHorizontalMerge(rowIndex + m, columnIndex, columnIndex + columnCount - 1);
                    }
                    table.applyVerticalMerge(columnIndex, rowIndex, rowIndex + rowCount - 1);
                }
                //复制内容
                TableCell wCell = table.getRows().get(r).getCells().get(c);
                if (!xCell.getDisplayedText().isEmpty()) {
                    range=wCell.addParagraph().appendText(xCell.getDisplayedText());
                    copyStyle(range, xCell, wCell);
                } else {
                    wCell.getCellFormat().setBackColor(xCell.getStyle().getColor());
                }
            }
        }
        //body.getChildObjects().remove(paragraph);
        body.getChildObjects().insert(index,table);

        Paragraph newParagraph = new Paragraph(doc); // 创建一个新的段落
        newParagraph.appendBreak(BreakType.Line_Break); // 添加一个换行符
        body.getChildObjects().insert(index + 1, newParagraph);

        doc.saveToFile(path,com.spire.doc.FileFormat.Docx_2013);
    }



    /**
     * @param cell
     * @param doc
     * @param path
     */
    private static void copyToWord(CellRange cell, Document doc, String path,TextRange range) {
        //添加表格
        Table table=new Table(doc,true);
        table.resetCells(cell.getRowCount()-1, cell.getColumnCount());
        Paragraph paragraph=range.getOwnerParagraph();
        Body body=paragraph.ownerTextBody();
        int index=body.getChildObjects().indexOf(paragraph);

        for (int r = 2; r <= cell.getRowCount(); r++) {
            for (int c = 1; c <= cell.getLastColumn(); c++) {
                CellRange xCell = cell.get(r, c);
                CellRange mergeArea = xCell.getMergeArea();
                //合并单元格
                if (mergeArea != null && mergeArea.getRow() == r && mergeArea.getColumn() == c) {
                    int rowIndex = mergeArea.getRow();
                    int columnIndex = mergeArea.getColumn();
                    int rowCount = mergeArea.getRowCount();
                    int columnCount = mergeArea.getColumnCount();
                    for (int m = 0; m < rowCount; m++) {
                        table.applyHorizontalMerge(rowIndex - 2 + m, columnIndex - 1, columnIndex + columnCount - 2);
                    }
                    table.applyVerticalMerge(columnIndex - 1, rowIndex - 2, rowIndex + rowCount - 3);
                }
                //复制内容
                TableCell wCell = table.getRows().get(r - 2).getCells().get(c - 1);
                if (!xCell.getDisplayedText().isEmpty()) {
                    range=wCell.addParagraph().appendText(xCell.getDisplayedText());
                    //TextRange textRange = wCell.addParagraph().appendText(xCell.getDisplayedText());
                    copyStyle(range, xCell, wCell);
                } else {
                    wCell.getCellFormat().setBackColor(xCell.getStyle().getColor());
                }
            }
        }
        body.getChildObjects().remove(paragraph);
        body.getChildObjects().insert(index,table);
        doc.saveToFile(path,com.spire.doc.FileFormat.Docx_2013);
    }



    /**
     *
     * @param wTextRange
     * @param xCell
     * @param wCell
     */
    private static void copyStyle(TextRange wTextRange, CellRange xCell, TableCell wCell) {
        //复制字体样式
        wTextRange.getCharacterFormat().setTextColor(xCell.getStyle().getFont().getColor());
        wTextRange.getCharacterFormat().setFontSize((float) xCell.getStyle().getFont().getSize());
        wTextRange.getCharacterFormat().setFontName(xCell.getStyle().getFont().getFontName());
        wTextRange.getCharacterFormat().setBold(xCell.getStyle().getFont().isBold());
        wTextRange.getCharacterFormat().setItalic(xCell.getStyle().getFont().isItalic());
        //复制背景色
        wCell.getCellFormat().setBackColor(xCell.getStyle().getColor());
        //复制排列方式
        switch (xCell.getHorizontalAlignment()) {
            case Left:
                wTextRange.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                break;
            case Center:
                wTextRange.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                break;
            case Right:
                wTextRange.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
                break;
            default:
                break;
        }
        switch (xCell.getVerticalAlignment()) {
            case Bottom:
                wCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Bottom);
                break;
            case Center:
                wCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                break;
            case Top:
                wCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Top);
                break;
            default:
                break;
        }
    }

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

    public TextSelection[] findStringInPagess(Document xw, String searchString, int startPage, int endPage) {
        for (int currentPage = startPage; currentPage <= endPage; currentPage++) {
            // 在当前页中查找字符串
            TextSelection[] textSelection = xw.findAllString(searchString, true, true);
            // 如果找到字符串，返回结果
            if (textSelection != null) {
                return textSelection;
            }
        }
        // 如果在指定范围内未找到字符串，返回null或适当的标识
        return null;
    }

}
