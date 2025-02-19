package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcQmgzsd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcQmgzsdVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmgzsdMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcQmgzsdService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellType;
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
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Service
public class JjgFbgcQlgcQmgzsdServiceImpl extends ServiceImpl<JjgFbgcQlgcQmgzsdMapper, JjgFbgcQlgcQmgzsd> implements JjgFbgcQlgcQmgzsdService {

    @Autowired
    private JjgFbgcQlgcQmgzsdMapper jjgFbgcQlgcQmgzsdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Autowired
    private SysUserService sysUserService;

    @Override
    public boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String userid = commonInfoVo.getUserid();
        //查询用户类型
        QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
        wrapperuser.eq("id",userid);
        SysUser one = sysUserService.getOne(wrapperuser);
        String type = one.getType();
        //拿到部门下所有用户的id
        List<Long> idlist = new ArrayList<>();
        if ("2".equals(type) || "4".equals(type)){
            //公司管理员
            Long deptId = one.getDeptId();
            QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
            wrapperid.eq("dept_id",deptId);
            List<SysUser> list = sysUserService.list(wrapperid);
            if (list!=null){
                for (SysUser user : list) {
                    Long id = user.getId();
                    idlist.add(id);
                }
            }
        }else if ("3".equals(type)){
            //普通用户
            idlist.add(Long.valueOf(userid));
        }
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmgzsdMapper.selectqlmc(proname,htd,fbgc,idlist);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    DBtoExcelql(proname,htd,fbgc,qlmc,idlist);
                }
            }
            return true;
        }else {
            return false;
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param fbgc
     * @param qlmc
     * @param idlist
     * @throws IOException
     */
    private void DBtoExcelql(String proname, String htd, String fbgc, String qlmc, List<Long> idlist) throws IOException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcQlgcQmgzsd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("qlmc",qlmc);
        wrapper.in("userid",idlist);
        wrapper.orderByAsc("zh");
        List<JjgFbgcQlgcQmgzsd> data = jjgFbgcQlgcQmgzsdMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"37构造深度手工铺沙法-"+qlmc+".xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
                /*File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String path = reportPath + File.separator + "构造深度手工铺沙法.xlsx";
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/构造深度手工铺沙法.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                createTable(gettableNum(data.size()),wb);
                if(DBtoExcel(data,wb,qlmc)){
                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }
                    //wb.setSheetHidden(wb.getSheetIndex("source"), true);
                    JjgFbgcCommonUtils.deleteEmptySheets(wb);

                    FileOutputStream fileOut = new FileOutputStream(f);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();

                }
                inputStream.close();
                out.close();
                wb.close();
            }catch (Exception e) {
                if(f.exists()){
                    f.delete();
                }
                throw new JjgysException(20001, "生成鉴定表错误，请检查数据的正确性");
            }

        }
    }

    /**
     *
     * @param data
     * @param wb
     * @return
     */
    private boolean DBtoExcel(List<JjgFbgcQlgcQmgzsd> data, XSSFWorkbook wb,String qlmc) {
        if (data.size()>0){
            XSSFSheet sheet = wb.getSheet("混凝土桥");
            int tableNO = 0;
            //首先填写表头
            fillCellTitle(data.get(0), tableNO, sheet,qlmc);
            //填写数据\
            int record = 0;
            String name = data.get(0).getJgmc();
            for(JjgFbgcQlgcQmgzsd row : data){
                if(name.equals(row.getJgmc()) && record <= 26){
                    fillCellBody(row, tableNO, record, sheet);
                    record ++;
                }
                else{
                    tableNO ++;
                    fillCellTitle(row, tableNO, sheet,qlmc);
                    name = row.getJgmc();
                    record = 0;
                    fillCellBody(row, tableNO, record, sheet);
                    record ++;
                }
            }
            XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(wb);
            double value = 0;
            XSSFRow row = sheet.getRow(sheet.getLastRowNum());
            int curRow = sheet.getLastRowNum();

            row.getCell(0).setCellValue("检测点数");
            row.getCell(1).setCellFormula("COUNT(L7:L"+curRow+")");
            value = e.evaluate(row.getCell(1)).getNumberValue();
            row.getCell(1).setCellFormula(null);
            row.getCell(1).setCellValue(value);

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 2, 3));
            row.getCell(2).setCellValue("合格点数");
            row.getCell(4).setCellFormula("COUNTIF(M7:M"+curRow+",\"√\")"+"+1-"+gettableNum(data.size()));//gettableNum()
            value = e.evaluate(row.getCell(4)).getNumberValue();
            row.getCell(4).setCellFormula(null);
            row.getCell(4).setCellValue(value);

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 5, 6));
            row.getCell(5).setCellValue("合格率");
            row.getCell(7).setCellFormula(row.getCell(4).getReference()+"*100/"+row.getCell(1).getReference());
            value = e.evaluate(row.getCell(7)).getNumberValue();
            row.getCell(7).setCellFormula(null);
            row.getCell(7).setCellValue(value);

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 8, 9));
            row.getCell(8).setCellValue("最大值");
            row.getCell(10).setCellFormula("MAX(L7:L"+curRow+")");
            value = e.evaluate(row.getCell(10)).getNumberValue();
            row.getCell(10).setCellFormula(null);
            row.getCell(10).setCellValue(value);

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 11, 12));
            row.getCell(11).setCellValue("最小值");
            row.getCell(13).setCellFormula("MIN(L7:L"+curRow+")");
            value = e.evaluate(row.getCell(13)).getNumberValue();
            row.getCell(13).setCellFormula(null);
            row.getCell(13).setCellValue(value);
            return true;

        }else {
            return false;
        }
    }

    /**
     *
     * @param row
     * @param tableNO
     * @param record
     * @param sheet
     */
    private void fillCellBody(JjgFbgcQlgcQmgzsd row, int tableNO, int record, XSSFSheet sheet) {
        XSSFRow currow = sheet.getRow(tableNO * 33 + 6 + record);
        currow.getCell(0).setCellValue(row.getZh());
        currow.getCell(1).setCellValue(row.getCd());
        currow.getCell(2).setCellValue(Double.valueOf(row.getCd1d1()));
        currow.getCell(3).setCellValue(Double.valueOf(row.getCd1d2()));
        currow.getCell(4).setCellFormula("ROUND(IF(ISERROR(31831/AVERAGE("+currow.getCell(2).getReference()+":"+currow.getCell(3).getReference()+")^2),\"\","
                + "31831/AVERAGE("+currow.getCell(2).getReference()+":"+currow.getCell(3).getReference()+")^2),2)");

        currow.getCell(5).setCellValue(Double.valueOf(row.getCd2d1()));
        currow.getCell(6).setCellValue(Double.valueOf(row.getCd2d2()));
        currow.getCell(7).setCellFormula("ROUND(IF(ISERROR(31831/AVERAGE("+currow.getCell(5).getReference()+":"+currow.getCell(6).getReference()+")^2),\"\","
                + "31831/AVERAGE("+currow.getCell(5).getReference()+":"+currow.getCell(6).getReference()+")^2),2)");

        currow.getCell(8).setCellValue(Double.valueOf(row.getCd3d1()));
        currow.getCell(9).setCellValue(Double.valueOf(row.getCd3d2()));
        currow.getCell(10).setCellFormula("ROUND(IF(ISERROR(31831/AVERAGE("+currow.getCell(8).getReference()+":"+currow.getCell(9).getReference()+")^2),\"\","
                + "31831/AVERAGE("+currow.getCell(8).getReference()+":"+currow.getCell(9).getReference()+")^2),2)");

        currow.getCell(11).setCellFormula("ROUND(IF(ISERROR(("
                +currow.getCell(4).getReference()+"+"
                +currow.getCell(7).getReference()+"+"
                +currow.getCell(10).getReference()+")/3),\"\",("
                +currow.getCell(4).getReference()+"+"
                +currow.getCell(7).getReference()+"+"
                +currow.getCell(10).getReference()+")/3),2)");

        currow.getCell(12).setCellFormula("IF(("+currow.getCell(11).getReference()+">="+row.getSjzxz()+")*AND("
                +currow.getCell(11).getReference()+"<="+row.getSjzdz()+"),\"√\",\"\")");

        currow.getCell(13).setCellFormula("IF("+currow.getCell(12).getReference()+"=\"\",\"×\",\"\")");
    }

    /**
     *
     * @param row
     * @param tableNO
     * @param sheet
     */
    private void fillCellTitle(JjgFbgcQlgcQmgzsd row, int tableNO, XSSFSheet sheet,String qlmc) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFRow titlerow = sheet.getRow(tableNO * 33 + 1);
        titlerow.getCell(1).setCellValue(row.getProname());
        titlerow.getCell(9).setCellValue(row.getHtd());
        titlerow = sheet.getRow(tableNO * 33 + 2);
        titlerow.getCell(1).setCellValue("桥面系("+qlmc+")");
        titlerow.getCell(9).setCellValue(row.getJgmc());
        titlerow = sheet.getRow(tableNO * 33 + 3);
        if(row.getSjzxz().equals(row.getSjzdz())){
            titlerow.getCell(1).setCellValue(""+row.getSjzxz());
        }
        else if(!"".equals(row.getSjzxz()) && "".equals(row.getSjzdz())){
            titlerow.getCell(1).setCellValue(""+row.getSjzxz());
        }
        else if("".equals(row.getSjzxz()) && !"".equals(row.getSjzdz())){
            titlerow.getCell(1).setCellValue("不大于"+row.getSjzdz());
        }
        else if(!"".equals(row.getSjzxz()) && !"".equals(row.getSjzdz())){
            titlerow.getCell(1).setCellValue("不小于"+row.getSjzxz()+"且不大于"+row.getSjzdz());
        }
        titlerow.getCell(9).setCellValue(simpleDateFormat.format(row.getJcsj()));
    }

    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "混凝土桥", "混凝土桥", 0, 32-5, i * (33-5));
        }
        wb.setPrintArea(wb.getSheetIndex("混凝土桥"), 0, 13, 0, record * (33-5)-1);
    }


    private int gettableNum(int size) {
        return size%22 <= 21 ? size/22+1 : size/22+2;
    }


    @Override
    public List<Map<String, Object>> lookJdbjgPage(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String userid = commonInfoVo.getUserid();
        String title = "混凝土路面构造深度质量鉴定表";
        String sheetname = "混凝土桥";
        String fbgc = "桥梁工程";
        //查询用户类型
        QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
        wrapperuser.eq("id",userid);
        SysUser one = sysUserService.getOne(wrapperuser);
        String type = one.getType();
        //拿到部门下所有用户的id
        List<Long> idlist = new ArrayList<>();
        if ("2".equals(type) || "4".equals(type)){
            //公司管理员
            Long deptId = one.getDeptId();
            QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
            wrapperid.eq("dept_id",deptId);
            List<SysUser> list = sysUserService.list(wrapperid);
            if (list!=null){
                for (SysUser user : list) {
                    Long id = user.getId();
                    idlist.add(id);
                }
            }
        }else if ("3".equals(type)){
            //普通用户
            idlist.add(Long.valueOf(userid));
        }
        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmgzsdMapper.selectqlmc(proname,htd,fbgc,idlist);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist) {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    Map<String, Object> lookqljdb = lookqljdb(proname, htd, qlmc, sheetname,title);
                    mapList.add(lookqljdb);
                }
            }
            return mapList;
        }else {
            return null;
        }

    }

    @Override
    public List<Map<String, Object>> getqlname(String proname, String htd) {
        List<Map<String, Object>> list = jjgFbgcQlgcQmgzsdMapper.getqlName(proname,htd);
        return list;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String title = "混凝土路面构造深度质量鉴定表";
        String sheetname = "混凝土桥";

        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmgzsdMapper.selectqlmc2(proname,htd);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist) {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    Map<String, Object> lookqljdb = lookqljdb(proname, htd, qlmc, sheetname,title);
                    if (lookqljdb!=null){
                        mapList.add(lookqljdb);
                    }

                }
            }
            return mapList;
        }else {
            return null;
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param qlmc
     * @param sheetname
     * @throws IOException
     */
    private Map<String, Object> lookqljdb(String proname,String htd,String qlmc,String sheetname,String title) throws IOException {
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "37构造深度手工铺沙法-"+qlmc+".xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = wb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名
            //List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();
            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(1).setCellType(CellType.STRING);//总点数
                slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格率
                slSheet.getRow(3).getCell(1).setCellType(CellType.STRING);//合格率
                double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(1).getStringCellValue());
                double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                String zdsz1 = decf.format(zds);
                String hgdsz1 = decf.format(hgds);
                String hglz1 = df.format(hgl);
                jgmap.put("检测项目", qlmc);
                jgmap.put("检测点数", zdsz1);
                jgmap.put("合格点数", hgdsz1);
                jgmap.put("合格率", hglz1);
                jgmap.put("yxpc", slSheet.getRow(3).getCell(1).getStringCellValue());
            }
            return jgmap;
        }
    }

    @Override
    public void exportqmgzsd(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "11桥面构造深度手工铺砂法实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcQmgzsdVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"11桥面构造深度手工铺砂法实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    @Transactional
    public void importqmgzsd(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcQmgzsdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcQmgzsdVo>(JjgFbgcQlgcQmgzsdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcQmgzsdVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcQlgcQmgzsdVo qmgzsdVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(qmgzsdVo.getQlmc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，桥梁名称为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(qmgzsdVo.getJgmc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，结构名称为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(qmgzsdVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(qmgzsdVo.getCd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，车道值为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(qmgzsdVo.getSjzxz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，设计最小值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(qmgzsdVo.getSjzdz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，设计最大值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(qmgzsdVo.getCd1d1()) || StringUtils.isEmpty(qmgzsdVo.getCd1d1())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点1D1值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(qmgzsdVo.getCd1d2()) || StringUtils.isEmpty(qmgzsdVo.getCd1d2())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点1D2值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(qmgzsdVo.getCd2d1()) || StringUtils.isEmpty(qmgzsdVo.getCd2d1())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点2D1值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(qmgzsdVo.getCd2d2()) || StringUtils.isEmpty(qmgzsdVo.getCd2d2())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点2D2值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(qmgzsdVo.getCd3d1()) || StringUtils.isEmpty(qmgzsdVo.getCd3d1())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点3D1值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(qmgzsdVo.getCd3d2()) || StringUtils.isEmpty(qmgzsdVo.getCd3d2())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中11桥面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点3D2值有误，请修改后重新上传");
                                        }
                                        JjgFbgcQlgcQmgzsd fbgcQlgcQmgzsd = new JjgFbgcQlgcQmgzsd();
                                        BeanUtils.copyProperties(qmgzsdVo,fbgcQlgcQmgzsd);
                                        fbgcQlgcQmgzsd.setCreatetime(new Date());
                                        fbgcQlgcQmgzsd.setUserid(commonInfoVo.getUserid());
                                        fbgcQlgcQmgzsd.setProname(commonInfoVo.getProname());
                                        fbgcQlgcQmgzsd.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcQmgzsd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcQmgzsdMapper.insert(fbgcQlgcQmgzsd);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中,11桥面构造深度手工铺砂法实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中,11桥面构造深度手工铺砂法实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }/*catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中,11桥面构造深度手工铺砂法实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }*/

    }

    @Override
    public List<Map<String, Object>> selectqlmc(String proname, String htd, String fbgc,String userid) {
        //查询用户类型
        QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
        wrapperuser.eq("id",userid);
        SysUser one = sysUserService.getOne(wrapperuser);
        String type = one.getType();
        //拿到部门下所有用户的id
        List<Long> idlist = new ArrayList<>();
        if ("2".equals(type) || "4".equals(type)){
            //公司管理员
            Long deptId = one.getDeptId();
            QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
            wrapperid.eq("dept_id",deptId);
            List<SysUser> list = sysUserService.list(wrapperid);
            if (list!=null){
                for (SysUser user : list) {
                    Long id = user.getId();
                    idlist.add(id);
                }
            }
        }else if ("3".equals(type)){
            //普通用户
            idlist.add(Long.valueOf(userid));
        }
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmgzsdMapper.selectqlmc(proname,htd,fbgc,idlist);
        return qlmclist;
    }

    /**
     * 报告中用的数据
     * @param commonInfoVo
     * @param value
     * @return
     * @throws IOException
     */
    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo, String value) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + value);
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = wb.getSheet("混凝土桥");
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名
            //List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();
            List<Map<String,Object>> list = new ArrayList<>();
            if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(1).setCellType(CellType.STRING);//总点数
                slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格率
                slSheet.getRow(3).getCell(1).setCellType(CellType.STRING);//合格率
                double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(1).getStringCellValue());
                double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                String zdsz1 = decf.format(zds);
                String hgdsz1 = decf.format(hgds);
                String hglz1 = df.format(hgl);
                jgmap.put("检测项目", StringUtils.substringBetween(value, "-", "."));
                jgmap.put("检测点数", zdsz1);
                jgmap.put("合格点数", hgdsz1);
                jgmap.put("合格率", hglz1);
                jgmap.put("yxpc", slSheet.getRow(3).getCell(1).getStringCellValue());
            }
            list.add(jgmap);
            return list;
        }

    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcQlgcQmgzsdMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcQlgcQmgzsdMapper.selectnumname(proname);
        return selectnum;
    }

    @Override
    public int createMoreRecords(List<JjgFbgcQlgcQmgzsd> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcQlgcQmgzsdMapper, userID);
    }
}
