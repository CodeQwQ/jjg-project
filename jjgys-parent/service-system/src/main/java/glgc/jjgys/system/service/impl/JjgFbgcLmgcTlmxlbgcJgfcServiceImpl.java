package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcTlmxlbgcJgfc;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcTlmxlbgcJgfcVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcTlmxlbgcJgfcMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcTlmxlbgcJgfcService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-09-23
 */
@Service
public class JjgFbgcLmgcTlmxlbgcJgfcServiceImpl extends ServiceImpl<JjgFbgcLmgcTlmxlbgcJgfcMapper, JjgFbgcLmgcTlmxlbgcJgfc> implements JjgFbgcLmgcTlmxlbgcJgfcService {

    @Autowired
    private JjgFbgcLmgcTlmxlbgcJgfcMapper jjgFbgcLmgcTlmxlbgcJgfcMapper;

    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;


    @Override
    public void exportxlbgs(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "砼路面相邻板高差实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcTlmxlbgcJgfcVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"砼路面相邻板高差实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    @Transactional
    public void importxlbgs(MultipartFile file, String proname) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcTlmxlbgcJgfcVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcTlmxlbgcJgfcVo>(JjgFbgcLmgcTlmxlbgcJgfcVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcTlmxlbgcJgfcVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcLmgcTlmxlbgcJgfcVo tlmxlbgcVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(tlmxlbgcVo.getZh())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(tlmxlbgcVo.getQywzmc())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，取样位置名称为空，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(tlmxlbgcVo.getScz1())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，实测值1有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(tlmxlbgcVo.getScz2())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，实测值2有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(tlmxlbgcVo.getScz3())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，实测值3有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(tlmxlbgcVo.getScz4())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，实测值4有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(tlmxlbgcVo.getBgcgdz())) {
                                            throw new JjgysException(20001, "第"+rowNumber+"行的数据中，板高差规定值有误，请修改后重新上传");
                                        }
                                        JjgFbgcLmgcTlmxlbgcJgfc tlmxlbgc = new JjgFbgcLmgcTlmxlbgcJgfc();
                                        BeanUtils.copyProperties(tlmxlbgcVo,tlmxlbgc);
                                        tlmxlbgc.setCreatetime(new Date());
                                        tlmxlbgc.setProname(proname);
                                        tlmxlbgc.setFbgc("路面工程");
                                        jjgFbgcLmgcTlmxlbgcJgfcMapper.insert(tlmxlbgc);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"请将检测日期修改为2021/01/01的格式，然后重新上传");
        }

    }

    @Override
    public boolean generateJdb(String proname) throws IOException, ParseException {
        List<String> htds =  jjgFbgcLmgcTlmxlbgcJgfcMapper.gethtd(proname);
        if (htds.size()>0){
            for (String htd : htds) {
                gethtdjdb(proname,htd);
            }
            return true;
        }else {
            return false;
        }
    }

    @Override
    public List<Map<String, Object>> selecthtd(String proname) {
        List<Map<String,Object>> htdList = jjgFbgcLmgcTlmxlbgcJgfcMapper.selecthtd(proname);
        return htdList;
    }

    /**
     *
     * @param proname
     * @param htd
     * @throws IOException
     * @throws ParseException
     */
    private void gethtdjdb(String proname, String htd) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        //获取数据
        QueryWrapper<JjgFbgcLmgcTlmxlbgcJgfc> wrapper=new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.orderByAsc("qywzmc","zh");
        List<JjgFbgcLmgcTlmxlbgcJgfc> data = jjgFbgcLmgcTlmxlbgcJgfcMapper.selectList(wrapper);

        File f = new File(jgfilepath+File.separator+proname+File.separator+htd+File.separator+"17混凝土路面相邻板高差.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(jgfilepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
                /*File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String path =reportPath +File.separator+ "混凝土路面相邻板高差.xlsx";
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/混凝土路面相邻板高差.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                createTable(gettableNum(data.size()),wb);
                if(DBtoExcel(data,wb)){
                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        if (shouldBeCalculate(wb.getSheetAt(j))) {
                            calculateHeightDifferenceSheet(wb.getSheetAt(j));
                            getTunnelTotal(wb.getSheetAt(j));
                        }
                    }

                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }

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
     * @param sheet
     */
    private void getTunnelTotal(XSSFSheet sheet) {
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow startrow = null;
        XSSFRow endrow = null;
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合计".equals(row.getCell(0).toString())) {
                endrow = sheet.getRow(i-1);
                startrow.createCell(8).setCellFormula("COUNT("
                        +startrow.getCell(3).getReference()+":"
                        +endrow.getCell(3).getReference()+")");//=COUNT(D7:D36)
                startrow.createCell(9).setCellFormula("COUNTIF("
                        +startrow.getCell(4).getReference()+":"
                        +endrow.getCell(4).getReference()+",\"√\")");//=COUNTIF(E7:E36,"√")
                startrow.createCell(10).setCellFormula(startrow.getCell(9).getReference()+"*100/"
                        +startrow.getCell(8).getReference());
                break;
            }
            if(flag && row.getCell(1) != null && !"".equals(row.getCell(1).toString())){
                if(!name.equals(row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"))) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(8).setCellFormula("COUNT("
                            +startrow.getCell(3).getReference()+":"
                            +endrow.getCell(3).getReference()+")");//=COUNT(D7:D36)
                    startrow.createCell(9).setCellFormula("COUNTIF("
                            +startrow.getCell(4).getReference()+":"
                            +endrow.getCell(4).getReference()+",\"√\")");//=COUNTIF(E7:E36,"√")
                    startrow.createCell(10).setCellFormula(startrow.getCell(9).getReference()+"*100/"
                            +startrow.getCell(8).getReference());
                    name = row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"));
                    startrow = row;
                }
                if("".equals(name)){
                    /*
                     * 隧道要分左右幅统计，但渗水系数没有分开统计，所以要根据桩号的z/y来判断
                     */
                    name = row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"));
                    //System.out.println("name1 = "+name);
                    startrow = row;
                }

            }
            if ("序号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
            }

        }

    }

    /**
     *
     * @param sheet
     */
    private void calculateHeightDifferenceSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 可以计算
            if (row.getCell(3).getCellType() == Cell.CELL_TYPE_NUMERIC
                    && flag) {
                row.getCell(4).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(3).getReference()
                                + "<=$C$4,\"√\",\"\"))");// E=IF(D7="","",IF(D7<=$C$4,"√",""))
                row.getCell(6).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + ">$C$4,\"×\",\"\")");// G=IF(D7>$C$4,"×","")
                rowend = row;
            }
            // 可以计算啦
            if ("序号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                rowstart = sheet.getRow(i + 1);
                rowend = rowstart;
            }
            if ("合计".equals(row.getCell(0).toString())) {

                row.getCell(3).setCellFormula(
                        "COUNT(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)
                row.getCell(5).setCellFormula(
                        "COUNTIF(" + rowstart.getCell(4).getReference() + ":"
                                + rowend.getCell(4).getReference() + ",\"√\")");// =COUNTIF(E7:F30,"√")
                row.getCell(7).setCellFormula(
                        row.getCell(5).getReference() + "/"
                                + row.getCell(3).getReference() + "*100");// =F36/D36*100
                row.createCell(8).setCellFormula(
                        "MAX(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)
                row.createCell(9).setCellFormula(
                        "MIN(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)
            }
        }
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        int num = size*4;
        return num%30 <= 29 ? num/30+1 : num/30+2;
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
            if(i < record-1){
                RowCopy.copyRows(wb, "相邻板高差", "相邻板高差", 6, 35, (i - 1) * 30 + 36);
            }
            else{
                RowCopy.copyRows(wb, "相邻板高差", "相邻板高差", 6, 34, (i - 1) * 30 + 36);
            }
        }
        if(record == 1){
            wb.getSheet("相邻板高差").shiftRows(36, 36, -1);
        }
        RowCopy.copyRows(wb, "source", "相邻板高差", 0, 0,(record) * 30 + 5);
        wb.setPrintArea(wb.getSheetIndex("相邻板高差"), 0, 7, 0,(record) * 30 + 5);
    }


    /**
     *
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcLmgcTlmxlbgcJgfc> data, XSSFWorkbook wb) throws ParseException {
        System.out.println(data);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("相邻板高差");
        int index = 6;
        sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
        sheet.getRow(1).getCell(6).setCellValue(data.get(0).getHtd());
        sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
        sheet.getRow(2).getCell(6).setCellValue("路面面层");
        sheet.getRow(3).getCell(2).setCellValue(Double.valueOf(data.get(0).getBgcgdz()));
        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for(int i =1; i < data.size(); i++){
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
        }
        sheet.getRow(3).getCell(6).setCellValue(date);
        for(int i =0; i < data.size(); i++){
            sheet.addMergedRegion(new CellRangeAddress(index, index+3, 0, 0));
            sheet.getRow(index).getCell(0).setCellValue(i+1);

            sheet.addMergedRegion(new CellRangeAddress(index, index+3, 1, 2));
            sheet.getRow(index).getCell(1).setCellValue(data.get(i).getQywzmc()+ data.get(i).getZh());

            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz1()));
            sheet.addMergedRegion(new CellRangeAddress(index, index, 4, 5));
            sheet.addMergedRegion(new CellRangeAddress(index, index, 6, 7));
            sheet.getRow(index+1).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz2()));
            sheet.addMergedRegion(new CellRangeAddress(index+1, index+1, 4, 5));
            sheet.addMergedRegion(new CellRangeAddress(index+1, index+1, 6, 7));
            sheet.getRow(index+2).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz3()));
            sheet.addMergedRegion(new CellRangeAddress(index+2, index+2, 4, 5));
            sheet.addMergedRegion(new CellRangeAddress(index+2, index+2, 6, 7));
            sheet.getRow(index+3).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz4()));
            sheet.addMergedRegion(new CellRangeAddress(index+3, index+3, 4, 5));
            sheet.addMergedRegion(new CellRangeAddress(index+3, index+3, 6, 7));
            index += 4;
        }
        return true;


    }

    /**
     *
     * @param sheet
     * @return
     */
    private boolean shouldBeCalculate(XSSFSheet sheet) {
        String title = null;
        title = sheet.getRow(0).getCell(0).getStringCellValue();
        if (title.endsWith("鉴定表")) {
            return true;
        }
        return false;
    }

    @Override
    public int createMoreRecords(List<JjgFbgcLmgcTlmxlbgcJgfc> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcLmgcTlmxlbgcJgfcMapper, userID);
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(String proname) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> mapList = new ArrayList<>();
        List<String> gethtd = jjgFbgcLmgcTlmxlbgcJgfcMapper.gethtd(proname);
        if (gethtd != null && gethtd.size()>0){
            for (String htd : gethtd) {
                String title = "混凝土路面相邻板高差质量鉴定表";
                String sheetname = "相邻板高差";

                File f = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "17混凝土路面相邻板高差.xlsx");
                if (!f.exists()) {
                    return null;
                } else {
                    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
                    XSSFSheet slSheet = wb.getSheet(sheetname);
                    XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(6);//合同段名
                    XSSFCell hd = slSheet.getRow(2).getCell(2);//分布工程名

                    Map<String, Object> jgmap = new HashMap<>();

                    if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(5).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(3).getCell(2).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum).getCell(9).setCellType(CellType.STRING);
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(5).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        jgmap.put("htd", htd);
                        jgmap.put("总点数", zdsz1);
                        jgmap.put("规定值", slSheet.getRow(3).getCell(2).getStringCellValue());
                        jgmap.put("合格点数", hgdsz1);
                        jgmap.put("合格率", hglz1);
                        jgmap.put("max", slSheet.getRow(lastRowNum).getCell(8).getStringCellValue());
                        jgmap.put("min", slSheet.getRow(lastRowNum).getCell(9).getStringCellValue());
                        mapList.add(jgmap);
                    }
                }
            }
        }
        return mapList;
    }
}
