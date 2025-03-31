package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcLcf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLmwcLcfVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLmwcLcfMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcLmwcLcfService;
import glgc.jjgys.system.service.SysUserService;
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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-04-18
 */
@Service
public class JjgFbgcLmgcLmwcLcfServiceImpl extends ServiceImpl<JjgFbgcLmgcLmwcLcfMapper, JjgFbgcLmgcLmwcLcf> implements JjgFbgcLmgcLmwcLcfService {

    @Autowired
    private JjgFbgcLmgcLmwcLcfMapper jjgFbgcLmgcLmwcLcfMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Autowired
    private SysUserService sysUserService;


    @Override
    public boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLmgcLmwcLcf> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        String userid = commonInfoVo.getUserid();

        //查询用户类型
        QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
        wrapperuser.eq("id",userid);
        SysUser one = sysUserService.getOne(wrapperuser);
        String type = one.getType();
        if ("2".equals(type) || "4".equals(type)){
            //公司管理员
            Long deptId = one.getDeptId();
            QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
            wrapperid.eq("dept_id",deptId);
            List<SysUser> list = sysUserService.list(wrapperid);
            //拿到部门下所有用户的id
            List<Long> idlist = new ArrayList<>();

            if (list!=null){
                for (SysUser user : list) {
                    Long id = user.getId();
                    idlist.add(id);
                }
            }
            wrapper.in("userid",idlist);
        }else if ("3".equals(type)){
            //普通用户
            //String username = project.getUsername();
            wrapper.like("userid",userid);
        }
        wrapper.orderByAsc("xh","zh");
        List<JjgFbgcLmgcLmwcLcf> data = jjgFbgcLmgcLmwcLcfMapper.selectList(wrapper);

        //获取“温度修正”所需要的数据
        //List<Map<String,Object>> wdata = jjgFbgcLmgcLmwcLcfMapper.selectwdata(proname,htd,fbgc);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"13路面弯沉(落锤法).xlsx");
        if (data == null || data.size()==0){
            return false;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
                /*File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String path = reportPath + File.separator + "路面弯沉17规范落锤法.xlsx";
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/路面弯沉17规范落锤法.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                //ArrayList<String[]> ref = createTemperatureSheet(wb.getSheet("温度修正"),wdata);
                //int num = gettableNum(proname, htd, fbgc);
                //createTable(num,wb);
                if(DBtoExcel(data,wb)){
                    //右侧的数据
                    //calculateTempDate(wb);
//                    ArrayList<String> totalref = getTotalMark(wb.getSheet("路面弯沉"));
//                    String time = getLastTime(wb.getSheet("路面弯沉"));
                    List<Map<String,Object>> totalrefzf = getTotalMark(wb.getSheet("路面弯沉左幅"),"路面弯沉左幅");
                    List<Map<String,Object>> totalrefyf = getTotalMark(wb.getSheet("路面弯沉右幅"),"路面弯沉右幅");
                    totalrefzf.addAll(totalrefyf);
                    //创建“评定单元”sheet
                    createEvaluateTable(totalrefzf,wb);
                    String time = getLastTime(wb.getSheet("路面弯沉左幅"));
                    //引用数据，然后计算
                    //completeTotleTable(totalref, time,wb,proname,htd,fbgc);
                    completeTotleTable(wb.getSheet("评定单元"), totalrefzf, time,wb,proname,htd);
                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }
                    ArrayList<String> delsheets = new ArrayList<String>();;
                    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                        String sheetName = wb.getSheetAt(i).getSheetName();
                        if(sheetName.contains("路面弯沉") && "".equals(wb.getSheet(sheetName).getRow(11).getCell(0).getStringCellValue())){
                            delsheets.add(wb.getSheetAt(i).getSheetName());
                        }
                    }
                    System.out.println(delsheets);
                    for (int i = 0; i < delsheets.size(); i++) {
                        wb.setSheetHidden(wb.getSheetIndex(delsheets.get(i)), true);
                    }
                    FileOutputStream fileOut = new FileOutputStream(f);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();

                }
                inputStream.close();
                out.close();
                wb.close();
                return true;
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
     * @param ref
     * @param time
     * @param xwb
     * @param proname
     * @param htd
     */
    private void completeTotleTable(XSSFSheet sheet, List<Map<String,Object>> ref, String time, XSSFWorkbook xwb, String proname, String htd) {
        //XSSFSheet sheet = xwb.getSheet("评定单元");
        String sheetname = "";
        XSSFCellStyle cellstyle = xwb.createCellStyle();
        XSSFFont font = xwb.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle.setWrapText(true);//自动换行

        XSSFDataFormat df = xwb.createDataFormat();
        cellstyle.setDataFormat(df.getFormat("#,##0.0"));

        XSSFFont font2 = xwb.createFont();
        font2.setFontHeightInPoints((short) 11);
        font2.setColor(Font.COLOR_RED);
        font2.setFontName("宋体");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");

        //XSSFRow row;
        int index = 5, rowstart = 5;
        int i = 1;
        /*
         * 此处先填写“评定单元”sheet的表头数据
         */
        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(7).setCellValue(htd);
        sheet.getRow(2).getCell(2).setCellValue("路面工程");
        sheet.getRow(2).getCell(7).setCellValue(simpleDateFormat.format(new Date()));
        String keyToCheck1 = "路面弯沉左幅";
        String keyToCheck2 = "路面弯沉右幅";
        for (Map<String, Object> map : ref) {
            if (map.containsKey(keyToCheck1)) {
                String[] split = map.get(keyToCheck1).toString().split(",");
                sheet.getRow(index).getCell(0).setCellValue(String.valueOf(i));
                sheet.getRow(index).getCell(1).setCellFormula(keyToCheck1 + "!" + split[0]);
                sheet.addMergedRegion(new CellRangeAddress(index, index, 1, 4));
                sheet.getRow(index).getCell(5).setCellFormula(keyToCheck1 + "!" + split[1]);
                sheet.getRow(index).getCell(6).setCellFormula(keyToCheck1 + "!" + split[2]);
                sheet.getRow(index).getCell(7).setCellFormula(
                        "IF(" + sheet.getRow(index).getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + sheet.getRow(index).getCell(6).getReference() + "<"
                                + sheet.getRow(index).getCell(5).getReference()
                                + ",\"√\",\"\"))");
                //sheet.getRow(index).getCell(7).setCellStyle(cellstyle);
                sheet.getRow(index).getCell(8).setCellFormula(
                        "IF(" + sheet.getRow(index).getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + sheet.getRow(index).getCell(6).getReference() + ">="
                                + sheet.getRow(index).getCell(5).getReference()
                                + ",\"×\",\"\"))");
                /*if("×".equals(sheet.getRow(index).getCell(8).getRawValue())){
                    sheet.getRow(index).getCell(8).setCellStyle(cellstyle);
                }*/
                //row.getCell(8).setCellStyle(cellstyle);
                index++;
                i++;
            } else if (map.containsKey(keyToCheck2)) {
                String[] split = map.get(keyToCheck2).toString().split(",");
                sheet.getRow(index).getCell(0).setCellValue(String.valueOf(i));
                sheet.getRow(index).getCell(1).setCellFormula(keyToCheck2 + "!" + split[0]);
                sheet.addMergedRegion(new CellRangeAddress(index, index, 1, 4));
                sheet.getRow(index).getCell(5).setCellFormula(keyToCheck2 + "!" + split[1]);
                sheet.getRow(index).getCell(6).setCellFormula(keyToCheck2 + "!" + split[2]);
                sheet.getRow(index).getCell(7).setCellFormula(
                        "IF(" + sheet.getRow(index).getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + sheet.getRow(index).getCell(6).getReference() + "<"
                                + sheet.getRow(index).getCell(5).getReference()
                                + ",\"√\",\"\"))");
                sheet.getRow(index).getCell(8).setCellFormula(
                        "IF(" + sheet.getRow(index).getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + sheet.getRow(index).getCell(6).getReference() + ">="
                                + sheet.getRow(index).getCell(5).getReference()
                                + ",\"×\",\"\"))");
                index++;
                i++;
            }
        }
        sheet.getRow(sheet.getLastRowNum()).getCell(4).setCellFormula(
                "COUNT("
                        + sheet.getRow(rowstart).getCell(5)
                        .getReference() + ":"
                        + sheet.getRow(sheet.getLastRowNum() - 1).getCell(5).getReference()
                        + ")");// =COUNT(F6:F27)

        sheet.getRow(sheet.getLastRowNum()).getCell(6).setCellFormula(
                "COUNTIF("
                        + sheet.getRow(rowstart).getCell(7)
                        .getReference() + ":"
                        + sheet.getRow(sheet.getLastRowNum() - 1).getCell(7).getReference()
                        + ",\"√\")");// =COUNTIF(H6:H27,"√")

        sheet.getRow(sheet.getLastRowNum()).getCell(8).setCellFormula(
                sheet.getRow(sheet.getLastRowNum()).getCell(6).getReference() + "/"
                        + sheet.getRow(sheet.getLastRowNum()).getCell(4).getReference() + "*100");

        /*double value = 0.0;
        for (int i = rowstart; i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (index < ref.size()) {
                row.createCell(0).setCellValue(index/3+1);
                XSSFCellStyle cellstyle2 = (XSSFCellStyle) cellstyle.clone();
                cellstyle2.setDataFormat(df.getFormat("#,##0"));
                row.getCell(0).setCellStyle(cellstyle2);

                row.createCell(1).setCellFormula("路面弯沉"+"!" + ref.get(index));
                sheet.addMergedRegion(new CellRangeAddress(i, i, 1, 4));
                row.getCell(1).setCellStyle(cellstyle);

                row.createCell(5).setCellFormula("路面弯沉"+"!" + ref.get(index + 1));
                row.getCell(5).setCellStyle(cellstyle);

                row.createCell(6).setCellFormula("路面弯沉"+"!" + ref.get(index + 2));
                row.getCell(6).setCellStyle(cellstyle);
                // =IF(G6="","",IF(G6<F6,"√",""))
                row.createCell(7).setCellFormula(
                        "IF(" + row.getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(6).getReference() + "<"
                                + row.getCell(5).getReference()
                                + ",\"√\",\"\"))");
                row.getCell(7).setCellStyle(cellstyle);
                // =IF(G6="","",IF(G6>=F6,"×",""))
                row.createCell(8).setCellFormula(
                        "IF(" + row.getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(6).getReference() + ">="
                                + row.getCell(5).getReference()
                                + ",\"×\",\"\"))");
                if("×".equals(row.getCell(8).getRawValue())){
                    cellstyle.setFont(font2);
                    row.getCell(8).setCellStyle(cellstyle);
                }
                cellstyle.setFont(font);
                row.getCell(8).setCellStyle(cellstyle);
                index += 3;
            }
            if ("合计".equals(row.getCell(0).toString())) {
                row.getCell(4).setCellFormula(
                        "COUNT("
                                + sheet.getRow(rowstart).getCell(5)
                                .getReference() + ":"
                                + sheet.getRow(i - 1).getCell(5).getReference()
                                + ")");// =COUNT(F6:F27)

                row.getCell(6).setCellFormula(
                        "COUNTIF("
                                + sheet.getRow(rowstart).getCell(7)
                                .getReference() + ":"
                                + sheet.getRow(i - 1).getCell(7).getReference()
                                + ",\"√\")");// =COUNTIF(H6:H27,"√")
                row.getCell(8).setCellFormula(
                        row.getCell(6).getReference() + "/"
                                + row.getCell(4).getReference() + "*100");// =G28/E28*100
            }
        }*/
    }

    /**
     *
     * @param sheet
     * @return
     */
    private String getLastTime(XSSFSheet sheet) {
        String time = "1900.1.1";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFRow row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (row.getCell(0) != null && "分部工程名称".equals(row.getCell(0).toString()) && row.getCell(11) != null && !"".equals(row.getCell(11).toString())) {
                try {
                    Date dt1 = simpleDateFormat.parse(time);
                    Date dt2 = simpleDateFormat.parse(row.getCell(11).toString());
                    if(dt1.getTime() < dt2.getTime()){
                        time = row.getCell(11).toString();
                    }
                } catch (ParseException e) {
                    e.printStackTrace();
                }
            }
        }
        return time;
    }

    /**
     *
     * @param ref
     * @param xwb
     */
    private void createEvaluateTable(List<Map<String,Object>>  ref, XSSFWorkbook xwb) {
        int record = 0;
        record = ref.size()/3;
        for(int i = 1; i < record/23+1; i++){
            if(i < record/23){
                RowCopy.copyRows(xwb, "评定单元", "评定单元", 5, 27, (i-1)*23+28);
            }
            else{
                RowCopy.copyRows(xwb, "评定单元", "评定单元", 5, 26, (i-1)*23+28);
            }
        }
        if(record/23 == 0){
            xwb.getSheet("评定单元").shiftRows(28, 28, -1, true , false);
        }
        RowCopy.copyRows(xwb, "source", "评定单元", 0, 1, (record/23+1)*23+4);
        xwb.setPrintArea(xwb.getSheetIndex("评定单元"), 0, 8, 0, (record/23+1)*23+4);

    }

    /**
     *
     * @param sheet
     * @return
     */
    private List<Map<String,Object>> getTotalMark(XSSFSheet sheet,String sheetname) {
        List<Map<String,Object>> result = new ArrayList<>();
        XSSFRow row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("桩号/部位".equals(row.getCell(0).toString()) && row.getCell(13) != null && !"".equals(row.getCell(13).toString())) {
                Map<String,Object> map = new HashMap<>();
                map.put(sheetname,row.getCell(1).getReference()+","+row.getCell(9).getReference()+","+sheet.getRow(i+2).getCell(9).getReference());
                result.add(map);
//                ref.add(row.getCell(1).getReference());
//                // 验收弯沉值（0.01mm）
//                ref.add(row.getCell(9).getReference());
            }
            // 弯沉代表值(0.01mm)
//            if ("代表弯沉值(0.01mm)".equals(row.getCell(6).toString()) && row.getCell(13) != null && !"".equals(row.getCell(13).toString())) {
//                ref.add(row.getCell(9).getReference());
//            }

        }
        return result;
    }

    /**
     *
     * @param xwb
     */
    private void calculateTempDate(XSSFWorkbook xwb,String sheetname) {
        XSSFSheet sheet = xwb.getSheet(sheetname);
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(xwb);
        XSSFRow row;
        boolean flag = false;
        int start_row = 0;
        String left_top = "", right_bottom = "";
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if(row.getCell(2).getStringCellValue().equals("温度影响系数")) {
                row.getCell(3).setCellFormula(
                        "EXP((0.000009*(LN("+sheet.getRow(i+2).getCell(3).getReference() +")-1)*"+
                                sheet.getRow(i+2).getCell(1).getReference() +"+0.004)*(20-"+
                                sheet.getRow(i+4).getCell(4).getReference() +"))");//温度影响系数
            }

            if ("桩号/部位".equals(row.getCell(0).toString()) && !name.equals(row.getCell(1).toString())) {
                if("".equals(name)){
                    name = row.getCell(1).toString();
                    left_top = sheet.getRow(i + 5).getCell(15).getReference();
                    continue;
                }
                else{
                    try{
                        right_bottom = sheet.getRow(i - 5).getCell(16).getReference();
                    }
                    catch(Exception e){
                        right_bottom = sheet.getRow(i - 5).createCell(16).getReference();
                    }
                    fillTempDate(xwb,evaluator, sheet, start_row, i - 3, left_top, right_bottom);
                    start_row = i - 3;
                    name = row.getCell(1).toString();
                    left_top = sheet.getRow(i + 5).getCell(15).getReference();
                }
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                continue;
            }
            if (row.getCell(0) != null && row.getCell(0).toString().startsWith("备注")) {
                flag = false;
                continue;
            }
            // 计算左边的值
            if (flag && row.getCell(3).toString() != null && !"".equals(row.getCell(3).toString()))
            {
                row.createCell(15).setCellFormula(row.getCell(3).getReference());// =E22
            }
            // 计算右边的值
            if (flag && row.getCell(9).toString() != null && !"".equals(row.getCell(9).toString()))
            {
                row.createCell(16).setCellFormula(row.getCell(9).getReference());
            }
        }
        try{
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(16).getReference();
        }
        catch(Exception e){
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).createCell(16).getReference();
        }
        fillTempDate(xwb,evaluator, sheet, start_row, sheet.getPhysicalNumberOfRows(), left_top, right_bottom);
    }

    /**
     *
     * @param evaluator
     * @param sheet
     * @param start_row
     * @param end_row
     * @param left_top
     * @param right_bottom
     */
    private void fillTempDate(XSSFWorkbook wb,XSSFFormulaEvaluator evaluator, XSSFSheet sheet, int start_row, int end_row, String left_top, String right_bottom) {
        int tables = (end_row - start_row)/34;
        XSSFRow row2, row3, row4, row5, row6, row7, row8, row9, row10;
        row2 = sheet.getRow(start_row + 1);
        row3 = sheet.getRow(start_row + 2);
        row4 = sheet.getRow(start_row + 3);
        row5 = sheet.getRow(start_row + 4);
        row6 = sheet.getRow(start_row + 5);
        row7 = sheet.getRow(start_row + 6);
        row8 = sheet.getRow(start_row + 7);
        row9 = sheet.getRow(start_row + 8);
        row10 = sheet.getRow(start_row + 9);
        sheet.getRow(start_row + 10);
        // 舍前均方差
        row2.getCell(13).setCellFormula(
                "IF(ISERROR(STDEV(" + left_top + ":" + right_bottom
                        + ")),\"\",STDEV(" + left_top + ":" + right_bottom
                        + "))");// =IF(ISERROR(STDEV(S10:V34)),"",STDEV(S10:V34))
        // 舍前平均值
        row3.getCell(13).setCellFormula(
                "IF(ISERROR(AVERAGE(" + left_top + ":" + right_bottom
                        + ")),\"\",AVERAGE(" + left_top + ":" + right_bottom
                        + "))");// =IF(ISERROR(AVERAGE(S10:V34)),"",AVERAGE(S10:V34))
        // 代表值=IF(ISERROR((Q3+F6*Q2)*F5*H5*H6),"",(Q3+F6*Q2)*F5*H5*H6)
        // 代表值=IF(ISERROR((N3+D6*N2)*D5*F5*F6),"",(N3+D6*N2)*D5*F5*F6)
        row4.getCell(13).setCellFormula(
                "IF(ISERROR((" + row3.getCell(13).getReference() + "+"
                        + row6.getCell(3).getReference() + "*"
                        + row2.getCell(13).getReference() + ")*"
                        + row5.getCell(3).getReference()+ "*"
                        + row5.getCell(5).getReference()+ "*"
                        + row6.getCell(5).getReference()+ "),\"\",("
                        + row3.getCell(13).getReference() + "+"
                        + row6.getCell(3).getReference() + "*"
                        + row2.getCell(13).getReference() + ")*"
                        + row5.getCell(3).getReference()+ "*"
                        + row5.getCell(5).getReference()+ "*"
                        + row6.getCell(5).getReference()+ ")");// =IF(ISERROR(Q3+C5*Q2*F5),"",(Q3+C5*Q2)*F5*H5)
        // 特异值下限
        row5.getCell(13).setCellFormula(
                "IF(ISERROR(" + row3.getCell(13).getReference() + "-2*"
                        + row2.getCell(13).getReference() + "),\"\","
                        + row3.getCell(13).getReference() + "-2*"
                        + row2.getCell(13).getReference() + ")");// =IF(ISERROR(Q3-2*Q2),"",Q3-2*Q2)
        // 特异值上限
        row6.getCell(13).setCellFormula(
                "IF(ISERROR(" + row3.getCell(13).getReference() + "+2*"
                        + row2.getCell(13).getReference() + "),\"\","
                        + row3.getCell(13).getReference() + "+2*"
                        + row2.getCell(13).getReference() + ")");// =IF(ISERROR(Q3+2*Q2),"",Q3+2*Q2)
        // 总测点数
        row7.getCell(13).setCellFormula(
                "COUNT(" + left_top + ":" + right_bottom + ")");// =COUNT(S10:V34)
        /*
         * 平均弯沉值(0.01mm)L5=IF(Q3>L4,Q8,Q3) 弯沉代表值(0.01mm)L6=IF(Q4<L4,Q9,Q4)
         * 结论L7=IF(L6>L4,"×","√")
         */
        fillSheetBody(wb,5, row5.getCell(8), row3.getCell(13), row4.getCell(8),
                row8.getCell(16));
        fillSheetBody(wb,6, row6.getCell(8), row4.getCell(13), row4.getCell(8),
                row9.getCell(16));
        fillSheetBody(wb,7, row7.getCell(8), row6.getCell(8), row4.getCell(8));
        XSSFRow t_row1, t_row2, t_row3;
        for (int i = 1; i < tables; i++) {
            t_row1 = sheet.getRow(start_row + i*34 + 4);
            t_row2 = sheet.getRow(start_row + i*34 + 5);
            t_row3 = sheet.getRow(start_row + i*34 + 6);
            fillSheetBody(wb,5, t_row1.getCell(8), row3.getCell(13), row4.getCell(8),
                    row8.getCell(16));
            fillSheetBody(wb,6, t_row2.getCell(8), row4.getCell(13), row4.getCell(8),
                    row9.getCell(16));
            fillSheetBody(wb,7, t_row3.getCell(8), row6.getCell(11), row4.getCell(8));
        }
    }

    /**
     *
     * @param xwb
     * @param rowNum
     * @param cell
     */
    private void fillSheetBody(XSSFWorkbook xwb,int rowNum, XSSFCell... cell) {
        /*
         * 前两项为四个参数，后一项为三个参数 平均弯沉值(0.01mm)L5=IF(Q3<L4,Q8,Q3)
         * 弯沉代表值(0.01mm)L6=IF(Q4>L4,Q9,Q4) 结论L7=IF(L6>L4,"×","√")
         */
        XSSFCellStyle cellstyle = xwb.createCellStyle();
        XSSFFont font=xwb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setColor(Font.COLOR_RED);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle.setWrapText(true);//自动换行
        XSSFCell cell1, cell2, cell3, cell4;
        if (cell.length == 4) {
            if(rowNum == 5){
                cell1 = cell[0];
                cell2 = cell[1];
                cell3 = cell[2];
                cell4 = cell[3];
                cell1.setCellFormula(cell2.getReference());
            }else{
                cell1 = cell[0];
                cell2 = cell[1];
                cell3 = cell[2];
                cell4 = cell[3];
                cell1.setCellFormula(cell2.getReference());
            }
        }
        if (cell.length == 3) {
            cell1 = cell[0];
            cell2 = cell[1];
            cell3 = cell[2];
            cell1.setCellFormula("IF(" + cell2.getReference() + ">"
                    + cell3.getReference() + ",\"×\",\"√\")");
            XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(xwb);
            if("×".equals(cell1.getRawValue())){
                cell1.setCellStyle(cellstyle);
            }
        }
    }

    /**
     *
     * @param data
     * @param wb
     * @return
     */
    private boolean DBtoExcel(List<JjgFbgcLmgcLmwcLcf> data, XSSFWorkbook wb) {
        Collections.sort(data, new Comparator<JjgFbgcLmgcLmwcLcf>() {
            @Override
            public int compare(JjgFbgcLmgcLmwcLcf o1, JjgFbgcLmgcLmwcLcf o2) {
                // 名字相同时按照 qdzh 排序
                String zh = o1.getZh();
                String zh1 = o2.getZh();
                int cmp = zh.compareTo(zh1);
                if (cmp != 0) {
                    return cmp;
                }
                // 名字相同时按照 qdzh 排序
                String cd = o1.getCd();
                String cd1 = o2.getCd();
                return cd.compareTo(cd1);
            }
        });
        List<JjgFbgcLmgcLmwcLcf> zfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmwcLcf> yfdata = new ArrayList<>();
        for (JjgFbgcLmgcLmwcLcf datum : data) {
            String zh = datum.getZh().substring(0,1);
            if ("Y".equals(zh)){
                yfdata.add(datum);
            }else {
                zfdata.add(datum);
            }

        }
        if (zfdata.size()>0){
            writezyfdata(zfdata, wb,"路面弯沉左幅");
            calculateTempDate(wb,"路面弯沉左幅");
        }
        if (yfdata.size()>0){
            writezyfdata(yfdata, wb,"路面弯沉右幅");
            calculateTempDate(wb,"路面弯沉右幅");
        }
        return true;

    }

    /**
     *
     * @param data
     * @param wb
     * @param sheetname
     */
    private void writezyfdata(List<JjgFbgcLmgcLmwcLcf> data, XSSFWorkbook wb,String sheetname) {
        Collections.sort(data, new Comparator<JjgFbgcLmgcLmwcLcf>() {
            @Override
            public int compare(JjgFbgcLmgcLmwcLcf o1, JjgFbgcLmgcLmwcLcf o2) {
                String zh = o1.getXh();
                String zh1 = o2.getXh();
                int cmp = zh.compareTo(zh1);
                if (cmp != 0) {
                    return cmp;
                }
                // 名字相同时按照 qdzh 排序
                String cd = o1.getZh();
                String cd1 = o2.getZh();
                return cd.compareTo(cd1);
            }
        });
        DecimalFormat nf = new DecimalFormat("0.00");
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFFont font= wb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平

        XSSFSheet sheet = wb.getSheet(sheetname);
        int num = gettableNum(data);
        createTable(num,wb,sheetname);

        String gcbw = data.get(0).getGcbw();
        String xuhao = data.get(0).getXh();
        HashMap<Integer, String> map = new HashMap<Integer, String>();
        String zhuanghao = gcbw+";\n";
        for(JjgFbgcLmgcLmwcLcf row: data){
            if(xuhao.equals(row.getXh())){
                if(!zhuanghao.contains(row.getGcbw())){
                    zhuanghao += " "+row.getGcbw()+";\n";
                }
            }
            else{
                if(map.get(Integer.valueOf(xuhao)) == null){
                    map.put(Integer.valueOf(xuhao), zhuanghao.substring(0,zhuanghao.length()-1));
                }
                else{
                    map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zhuanghao.substring(0,zhuanghao.length()-1));
                }
                xuhao = row.getXh();
                zhuanghao = row.getGcbw()+";\n";
            }
        }
        if(!"".equals(zhuanghao)){
            if(map.get(Integer.valueOf(xuhao)) == null){
                map.put(Integer.valueOf(xuhao), zhuanghao.substring(0,zhuanghao.length()-2));
            }
            else{
                map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+zhuanghao.substring(0,zhuanghao.length()-2));
            }
        }
        int index = 0;
        int tableNum = 0;
        xuhao = data.get(0).getXh();
        ArrayList<String> compaction = new ArrayList<String>();
        fillTitleCellData(sheet,tableNum,  data.get(0),map.get(Integer.valueOf(xuhao)));
        for(JjgFbgcLmgcLmwcLcf row: data){
            if(xuhao.equals(row.getXh())){
                if(index == 50){
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum, row, map.get(Integer.valueOf(xuhao)));
                    index = 0;
                }
                if(row.getBz() == null || "".equals(row.getBz())){
                    fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                    index ++;
                }
                else{
                    if(index > 22){
                        if(index > 47){
                            tableNum ++;
                            fillTitleCellData(sheet, tableNum, row,map.get(Integer.valueOf(xuhao)));
                            index = 0;
                        }else if(index < 25){
                            index = 25;
                        }
                    }
                    fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                    index += 3;
                }
            }
            else{
                xuhao = row.getXh();
                tableNum ++;
                index = 0;
                fillTitleCellData(sheet, tableNum, row,map.get(Integer.valueOf(xuhao)));
                if(row.getBz() == null || "".equals(row.getBz())){
                    fillCommonCellData(sheet, tableNum, index, row,cellstyle);
                    index ++;
                }
                else{
                    if(index > 22){
                        if(index > 47){
                            tableNum ++;
                            fillTitleCellData(sheet, tableNum,  row, map.get(Integer.valueOf(xuhao)));
                            index = 0;
                        }else if(index < 25){
                            index = 25;
                        }
                    }
                    fillCommonCellData(sheet, tableNum, index, row,cellstyle);
                    index += 3;
                }

            }
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     * @param cellstyle
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcLmgcLmwcLcf row, XSSFCellStyle cellstyle) {
        if(row.getBz() == null || "".equals(row.getBz())){
            sheet.getRow(tableNum*34+8+index%25).getCell(6*(index/25)).setCellValue(row.getZh());
            sheet.getRow(tableNum*34+8+index%25).getCell(2+6*(index/25)).setCellValue(row.getCd());
            /*if("Z".equals(row.getGcbw().substring(0,1)) && !row.getCd().startsWith("左")){
                sheet.getRow(tableNum*34+8+index%25).getCell(2+6*(index/25)).setCellValue("左"+row.getCd());
            }
            else if("Y".equals(row.getGcbw().substring(0,1)) && !row.getCd().startsWith("右")){
                sheet.getRow(tableNum*34+8+index%25).getCell(2+6*(index/25)).setCellValue("右"+row.getCd());
            }
            else{
                if(row.getCd().contains(".")){
                    sheet.getRow(tableNum*34+8+index%25).getCell(2+6*(index/25)).setCellValue(Double.valueOf(row.getCd()).intValue());
                }
                else{
                    sheet.getRow(tableNum*34+8+index%25).getCell(2+6*(index/25)).setCellValue(row.getCd());
                }
            }*/
            try{
                sheet.getRow(tableNum*34+8+index%25).getCell(3+6*(index/25)).setCellValue(Double.valueOf(row.getScwcz()));
                sheet.getRow(tableNum*34+8+index%25).getCell(4+6*(index/25)).setCellValue(Double.valueOf(row.getLbwd()));
            }catch(Exception e){
                e.printStackTrace();

            }
        }
        else{
            sheet.getRow(tableNum*34+8+index%25).getCell(6*(index/25)).setCellValue(StringUtils.substringBefore(row.getZh(),"--"));
            sheet.getRow(tableNum*34+8+index%25+1).getCell(6*(index/25)).setCellValue("～");
            sheet.getRow(tableNum*34+8+index%25+2).getCell(6*(index/25)).setCellValue(StringUtils.substringAfter(row.getZh(), "--"));

            try{
                sheet.addMergedRegion(new CellRangeAddress(tableNum*34+8+index%25, tableNum*34+8+index%25+2, 1+6*(index/25), 3+6*(index/25)));
            }catch(Exception e){
                e.printStackTrace();

            }
            sheet.getRow(tableNum*34+8+index%25).getCell(1+6*(index/25)).setCellValue(row.getBz());
            sheet.getRow(tableNum*34+8+index%25).getCell(1+6*(index/25)).setCellStyle(cellstyle);
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param row
     * @param position
     */
    private void fillTitleCellData(XSSFSheet sheet,int tableNum, JjgFbgcLmgcLmwcLcf row, String position) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        sheet.getRow(tableNum * 34 + 1).getCell(1).setCellValue(row.getProname());
        sheet.getRow(tableNum * 34 + 1).getCell(9).setCellValue(row.getHtd());
        sheet.getRow(tableNum*34+2).getCell(1).setCellValue(row.getFbgc());
        sheet.getRow(tableNum*34+2).getCell(9).setCellValue(simpleDateFormat.format(row.getJcsj()));//检测日期
        sheet.getRow(tableNum*34+3).getCell(1).setCellValue(position);//工程部位

        sheet.getRow(tableNum*34+3).getCell(9).setCellType(CellType.NUMERIC);
        sheet.getRow(tableNum*34+3).getCell(9).setCellValue(Double.parseDouble(row.getYswcz()));//(nf2.format(Double.parseDouble(row[7])));//验收弯沉值
        sheet.getRow(tableNum*34+4).getCell(1).setCellValue(row.getJgcc());//结构层次
        if(row.getJjyxxs() == null || "".equals(row.getJjyxxs())){
            sheet.getRow(tableNum*34+4).getCell(5).setCellValue(1);//季节影响系数
        }
        else{
            sheet.getRow(tableNum*34+4).getCell(5).setCellValue(Double.valueOf(row.getJjyxxs()));//季节影响系数
        }

        sheet.getRow(tableNum*34+5).getCell(1).setCellValue(row.getJglx());//结构类型
        sheet.getRow(tableNum*34+5).getCell(3).setCellValue(Double.valueOf(row.getMbkkzb()));//目标可靠指标
        sheet.getRow(tableNum*34+5).getCell(5).setCellValue(Double.valueOf(row.getSdyxxs()));//湿度影响系数
        sheet.getRow(tableNum*34+6).getCell(1).setCellValue(Double.valueOf(row.getClchd()));//材料层厚度Ha
        sheet.getRow(tableNum*34+6).getCell(3).setCellValue(Double.valueOf(row.getLjdmhtml()));//路基顶面回弹模量E0
        sheet.getRow(tableNum*34+6).getCell(5).setCellValue(Double.valueOf(row.getLcz()));//落锤重（T)
        if (row.getBz() == null || "".equals(row.getBz())){
            sheet.getRow(tableNum*34+33).getCell(0).setCellValue("备注：");//备注
        }else {
            sheet.getRow(tableNum*34+33).getCell(0).setCellValue("备注："+row.getBz());//备注
        }

    }

    /**
     *
     * @param sheet
     * @param data
     * @return
     */
    private ArrayList<String[]> createTemperatureSheet(XSSFSheet sheet, List<Map<String, Object>> data) {
        ArrayList<String[]> tem = new ArrayList<String[]>();
        XSSFRow row;
        DecimalFormat nf = new DecimalFormat("0");
        for (int i = 0; i < data.size(); i++) {
            String[] ref = new String[]{"",""};
            row = sheet.getRow(i + 2);

            row.createCell(0).setCellValue(nf.format(i+1));
            row.createCell(1).setCellValue(data.get(i).get("gcbw").toString());
            ref[0] = data.get(i).get("gcbw").toString();//工程部位
            row.createCell(2).setCellValue(Double.parseDouble(data.get(i).get("lqczhd").toString()));
            row.createCell(3).setCellValue(Double.parseDouble(data.get(i).get("lqbmwd").toString()));
            row.createCell(4).setCellValue(Double.parseDouble(data.get(i).get("cqpjwd").toString()));

            row.createCell(5).setCellFormula(
                    "(-2.65+0.52*" + row.getCell(2).getReference()
                            + ")+(0.62-0.008*" + row.getCell(2).getReference()
                            + ")*(" + row.getCell(4).getReference() + "+"
                            + row.getCell(3).getReference() + ")");// =(-2.65+0.52*C3)+(0.62-0.008*C3)*(E3+D3)
            row.createCell(6).setCellFormula(
                    "ROUND(IF((" + row.getCell(3).getReference() + "<=22)*AND("
                            + row.getCell(3).getReference() + ">=18),1,IF("
                            + row.getCell(5).getReference()
                            + ">=20,POWER(2.71828,(1/"
                            + row.getCell(5).getReference()
                            + "-0.05)*C3),POWER(2.71828,0.002*(20-"
                            + row.getCell(5).getReference() + ")*"
                            + row.getCell(2).getReference() + "))),2)");// =IF((D3<=22)*AND(D3>=18),1,IF(F3>=20,POWER(2.71828,(1/F3-0.05)*C3),POWER(2.71828,0.002*(20-F3)*C3)))
            ref[1] = row.getCell(6).getReference();
            tem.add(ref);
        }
        return tem;


    }

    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb,String sheetname) {
        for(int i = 1; i < tableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 33, i*34);
        }
        if(tableNum > 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 10, 0, tableNum*34-1);
        }
    }

    /**
     *
     * @param data
     * @return
     */
    private int gettableNum(List<JjgFbgcLmgcLmwcLcf> data) {
        Map<String, List<JjgFbgcLmgcLmwcLcf>> classifiedData = new HashMap<>();

        for (JjgFbgcLmgcLmwcLcf item : data) {
            String xh = item.getXh();
            // 检查是否已存在该分类，如果不存在则创建一个新的List
            classifiedData.computeIfAbsent(xh, k -> new ArrayList<>()).add(item);
        }
        int num = 0;
        // 分别统计每个List的长度
        for (Map.Entry<String, List<JjgFbgcLmgcLmwcLcf>> entry : classifiedData.entrySet()) {
            String xh = entry.getKey();
            List<JjgFbgcLmgcLmwcLcf> dataList = entry.getValue();
            int dataSize = dataList.size();
            num += dataSize % 50 == 0 ? dataSize / 50 : dataSize / 50 + 1;
        }
        return num;

    }


    /*private int gettableNum(String proname, String htd, String fbgc) {
        QueryWrapper<JjgFbgcLmgcLmwcLcf> queryWrapper = new QueryWrapper();
        queryWrapper.select("DISTINCT xh").like("proname",proname).like("htd",htd).like("fbgc",fbgc);
        List<JjgFbgcLmgcLmwcLcf> xhnum = jjgFbgcLmgcLmwcLcfMapper.selectList(queryWrapper);
        List<Integer> xuhaoList = new ArrayList<>();
        for (JjgFbgcLmgcLmwcLcf x:xhnum){
            xuhaoList.add(Integer.valueOf(x.getXh()));
        }
        int tableNum = 0;
        for(int i = 0; i < xuhaoList.size(); i++){
            QueryWrapper<JjgFbgcLmgcLmwcLcf> queryWrapperxh = new QueryWrapper();
            queryWrapperxh.like("proname",proname).like("htd",htd).like("fbgc",fbgc).like("xh", xuhaoList.get(i));
            List<JjgFbgcLmgcLmwcLcf> jjgFbgcLjgcLjwcs = jjgFbgcLmgcLmwcLcfMapper.selectList(queryWrapperxh);
            int dataNum = jjgFbgcLjgcLjwcs.size();
            tableNum+= dataNum%50 == 0 ? dataNum/50 : dataNum/50+1;
        }
        return tableNum;
    }*/

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo, int flag) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.#");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String title = "弯沉质量鉴定结果汇总表";
        String sheetname = "评定单元";
        String sheetname1 = "路面弯沉左幅";
        String sheetname2 = "路面弯沉右幅";
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "13路面弯沉(落锤法).xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFSheet slSheet1 = xwb.getSheet(sheetname1);

            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(7);//合同段名
            XSSFCell hd = slSheet.getRow(2).getCell(2);//分布工程名
            List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();

            /*
            2025.03.28 修改思路
            在读取“评定单元”的时候，读取数据的时候把“检测路段”给加上，后面流式传输的时候（flag = 2），根据这个检测路段和“验收弯沉值”来进行分类；
            检测路段的识别需要进行字符串的处理
            * */
            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                //取评定单元中的所有数据
                int pddynum = slSheet.getLastRowNum();
                List<Map<String, Object>> tempmapList = new ArrayList<>();
                for (int i = 5;i<pddynum-1;i++){
                    slSheet.getRow(i).getCell(1).setCellType(CellType.STRING);// 检测段落
                    slSheet.getRow(i).getCell(5).setCellType(CellType.STRING);//验收弯沉值（0.01mm）
                    slSheet.getRow(i).getCell(6).setCellType(CellType.STRING);//弯沉代表值
                    slSheet.getRow(i).getCell(7).setCellType(CellType.STRING);//弯沉代表值
                    slSheet.getRow(i).getCell(8).setCellType(CellType.STRING);//弯沉代表值
                    Map<String,Object> maptemp = new HashMap<>();
                    String name = processString(slSheet.getRow(i).getCell(1).getStringCellValue());
                    String value1 = slSheet.getRow(i).getCell(5).getStringCellValue();
                    String value2 = slSheet.getRow(i).getCell(6).getStringCellValue();
                    String value3 = slSheet.getRow(i).getCell(7).getStringCellValue();
                    String value4 = slSheet.getRow(i).getCell(8).getStringCellValue();
                    if (!"".equals(value1) && !"".equals(value2) ){
                        maptemp.put("yswcz",value1);
                        maptemp.put("wcdbz",value2);
                        if ("√".equals(value3)){
                            maptemp.put("wcjl","合格");
                        }else if ("×".equals(value4)){
                            maptemp.put("wcjl","不合格");
                        }
                        maptemp.put("jcdl",name);
                        tempmapList.add(maptemp);
                    }
                }

                // 2025.03.28这段按照“验收弯沉值”进行分组，有点疑惑，现在为了不改变太多的代码，先假定弯沉值是一样的，并且分开两套代码，避免影响flag = 1的时候
                //按yswcz进行分组统计wcdbz的最大值和最小值，以及总点数，合格点数，合格率
                if(flag == 1){
                    Map<String, List<Map<String, Object>>> result = tempmapList.stream()
                            .collect(Collectors.groupingBy(map -> (String) map.get("yswcz")));
                    result.forEach((group, grouphtdData) -> {
                        int zds = grouphtdData.size();
                        int count=0;
                        for (Map<String, Object> grouphtdDatum : grouphtdData) {
                            String wcjl = grouphtdDatum.get("wcjl").toString();
                            if ("合格".equals(wcjl)){
                                count++;
                            }
                        }
                        Map<String, Object> map = new HashMap<>();
                        map.put("max", decf.format(findMaxValue(grouphtdData, "wcdbz")));
                        map.put("min",decf.format(findMinValue(grouphtdData, "wcdbz")));
                        map.put("代表值",decf.format(findMinValue(grouphtdData, "wcdbz"))+"~"+decf.format(findMaxValue(grouphtdData, "wcdbz")));
                        map.put("规定值",group);
                        map.put("检测单元数",zds);
                        map.put("合格单元数",count);
                        map.put("合格率",(zds != 0) ? df.format(count/zds*100) : "0");
                        mapList.add(map);

                    });
                }else if(flag == 2){
                    // 按照验收弯沉值 + 检测段落进行分组，模仿上面写即可了
                    Map<String, List<Map<String, Object>>> result = tempmapList.stream()
                            .collect(Collectors.groupingBy(map -> {
                                // 动态生成分组键
                                String fbgcmc = map.get("jcdl").toString();
                                String yswcz = map.get("yswcz").toString();

                                // 判断是否符合特殊分组条件
                                if (fbgcmc.equals("路面") || fbgcmc.contains("互通")) {
                                    return yswcz + "|SPECIAL"; // 特殊分组标记
                                } else {
                                    return yswcz + "|NORMAL|" + fbgcmc; // 普通分组标记
                                }
                            }));

                    result.forEach((group, grouphtdData) -> {
                        // 公共计算逻辑
                        int zds = grouphtdData.size();
                        long count = grouphtdData.stream()
                                .filter(m -> "合格".equals(m.get("wcjl").toString()))
                                .count();

                        // 解析分组键
                        String[] segments = group.split("\\|");
                        String yswczValue = segments[0];
                        String groupType = segments[1];
                        String fbgcmcValue;

                        // 根据分组类型设置工程名称
                        if ("SPECIAL".equals(groupType)) {
                            // 合并显示特殊分组的名称
                            fbgcmcValue = "路面/互通类";
                        } else {
                            // 普通分组取原始名称
                            fbgcmcValue = segments[2];
                        }

                        Map<String, Object> map = new HashMap<>();
                        // 公共字段
                        map.put("max", decf.format(findMaxValue(grouphtdData, "wcdbz")));
                        map.put("min", decf.format(findMinValue(grouphtdData, "wcdbz")));
                        map.put("代表值", decf.format(findMinValue(grouphtdData, "wcdbz"))
                                + "~" + decf.format(findMaxValue(grouphtdData, "wcdbz")));
                        // 分组维度字段
                        map.put("规定值", yswczValue);
                        map.put("分部工程名称", fbgcmcValue);
                        // 统计字段
                        map.put("检测单元数", zds);
                        map.put("合格单元数", count);
                        map.put("合格率", (zds != 0) ? df.format((count * 100.0)/zds) : "0");

                        mapList.add(map);
                    });
                }


                /*List list = new ArrayList<>();
                int lastRowNum1 = slSheet1.getLastRowNum();
                for (int i = 5;i<lastRowNum1-1;i++){
                    slSheet1.getRow(i).getCell(6).setCellType(CellType.STRING);
                    slSheet1.getRow(i).getCell(9).setCellType(CellType.STRING);
                    String dbz = slSheet1.getRow(i).getCell(6).getStringCellValue();
                    if (dbz.contains("代表弯沉值")){
                        String value = slSheet1.getRow(i).getCell(9).getStringCellValue();
                        list.add(value);
                    }
                }
                int lastRowNumpddy = slSheet.getLastRowNum();
                List<Map<String,Object>> listtemp = new ArrayList<>();
                for (int i = 5;i<lastRowNumpddy-1;i++){
                    slSheet.getRow(i).getCell(5).setCellType(CellType.STRING);//验收弯沉值（0.01mm）
                    slSheet.getRow(i).getCell(6).setCellType(CellType.STRING);//弯沉代表值
                    slSheet.getRow(i).getCell(7).setCellType(CellType.STRING);//弯沉代表值
                    slSheet.getRow(i).getCell(8).setCellType(CellType.STRING);//弯沉代表值
                    Map<String,Object> maptemp = new HashMap<>();
                    String value1 = slSheet.getRow(i).getCell(5).getStringCellValue();
                    String value2 = slSheet.getRow(i).getCell(6).getStringCellValue();
                    String value3 = slSheet.getRow(i).getCell(7).getStringCellValue();
                    String value4 = slSheet.getRow(i).getCell(8).getStringCellValue();
                    if (!"".equals(value1) && !"".equals(value2) ){
                        maptemp.put("yswcz",value1);
                        maptemp.put("wcdbz",value2);
                        if ("√".equals(value3)){
                            maptemp.put("wcjl","合格");
                        }else if ("×".equals(value4)){
                            maptemp.put("wcjl","不合格");
                        }
                        mapList.add(maptemp);
                    }
                }
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//检测单元数
                slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格单元
                slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);//合格率
                slSheet.getRow(5).getCell(5).setCellType(CellType.STRING);//合格率
                jgmap.put("检测单元数", decf.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue())));
                jgmap.put("合格单元数", decf.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue())));
                jgmap.put("合格率", df.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(8).getStringCellValue())));
                jgmap.put("规定值", slSheet.getRow(5).getCell(5).getStringCellValue());
                jgmap.put("代表值", list);
                jgmap.put("listtemp", listtemp);
                mapList.add(jgmap);*/
                return mapList;

            }
            return null;
        }
    }


    // 这个函数是用来处理字符串，如果字符串中没有字母，则返回"路面"，如果有字母，则返回第一个字母之前的部分。
    public static String processString(String s) {
        Pattern p = Pattern.compile("[A-Za-z]");
        Matcher matcher = p.matcher(s);
        if (!matcher.find()) {
            return "路面"; // 没有字母的情况
        }

        int firstLetterPos = matcher.start();
        if (firstLetterPos == 0) {
            return "路面";
        } else {
            String chinesePart = s.substring(0, firstLetterPos).trim();
            return chinesePart.isEmpty() ? "路面" : chinesePart;
        }
    }


    // 查找最大值
    public static double findMaxValue(List<Map<String, Object>> list, String key) {
        double maxValue = Double.valueOf(list.get(0).get("wcdbz").toString()); // 初始化为Double的最小值
        if (list != null || !list.isEmpty() || key != null || !key.isEmpty()) {
            for (Map<String, Object> map : list) {
                if (map.containsKey(key)) {
                    String d = map.get(key).toString();
                    ///if (value instanceof Double) {
                        double currentValue = Double.valueOf(d);
                        if (currentValue > maxValue) {
                            maxValue = currentValue;
                        }
                    //}
                }
            }
        }
        return maxValue;
    }

    // 查找最小值
    public static double findMinValue(List<Map<String, Object>> list, String key) {
        double minValue = Double.valueOf(list.get(0).get("wcdbz").toString()); // 初始化为Double的最大值
        if (list != null || !list.isEmpty() || key != null || !key.isEmpty()) {
            for (Map<String, Object> map : list) {
                if (map.containsKey(key)) {
                    String d = map.get(key).toString();
                    //if (value instanceof Double) {
                    double currentValue = Double.valueOf(d);
                        if (currentValue < minValue) {
                            minValue = currentValue;
                        }
                    //}
                }
            }
        }

        return minValue;
    }

    @Override
    public void exportlmwclcf(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "02路面弯沉(落锤法)实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLmwcLcfVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"02路面弯沉(落锤法)实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }
    }

    @Override
    @Transactional
    public void importlmwclcf(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLmwcLcfVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLmwcLcfVo>(JjgFbgcLmgcLmwcLcfVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLmwcLcfVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcLmgcLmwcLcfVo lmwcLcfVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(lmwcLcfVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(lmwcLcfVo.getYswcz()) || StringUtils.isEmpty(lmwcLcfVo.getYswcz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，验收弯沉值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getMbkkzb())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，目标可靠指标值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getSdyxxs())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，湿度影响系数值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getJjyxxs())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，季节影响系数值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getJglx())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，结构类型为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getJgcc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，结构层次为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(lmwcLcfVo.getLcz()) || StringUtils.isEmpty(lmwcLcfVo.getLcz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，落锤重值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getGcbw())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，工程部位为空，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(lmwcLcfVo.getCd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，车道为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(lmwcLcfVo.getClchd()) || StringUtils.isEmpty(lmwcLcfVo.getClchd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，材料层厚度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getLjdmhtml())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，路基顶面回弹模量值有误，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(lmwcLcfVo.getScwcz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，实测弯沉值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getLbwd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，路表温度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getLqczhd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，沥青层总厚度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getLqbmwd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，沥青表面温度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getCqpjwd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，测前5h平均温度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcLcfVo.getXh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，序号值有误，请修改后重新上传");
                                        }
                                        JjgFbgcLmgcLmwcLcf lmwcLcf = new JjgFbgcLmgcLmwcLcf();
                                        BeanUtils.copyProperties(lmwcLcfVo,lmwcLcf);
                                        lmwcLcf.setCreatetime(new Date());
                                        lmwcLcf.setProname(commonInfoVo.getProname());
                                        lmwcLcf.setUserid(commonInfoVo.getUserid());
                                        lmwcLcf.setHtd(commonInfoVo.getHtd());
                                        lmwcLcf.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLmwcLcfMapper.insert(lmwcLcf);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中02路面弯沉(落锤法)实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }


    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLmgcLmwcLcfMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcLmgcLmwcLcfMapper.selectnumname(proname);
        return selectnum;
    }
}
