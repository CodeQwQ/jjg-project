package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLmwcVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLmwcMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcLmwcService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
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
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-04-16
 */
@Service
public class JjgFbgcLmgcLmwcServiceImpl extends ServiceImpl<JjgFbgcLmgcLmwcMapper, JjgFbgcLmgcLmwc> implements JjgFbgcLmgcLmwcService {

    @Autowired
    private JjgFbgcLmgcLmwcMapper jjgFbgcLmgcLmwcMapper;

    @Autowired
    private SysUserService sysUserService;


    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLmgcLmwc> wrapper=new QueryWrapper<>();
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
        wrapper.orderByAsc("xh");
        List<JjgFbgcLmgcLmwc> data = jjgFbgcLmgcLmwcMapper.selectList(wrapper);

        //获取“温度修正”所需要的数据
        //List<Map<String,Object>> wdata = jjgFbgcLmgcLmwcMapper.selectwdata(proname,htd,fbgc);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"13路面弯沉(贝克曼梁法).xlsx");
        if (data == null || data.size()==0){
            return false;
        }else {
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
                /*File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String path =reportPath +File.separator+ "路面弯沉17规范贝克曼梁法.xlsx";
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/路面弯沉17规范贝克曼梁法.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                //ArrayList<String[]> ref = createTemperatureSheet(wb.getSheet("温度修正"),wdata);
                //createTable(gettableNum(proname,htd,fbgc),wb);
                if(DBtoExcel(data,wb)){
                    //calculateTempDate(wb,"路面弯沉左幅");//wb.getSheet("路面弯沉")
                    //calculateTempDate(wb,"路面弯沉右幅");
                    List<Map<String,Object>> totalrefzf = getTotalMark(wb.getSheet("路面弯沉左幅"),"路面弯沉左幅");
                    List<Map<String,Object>> totalrefyf = getTotalMark(wb.getSheet("路面弯沉右幅"),"路面弯沉右幅");
                    totalrefzf.addAll(totalrefyf);
                    String time = getLastTime(wb.getSheet("路面弯沉左幅"));

                    //在评定单元添加合计的一行
                    createEvaluateTable(totalrefzf,wb);
                    //引用数据，然后计算
                    completeTotleTable(wb.getSheet("评定单元"), totalrefzf, time,wb,proname,htd);
                    //completeTotleTable(wb.getSheet("评定单元"), totalrefyf, time,wb,proname,htd,"路面弯沉右幅");

                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }
                    //JjgFbgcCommonUtils.deletwcEmptySheets(wb);
                    ArrayList<String> delsheets = new ArrayList<String>();;
                    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                        String sheetName = wb.getSheetAt(i).getSheetName();
                        if(sheetName.contains("路面弯沉") && "".equals(wb.getSheet(sheetName).getRow(11).getCell(1).getStringCellValue())){
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
    private void createEvaluateTable(List<Map<String,Object>> ref, XSSFWorkbook xwb) {
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
            if ("桩号/部位".equals(row.getCell(0).toString()) && row.getCell(16) != null && !"".equals(row.getCell(16).toString())) {
                Map<String,Object> map = new HashMap<>();
                /*ArrayList<String> ref = new ArrayList<String>();
                ref.add(row.getCell(2).getReference());
                // 验收弯沉值（0.01mm）
                ref.add(row.getCell(11).getReference());
                ref.add(sheet.getRow(i+2).getCell(11).getReference());*/
                map.put(sheetname,row.getCell(2).getReference()+","+row.getCell(11).getReference()+","+sheet.getRow(i+2).getCell(11).getReference());
                result.add(map);
            }
            // 弯沉代表值(0.01mm)
            /*if ("代表弯沉值(0.01mm)".equals(row.getCell(8).toString()) && row.getCell(16) != null && !"".equals(row.getCell(16).toString())) {
                //System.out.println(row.getCell(11).getReference());
                ref.add(row.getCell(11).getReference());
            }*/

        }


        return result;
    }

    /**
     *
     * @param wb
     */
    private void calculateTempDate(XSSFWorkbook wb,String sheetname) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(wb);
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
            if ("桩号/部位".equals(row.getCell(0).toString()) && !name.equals(row.getCell(2).toString())) {
                if("".equals(name)){
                    name = row.getCell(2).toString();
                    left_top = sheet.getRow(i + 7).getCell(18).getReference();
                    continue;
                }
                else{
                    try{
                        right_bottom = sheet.getRow(i - 5).getCell(21).getReference();
                    }
                    catch(Exception e){
                        right_bottom = sheet.getRow(i - 5).createCell(21).getReference();
                    }
                    fillTempDate(evaluator, wb, start_row, i - 3, left_top, right_bottom,sheetname);
                    start_row = i - 3;
                    name = row.getCell(2).toString();
                    left_top = sheet.getRow(i + 7).getCell(18).getReference();
                }
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                continue;
            }
            if (row.getCell(0) != null && row.getCell(0).toString().startsWith("备注")) {
                flag = false;
                continue;
            }
            // 计算左边的值
            if (flag && row.getCell(2).toString() != null && !"".equals(row.getCell(2).toString()))
            {
                row.getCell(4).setCellFormula(
                        "IF(" + row.getCell(2).getReference() + "<>\"\","
                                + row.getCell(2).getReference() + "*2,\"\")");
                row.getCell(5).setCellFormula(
                        "IF(" + row.getCell(3).getReference() + "<>\"\","
                                + row.getCell(3).getReference() + "*2,\"\")");// =IF(C22<>"",C22*2,"")
                row.createCell(18).setCellFormula(row.getCell(4).getReference());// =E22
                row.createCell(19).setCellFormula(row.getCell(5).getReference());
            }
            // 计算右边的值
            if (flag && row.getCell(9).toString() != null && !"".equals(row.getCell(9).toString()))
            {
                row.getCell(11).setCellFormula(
                        "IF(" + row.getCell(9).getReference() + "<>\"\","
                                + row.getCell(9).getReference() + "*2,\"\")");
                row.getCell(12).setCellFormula(
                        "IF(" + row.getCell(10).getReference() + "<>\"\","
                                + row.getCell(10).getReference() + "*2,\"\")");
                row.createCell(20).setCellFormula(row.getCell(11).getReference());
                row.createCell(21).setCellFormula(row.getCell(12).getReference());
            }
        }

        try{
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(21).getReference();
        }
        catch(Exception e){
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).createCell(21).getReference();
        }
        fillTempDate(evaluator, wb, start_row, sheet.getPhysicalNumberOfRows(), left_top, right_bottom,sheetname);


    }

    /**
     *
     * @param evaluator
     * @param start_row
     * @param end_row
     * @param left_top
     * @param right_bottom
     */
    private void fillTempDate(XSSFFormulaEvaluator evaluator, XSSFWorkbook wb, int start_row, int end_row, String left_top, String right_bottom,String sheetname) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        int tables = (end_row - start_row)/36;
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
        row2.getCell(16).setCellFormula(
                "IF(ISERROR(STDEV(" + left_top + ":" + right_bottom
                        + ")),\"\",STDEV(" + left_top + ":" + right_bottom
                        + "))");// =IF(ISERROR(STDEV(S10:V34)),"",STDEV(S10:V34))
        BigDecimal row2value = new BigDecimal(evaluator.evaluate(row2.getCell(16)).getNumberValue());
        row2.getCell(16).setCellFormula(null);
        row2.getCell(16).setCellValue(row2value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 舍前平均值
        row3.getCell(16).setCellFormula(
                "IF(ISERROR(AVERAGE(" + left_top + ":" + right_bottom
                        + ")),\"\",AVERAGE(" + left_top + ":" + right_bottom
                        + "))");// =IF(ISERROR(AVERAGE(S10:V34)),"",AVERAGE(S10:V34))
        BigDecimal row3value = new BigDecimal(evaluator.evaluate(row3.getCell(16)).getNumberValue());
        row3.getCell(16).setCellFormula(null);
        row3.getCell(16).setCellValue(row3value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());

        // 代表值=IF(ISERROR((Q3+F6*Q2)*F5*H5*H6),"",(Q3+F6*Q2)*F5*H5*H6)
        row4.getCell(16).setCellFormula(
                "IF(ISERROR((" + row3.getCell(16).getReference() + "+"
                        + row6.getCell(5).getReference() + "*"
                        + row2.getCell(16).getReference() + ")*"
                        + row5.getCell(5).getReference()+ "*"
                        + row5.getCell(7).getReference()+ "*"
                        + row6.getCell(7).getReference()+ "),\"\",("
                        + row3.getCell(16).getReference() + "+"
                        + row6.getCell(5).getReference() + "*"
                        + row2.getCell(16).getReference() + ")*"
                        + row5.getCell(5).getReference()+ "*"
                        + row5.getCell(7).getReference()+ "*"
                        + row6.getCell(7).getReference()+ ")");// =IF(ISERROR(Q3+C5*Q2*F5),"",(Q3+C5*Q2)*F5*H5)
        // 舍前代表值
        // 特异值下限
        row5.getCell(16).setCellFormula(
                "IF(ISERROR(" + row3.getCell(16).getReference() + "-2*"
                        + row2.getCell(16).getReference() + "),\"\","
                        + row3.getCell(16).getReference() + "-2*"
                        + row2.getCell(16).getReference() + ")");// =IF(ISERROR(Q3-2*Q2),"",Q3-2*Q2)
        // 特异值上限
        row6.getCell(16).setCellFormula(
                "IF(ISERROR(" + row3.getCell(16).getReference() + "+2*"
                        + row2.getCell(16).getReference() + "),\"\","
                        + row3.getCell(16).getReference() + "+2*"
                        + row2.getCell(16).getReference() + ")");// =IF(ISERROR(Q3+2*Q2),"",Q3+2*Q2)
        // 总测点数
        row7.getCell(16).setCellFormula(
                "COUNT(" + left_top + ":" + right_bottom + ")");// =COUNT(S10:V34)
        fillSheetBody(wb,5, row5.getCell(11), row3.getCell(16), row4.getCell(11),
                row8.getCell(16));
        fillSheetBody(wb,6, row6.getCell(11), row4.getCell(16), row4.getCell(11),
                row9.getCell(16));
        fillSheetBody(wb,7, row7.getCell(11), row6.getCell(11), row4.getCell(11));
        XSSFRow t_row1, t_row2, t_row3;
        for (int i = 1; i < tables; i++) {
            t_row1 = sheet.getRow(start_row + i*36 + 4);
            t_row2 = sheet.getRow(start_row + i*36 + 5);
            t_row3 = sheet.getRow(start_row + i*36 + 6);
            fillSheetBody(wb,5, t_row1.getCell(11), row3.getCell(16), row4.getCell(11),
                    row8.getCell(16));
            fillSheetBody(wb,6, t_row2.getCell(11), row4.getCell(16), row4.getCell(11),
                    row9.getCell(16));
            fillSheetBody(wb,7, t_row3.getCell(11), row6.getCell(11), row4.getCell(11));
        }
    }

    /**
     *
     * @param rowNum
     * @param cell
     */
    public void fillSheetBody(XSSFWorkbook xwb,int rowNum, XSSFCell... cell) {
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
     * @param sheet
     * @param ref
     * @param time
     * @param xwb
     * @param proname
     * @param htd
     */
    private void completeTotleTable(XSSFSheet sheet, List<Map<String,Object>> ref, String time, XSSFWorkbook xwb, String proname, String htd) {
        //[{路面弯沉左幅=C4,L4,L6}, {路面弯沉左幅=C40,L40,L42}]
        String sheetname = "";
        XSSFCellStyle cellstyle = xwb.createCellStyle();
        XSSFFont font=xwb.createFont();
        font.setFontHeightInPoints((short)11);
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

        XSSFCellStyle cellstyle2 = xwb.createCellStyle();
        XSSFFont font2=xwb.createFont();
        font2.setFontHeightInPoints((short)9);
        font2.setFontName("宋体");
        cellstyle2.setFont(font2);
        cellstyle2.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle2.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle2.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle2.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle2.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle2.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle2.setWrapText(true);//自动换行
        //XSSFRow row;
        int index = 5, rowstart = 5;
        int i = 1;
        /*
         * 填写“评定单元”sheet的表头数据
         */
        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(7).setCellValue(htd);
        sheet.getRow(2).getCell(2).setCellValue("路面工程");
        sheet.getRow(2).getCell(7).setCellValue(time);
        String keyToCheck1 = "路面弯沉左幅";
        String keyToCheck2 = "路面弯沉右幅";
        for (Map<String, Object> map : ref) {
            if (map.containsKey(keyToCheck1)){
                String[] split = map.get(keyToCheck1).toString().split(",");
                sheet.getRow(index).getCell(0).setCellValue(String.valueOf(i));
                sheet.getRow(index).getCell(1).setCellFormula(keyToCheck1+"!" + split[0]);
                sheet.addMergedRegion(new CellRangeAddress(index, index, 1, 4));
                sheet.getRow(index).getCell(5).setCellFormula(keyToCheck1+"!" + split[1]);
                sheet.getRow(index).getCell(6).setCellFormula(keyToCheck1+"!" + split[2]);
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
            }else if (map.containsKey(keyToCheck2)){
                String[] split = map.get(keyToCheck2).toString().split(",");
                sheet.getRow(index).getCell(0).setCellValue(String.valueOf(i));
                sheet.getRow(index).getCell(1).setCellFormula(keyToCheck2+"!" + split[0]);
                sheet.addMergedRegion(new CellRangeAddress(index, index, 1, 4));
                sheet.getRow(index).getCell(5).setCellFormula(keyToCheck2+"!" + split[1]);
                sheet.getRow(index).getCell(6).setCellFormula(keyToCheck2+"!" + split[2]);
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
                        + sheet.getRow(sheet.getLastRowNum()).getCell(4).getReference() + "*100");// =G28/E28*100
        /*for (int j = rowstart; j <= sheet.getPhysicalNumberOfRows(); j++) {
            XSSFRow row = sheet.getRow(j);
            System.out.println(row.getRowNum());
            if ("合计".equals(row.getCell(0).toString())) {
                row.getCell(4).setCellFormula(
                        "COUNT("
                                + sheet.getRow(rowstart).getCell(5)
                                .getReference() + ":"
                                + sheet.getRow(j - 1).getCell(5).getReference()
                                + ")");// =COUNT(F6:F27)

                row.getCell(6).setCellFormula(
                        "COUNTIF("
                                + sheet.getRow(rowstart).getCell(7)
                                .getReference() + ":"
                                + sheet.getRow(j - 1).getCell(7).getReference()
                                + ",\"√\")");// =COUNTIF(H6:H27,"√")

                row.getCell(8).setCellFormula(
                        row.getCell(6).getReference() + "/"
                                + row.getCell(4).getReference() + "*100");// =G28/E28*100
            }
        }*/


        /*double value = 0.0;
        for (int i = rowstart; i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (index < ref.size()) {
                String[] arr= new String[]{"路面弯沉左幅","路面弯沉右幅"};
                for (String s : arr) {
                    String[] split = ref.get(index).get(s).toString().split(",");
                    row.createCell(0).setCellValue(String.valueOf(index/3+1));
                    row.getCell(0).setCellStyle(cellstyle);

                    row.createCell(1).setCellFormula(s+"!" + split[0]);
                    sheet.addMergedRegion(new CellRangeAddress(i, i, 1, 4));
                    row.getCell(1).setCellStyle(cellstyle2);

                    row.createCell(5).setCellFormula(s+"!" + split[1]);
                    row.getCell(5).setCellStyle(cellstyle);
                    row.createCell(6).setCellFormula(s+"!" + split[2]);
                    row.getCell(6).setCellStyle(cellstyle);
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
                        row.getCell(8).setCellStyle(cellstyle);
                    }
                    row.getCell(8).setCellStyle(cellstyle);
                }
                index += 3;

            }
            if ("合计".equals(row.getCell(0).toString())) {
                //System.out.println("zuihouyihang");
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
     * @param data
     * @param wb
     * @return
     */
    private boolean DBtoExcel(List<JjgFbgcLmgcLmwc> data, XSSFWorkbook wb) {
        Collections.sort(data, new Comparator<JjgFbgcLmgcLmwc>() {
            @Override
            public int compare(JjgFbgcLmgcLmwc o1, JjgFbgcLmgcLmwc o2) {
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
        List<JjgFbgcLmgcLmwc> zfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmwc> yfdata = new ArrayList<>();
        for (JjgFbgcLmgcLmwc datum : data) {
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
    private void writezyfdata(List<JjgFbgcLmgcLmwc> data, XSSFWorkbook wb,String sheetname) {
        Collections.sort(data, new Comparator<JjgFbgcLmgcLmwc>() {
            @Override
            public int compare(JjgFbgcLmgcLmwc o1, JjgFbgcLmgcLmwc o2) {
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
        DecimalFormat nf = new DecimalFormat(".00");
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
        //int i = gettableNum(data.get(0).getProname(), data.get(0).getHtd(), data.get(0).getFbgc());
        int num = gettableNum(data);
        createTable(num,wb,sheetname);
        //createTable(gettableNum(data.get(0).getProname(),data.get(0).getHtd(),data.get(0).getFbgc()),wb,sheetname);

        String gcbw = data.get(0).getGcbw();
        String xuhao = data.get(0).getXh();
        HashMap<Integer, String> map = new HashMap<Integer, String>();
        String zhuanghao = gcbw+";\n";
        for(JjgFbgcLmgcLmwc row: data){
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
        fillTitleCellData(wb, tableNum, data.get(0), map.get(Integer.valueOf(xuhao)),sheetname);
        for(JjgFbgcLmgcLmwc row: data){
            if(xuhao.equals(row.getXh())){
                if(index == 50){
                    tableNum ++;
                    fillTitleCellData(wb, tableNum, row, map.get(Integer.valueOf(xuhao)),sheetname);
                    index = 0;
                }
                if(row.getBz() == null || "".equals(row.getBz())){
                    fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                    index ++;
                } else{
                    if(index > 22){
                        if(index > 47){
                            tableNum ++;
                            fillTitleCellData(wb, tableNum, row, map.get(Integer.valueOf(xuhao)),sheetname);
                            index = 0;
                        }else if(index < 25){
                            index = 25;
                        }
                    }
                    fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                    index += 3;
                }
            } else{
                xuhao = row.getXh();
                tableNum ++;
                index = 0;
                fillTitleCellData(wb, tableNum, row, map.get(Integer.valueOf(xuhao)),sheetname);
                if(row.getBz() == null || "".equals(row.getBz())){
                    fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                    index ++;
                }
                else{
                    if(index > 22){
                        if(index > 47){
                            tableNum ++;
                            fillTitleCellData(wb, tableNum, row, map.get(Integer.valueOf(xuhao)),sheetname);
                            index = 0;
                        }else if(index < 25){
                            index = 25;
                        }
                    }
                    fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                    index += 3;
                }

            }
        }
    }

    /**
     *
     * @param data
     * @return
     */
    private int gettableNum(List<JjgFbgcLmgcLmwc> data) {
        Map<String, List<JjgFbgcLmgcLmwc>> classifiedData = new HashMap<>();

        for (JjgFbgcLmgcLmwc item : data) {
            String xh = item.getXh();
            // 检查是否已存在该分类，如果不存在则创建一个新的List
            classifiedData.computeIfAbsent(xh, k -> new ArrayList<>()).add(item);
        }
        int num = 0;
        // 分别统计每个List的长度
        for (Map.Entry<String, List<JjgFbgcLmgcLmwc>> entry : classifiedData.entrySet()) {
            String xh = entry.getKey();
            List<JjgFbgcLmgcLmwc> dataList = entry.getValue();
            int dataSize = dataList.size();
            num += dataSize % 50 == 0 ? dataSize / 50 : dataSize / 50 + 1;
        }
        return num;

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     * @param cellstyle
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcLmgcLmwc row, XSSFCellStyle cellstyle) {
        if(row.getBz() == null || "".equals(row.getBz())){
            sheet.getRow(tableNum*36+10+index%25).getCell(7*(index/25)).setCellValue(row.getZh());
            sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellValue(row.getCd());
            /*if("Z".equals(row.getGcbw().substring(0,1)) && !row.getCd().startsWith("左")){
                sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellValue("左"+row.getCd());
            }
            else if("Y".equals(row.getGcbw().substring(0,1)) && !row.getCd().startsWith("右")){
                sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellValue("右"+row.getCd());
            }
            else{
                if(row.getCd().contains(".")){
                    sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellValue(Double.valueOf(row.getCd()).intValue());
                }
                else{
                    sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellValue(row.getCd());
                }
            }*/

            try{
                sheet.getRow(tableNum*36+10+index%25).getCell(2+7*(index/25)).setCellValue(Double.valueOf(row.getZz()).intValue());
                sheet.getRow(tableNum*36+10+index%25).getCell(3+7*(index/25)).setCellValue(Double.valueOf(row.getYz()).intValue());
                sheet.getRow(tableNum*36+10+index%25).getCell(6+7*(index/25)).setCellValue(Double.valueOf(row.getLqbmwd()).intValue());
            }catch(Exception e){
                e.printStackTrace();

            }
        }
        else{
            sheet.getRow(tableNum*36+10+index%25).getCell(7*(index/25)).setCellValue(StringUtils.substringBefore(row.getZh(),"--"));
            sheet.getRow(tableNum*36+10+index%25+1).getCell(7*(index/25)).setCellValue("～");
            sheet.getRow(tableNum*36+10+index%25+2).getCell(7*(index/25)).setCellValue(StringUtils.substringAfter(row.getZh(), "--"));

            try{
                sheet.addMergedRegion(new CellRangeAddress(tableNum*36+10+index%25, tableNum*36+10+index%25+2, 1+7*(index/25), 6+7*(index/25)));
            }catch(Exception e){
                e.printStackTrace();

            }
            sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellValue(row.getBz());
            sheet.getRow(tableNum*36+10+index%25).getCell(1+7*(index/25)).setCellStyle(cellstyle);
        }
    }

    /**
     *
     * @param wb
     * @param tableNum
     * @param row
     * @param position
     */
    private void fillTitleCellData(XSSFWorkbook wb, int tableNum, JjgFbgcLmgcLmwc row, String position,String sheetname) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        System.out.println(tableNum);
        sheet.getRow(tableNum * 36 + 1).getCell(2).setCellValue(row.getProname());
        sheet.getRow(tableNum * 36 + 1).getCell(11).setCellValue(row.getHtd());
        sheet.getRow(tableNum*36+2).getCell(2).setCellValue(row.getFbgc());
        sheet.getRow(tableNum*36+2).getCell(11).setCellValue(simpleDateFormat.format(row.getJcsj()));//检测日期
        sheet.getRow(tableNum*36+3).getCell(2).setCellValue(position);//工程部位
        sheet.getRow(tableNum*36+3).getCell(11).setCellType(CellType.NUMERIC);
        sheet.getRow(tableNum*36+3).getCell(11).setCellValue(Double.parseDouble(row.getYswcz()));//(nf2.format(Double.parseDouble(row[7])));//验收弯沉值
        sheet.getRow(tableNum*36+4).getCell(2).setCellValue(row.getJgcc());//结构层次
        sheet.getRow(tableNum*36+4).getCell(5).setCellValue(getwdxz(wb,row));//温度影响系数

        sheet.getRow(tableNum*36+5).getCell(5).setCellValue(Double.valueOf(row.getMbkkzb()));//保证率系数λ  目标可靠指标
        sheet.getRow(tableNum*36+5).getCell(7).setCellValue(Double.valueOf(row.getSdyxxs()));//湿度影响系数
        if(row.getJjyxxs() == null || "".equals(row.getJjyxxs())){
            sheet.getRow(tableNum*36+4).getCell(7).setCellValue(1);//季节影响系数
        }
        else{
            sheet.getRow(tableNum*36+4).getCell(7).setCellValue(Double.valueOf(row.getJjyxxs()));//季节影响系数
        }
        //材料层厚度Ha
        sheet.getRow(tableNum*36+6).getCell(2).setCellValue(Double.valueOf(row.getClchd()));
        //平衡湿度路基顶面回弹模量E0
        sheet.getRow(tableNum*36+6).getCell(6).setCellValue(Double.valueOf(row.getPhsdljdmhtml()));
        sheet.getRow(tableNum*36+5).getCell(2).setCellValue(row.getJglx());//结构类型
        sheet.getRow(tableNum*36+7).getCell(2).setCellValue(Double.valueOf(row.getHzz()));//后轴重
        sheet.getRow(tableNum*36+7).getCell(6).setCellValue(Double.valueOf(row.getLtqy()));//轮胎气压
        if (row.getBz() == null || "".equals(row.getBz())){
            sheet.getRow(tableNum*36+35).getCell(0).setCellValue("备注：");//备注
        }else {
            sheet.getRow(tableNum*36+35).getCell(0).setCellValue("备注："+row.getBz());//备注
        }

    }

    /**
     *
     * @param wb
     * @param row
     * @return
     */
    private double getwdxz(XSSFWorkbook wb,JjgFbgcLmgcLmwc row) {
        double res = 0.0;
        XSSFSheet sheet = wb.getSheet("温度修正");
        XSSFRow cellrow = sheet.getRow(2);
        cellrow.createCell(2).setCellValue(Double.parseDouble(row.getLqczhd()));
        cellrow.createCell(3).setCellValue(Double.parseDouble(row.getLqbmwd()));
        cellrow.createCell(4).setCellValue(Double.parseDouble(row.getCqpjwd()));

        cellrow.createCell(5).setCellFormula(
                "(-2.65+0.52*" + cellrow.getCell(2).getReference()
                        + ")+(0.62-0.008*" + cellrow.getCell(2).getReference()
                        + ")*(" + cellrow.getCell(4).getReference() + "+"
                        + cellrow.getCell(3).getReference() + ")");// =(-2.65+0.52*C3)+(0.62-0.008*C3)*(E3+D3)
        cellrow.createCell(6).setCellFormula(
                "ROUND(IF((" + cellrow.getCell(3).getReference() + "<=22)*AND("
                        + cellrow.getCell(3).getReference() + ">=18),1,IF("
                        + cellrow.getCell(5).getReference()
                        + ">=20,POWER(2.71828,(1/"
                        + cellrow.getCell(5).getReference()
                        + "-0.05)*C3),POWER(2.71828,0.002*(20-"
                        + cellrow.getCell(5).getReference() + ")*"
                        + cellrow.getCell(2).getReference() + "))),2)");// =IF((D3<=22)*AND(D3>=18),1,IF(F3>=20,POWER(2.71828,(1/F3-0.05)*C3),POWER(2.71828,0.002*(20-F3)*C3)))
        XSSFFormulaEvaluator evaluator=new XSSFFormulaEvaluator(wb);
        CellValue tempCellValue = evaluator.evaluate(cellrow.getCell(6));
        res = tempCellValue.getNumberValue();
        return res;
    }

    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb,String sheetname) {
        for(int i = 1; i < tableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 35, i*36);
        }
        if(tableNum > 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 13, 0, tableNum*36-1);
        }
    }

    /**
     * 获取要生成页数的数量
     * @param proname
     * @param htd
     * @param fbgc
     * @return
     */
    /*private int gettableNum(String proname, String htd, String fbgc) {
        QueryWrapper<JjgFbgcLmgcLmwc> queryWrapper = new QueryWrapper();
        queryWrapper.select("DISTINCT xh").like("proname",proname).like("htd",htd).like("fbgc",fbgc);
        List<JjgFbgcLmgcLmwc> xhnum = jjgFbgcLmgcLmwcMapper.selectList(queryWrapper);
        List<Integer> xuhaoList = new ArrayList<>();
        for (JjgFbgcLmgcLmwc x:xhnum){
            xuhaoList.add(Integer.valueOf(x.getXh()));
        }
        int tableNum = 0;

        for(int i = 0; i < xuhaoList.size(); i++){
            QueryWrapper<JjgFbgcLmgcLmwc> queryWrapperxh = new QueryWrapper();
            queryWrapperxh.like("proname",proname).like("htd",htd).like("fbgc",fbgc).like("xh", xuhaoList.get(i));
            List<JjgFbgcLmgcLmwc> lmgcLmwcs = jjgFbgcLmgcLmwcMapper.selectList(queryWrapperxh);
            int num=0;
            for (int j=0;j<lmgcLmwcs.size();j++){
                if (lmgcLmwcs.get(i).getBz() != null || "".equals(lmgcLmwcs.get(i).getBz())){
                    num+=2;
                }
            }
            int dataNum = lmgcLmwcs.size()+num;
            tableNum+= dataNum%50 == 0 ? dataNum/50 : dataNum/50+1;
        }
        return tableNum;

    }*/

    /**
     *
     * @param sheet
     * @param wdata
     * @return
     */
    private ArrayList<String[]> createTemperatureSheet(XSSFSheet sheet, List<Map<String, Object>> wdata) {
        ArrayList<String[]> tem = new ArrayList<String[]>();
        XSSFRow row;
        DecimalFormat nf = new DecimalFormat("0");
        for (int i = 0; i < wdata.size(); i++) {
            String[] ref = new String[]{"",""};
            row = sheet.getRow(i + 2);

            row.createCell(0).setCellValue(nf.format(i+1));//序号
            row.createCell(1).setCellValue(wdata.get(i).get("gcbw").toString());//工程部位
            ref[0] = wdata.get(i).get("gcbw").toString();//工程部位
            row.createCell(2).setCellValue(Double.parseDouble(wdata.get(i).get("lqczhd").toString()));
            row.createCell(3).setCellValue(Double.parseDouble(wdata.get(i).get("lqbmwd").toString()));
            row.createCell(4).setCellValue(Double.parseDouble(wdata.get(i).get("cqpjwd").toString()));

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
     * @param commonInfoVo
     * @return
     */
    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String title = "弯沉质量鉴定结果汇总表";
        String sheetname = "评定单元";
        String sheetname1 = "路面弯沉左幅";
        String sheetname2 = "路面弯沉右幅";
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "13路面弯沉(贝克曼梁法).xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFSheet slSheet1 = xwb.getSheet(sheetname1);
            XSSFSheet slSheet2 = xwb.getSheet(sheetname2);

            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(7);//合同段名

            List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();
            DecimalFormat df = new DecimalFormat("0.00");
            DecimalFormat decf = new DecimalFormat("0.#");

            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                int pddynum = slSheet.getLastRowNum();
                List<Map<String, Object>> tempmapList = new ArrayList<>();
                for (int i = 5;i<pddynum-1;i++){
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
                        tempmapList.add(maptemp);
                    }
                }
                //按yswcz进行分组统计wcdbz的最大值和最小值，以及总点数，合格点数，合格率
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
                /*List list = new ArrayList<>();
                int lastRowNum1 = slSheet1.getLastRowNum();
                for (int i = 5;i<lastRowNum1-1;i++){
                    slSheet1.getRow(i).getCell(8).setCellType(CellType.STRING);
                    slSheet1.getRow(i).getCell(11).setCellType(CellType.STRING);
                    String dbz = slSheet1.getRow(i).getCell(8).getStringCellValue();
                    if (dbz.contains("代表弯沉值")){
                        String value = slSheet1.getRow(i).getCell(11).getStringCellValue();
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


    /**
     *
     * @param response
     */
    @Override
    public void exportlmwc(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "02路面弯沉(贝克曼梁法)实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLmwcVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"02路面弯沉(贝克曼梁法)实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    /**
     *
     * @param file
     * @param commonInfoVo
     */
    @Override
    @Transactional
    public void importlmwc(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLmwcVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLmwcVo>(JjgFbgcLmgcLmwcVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLmwcVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcLmgcLmwcVo lmwcVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(lmwcVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(lmwcVo.getYswcz()) || StringUtils.isEmpty(lmwcVo.getYswcz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，验收弯沉值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getMbkkzb())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，目标可靠指标值有误，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(lmwcVo.getSdyxxs())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，湿度影响系数值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getJjyxxs())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，季节影响系数值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getJglx())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，结构类型为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getJgcc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，结构层次为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(lmwcVo.getHzz()) || StringUtils.isEmpty(lmwcVo.getHzz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，后轴重值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getGcbw())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，工程部位为空，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(lmwcVo.getCd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，车道为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(lmwcVo.getClchd()) || StringUtils.isEmpty(lmwcVo.getClchd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，材料层厚度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getPhsdljdmhtml())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，平衡湿度路基顶面回弹模量值有误，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(lmwcVo.getZz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，左值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getYz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，右值有误，请修改后重新上传");
                                        }

                                        if (StringUtils.isEmpty(lmwcVo.getLtqy())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，轮胎气压值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getLbwd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，路表温度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getLqczhd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，沥青层总厚度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getLqbmwd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，沥青表面温度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getCqpjwd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，测前5h平均温度值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmwcVo.getXh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，第"+rowNumber+"行的数据中，序号值有误，请修改后重新上传");
                                        }
                                        JjgFbgcLmgcLmwc lmgcLmwc = new JjgFbgcLmgcLmwc();
                                        BeanUtils.copyProperties(lmwcVo,lmgcLmwc);
                                        lmgcLmwc.setCreatetime(new Date());
                                        lmgcLmwc.setUserid(commonInfoVo.getUserid());
                                        lmgcLmwc.setProname(commonInfoVo.getProname());
                                        lmgcLmwc.setHtd(commonInfoVo.getHtd());
                                        lmgcLmwc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLmwcMapper.insert(lmgcLmwc);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中02路面弯沉(贝克曼梁法)实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }


    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLmgcLmwcMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcLmgcLmwcMapper.selectnumname(proname);
        return selectnum;
    }

    @Override
    public int createMoreRecords(List<JjgFbgcLmgcLmwc> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcLmgcLmwcMapper, userID);
    }
}
