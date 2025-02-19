package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcSbJgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcSbJgccVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcSbJgccMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcSbJgccService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
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

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Service
public class JjgFbgcQlgcSbJgccServiceImpl extends ServiceImpl<JjgFbgcQlgcSbJgccMapper, JjgFbgcQlgcSbJgcc> implements JjgFbgcQlgcSbJgccService {

    @Autowired
    private JjgFbgcQlgcSbJgccMapper jjgFbgcQlgcSbJgccMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Autowired
    private SysUserService sysUserService;


    @Override
    public boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcQlgcSbJgcc> wrapper=new QueryWrapper<>();
        wrapper.eq("proname",proname);
        wrapper.eq("htd",htd);
        wrapper.eq("fbgc",fbgc);
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
            wrapper.like("userid",userid);
        }
        wrapper.orderByAsc("qlmc");
        List<JjgFbgcQlgcSbJgcc> data = jjgFbgcQlgcSbJgccMapper.selectList(wrapper);

        //List<Map<String,Object>> selectnum = jjgFbgcQlgcSbJgccMapper.selectnum(proname, htd, fbgc);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"30桥梁上部主要结构尺寸.xlsx");
        if (data == null || data.size()==0){
            return false;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
               /* File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String name = "桥梁上部尺寸.xlsx";
                String path = reportPath + File.separator + name;
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/桥梁上部尺寸.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream in = new FileInputStream(f);
                wb = new XSSFWorkbook(in);

                int tablename = gettableNum(data);
                createTable(tablename,wb);
                if(DBtoExcel(data,wb,tablename)){
                    calculateSizeSheet(wb);
                    /*for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }*/

                    FileOutputStream fileOut = new FileOutputStream(f);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();

                }
                inputStream.close();
                in.close();
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
     * 对“尺寸”sheet进行计算处理
     * @param wb
     */
    private void calculateSizeSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("尺寸");
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        XSSFRow rowrecord = null;
        ArrayList<String> bridgename = new ArrayList<String>();
        ArrayList<XSSFRow> start = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> end = new ArrayList<XSSFRow>();
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 计算
            if (row.getCell(1) != null && !"".equals(row.getCell(1).toString()) && flag) {
                row.getCell(7).setCellFormula(
                        "IF(((" + row.getCell(5).getReference() + "-"
                                + row.getCell(4).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[0]+")*AND(("
                                + row.getCell(4).getReference() + "-"
                                + row.getCell(5).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[1]
                                + "),\"√\",\"\")");
                // H=IF(ABS(F6-E6)<=15,"√","")=IF(ABS(F6-E6)<=20,"√","")
                //=IF(((F6-E6)<=15)*AND((E6-F6)<=15),"√","")
                row.getCell(8).setCellFormula(
                        "IF(((" + row.getCell(5).getReference() + "-"
                                + row.getCell(4).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[0]+")*AND(("
                                + row.getCell(4).getReference() + "-"
                                + row.getCell(5).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[1]
                                + "),\"\",\"×\")");// I=IF(ABS(F6-E6)<=15,"","×")=IF(ABS(F6-E6)<=20,"","×")
                row.createCell(13).setCellFormula(
                        row.getCell(5).getReference() + "-"
                                + row.getCell(4).getReference());// =F6-E6
            }
            // 到下一个表了
            if (row.getCell(1) != null && !"".equals(row.getCell(0).toString())
                    && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (row.getCell(0) != null && !"".equals(row.getCell(0).toString())
                    && !name.equals(getBridgeName(row.getCell(0).toString())) && flag) {
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(sheet, rowrecord, start, end, 5, 7, 8, 9, 10, 11, 12,wb);
                }
                start.clear();
                end.clear();
                name = getBridgeName(row.getCell(0).toString());
                rowstart = sheet.getRow(i);
                rowrecord = rowstart;

                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
                bridgename.add(rowstart.getCell(0).toString());
            }
            // 同一座桥，但是在不同的单元格
            else if (row.getCell(0) != null && !"".equals(row.getCell(0).toString())
                    && name.equals(getBridgeName(row.getCell(0).toString())) && flag
                    && rowstart != row) {
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);

            }
            // 可以计算啦
            if (row.getCell(0) != null && "桥梁   名称".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
                rowrecord = rowstart;
                rowend = sheet.getRow(getCellEndRow(sheet, i + 1, 0));
                if (name.equals(getBridgeName(rowstart.getCell(0).toString()))) {
                    start.add(rowstart);
                    end.add(rowend);

                } else {
                    if (start.size() > 0) {
                        rowrecord = start.get(0);
                        setTotalData(sheet, rowrecord, start, end, 5, 7, 8, 9, 10, 11, 12,wb);
                    }
                    start.clear();
                    end.clear();
                    name = getBridgeName(rowstart.getCell(0).toString());
                    start.add(rowstart);
                    end.add(rowend);
                    bridgename.add(name);
                }
                System.out.println();
            }
        }
        if (start.size() > 0) {
            rowrecord = start.get(0);
            setTotalData(sheet, rowrecord, start, end, 5, 7, 8, 9, 10, 11, 12,wb);
        }

        XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(wb);
        row = sheet.getRow(2);
        row.createCell(9).setCellFormula("SUM("
                +sheet.getRow(3).createCell(9).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(9).getReference()+")");//=SUM(L7:L300)
        double value = e.evaluate(row.getCell(9)).getNumberValue();
        row.getCell(9).setCellFormula(null);
        row.getCell(9).setCellValue(value);


        row.createCell(10).setCellFormula("SUM("
                +sheet.getRow(3).createCell(10).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(10).getReference()+")");//=SUM(L7:L300)

        value = e.evaluate(row.getCell(10)).getNumberValue();
        row.getCell(10).setCellFormula(null);
        row.getCell(10).setCellValue(value);

        sheet.getRow(2).createCell(11).setCellFormula(sheet.getRow(2).getCell(10).getReference()+"*100/"
                +sheet.getRow(2).getCell(9).getReference());//合格率
        value = e.evaluate(row.getCell(11)).getNumberValue();
        row.getCell(11).setCellFormula(null);
        row.getCell(11).setCellValue(value);
    }

    /**
     * 根据传入的字符串，判断允许偏差的数值
     * @param offset
     * @return
     */
    private String[] analysisOffset(String offset) {
        String result[] = new String[2];
        if(offset.contains(",")){
            String [] temp = offset.split("[+,-]");
            if(offset.contains("+") && offset.contains("-")){
                result[0] = temp[1];
                result[1] = temp[temp.length-1];
            }else if(offset.contains("+")){
                result[0] = temp[1];
                result[1] = temp[temp.length-1];
            }else{
                result[0] = temp[0];
                result[1] = temp[temp.length-1];
            }
        }
        else{
            result[0] = offset.substring(1);
            result[1] = result[0];
        }
        return result;
    }

    /**
     * 根据给定的单元格起始行号，得到合并单元格的最后一行行号 如果给定的初始行号不是合并单元格，那么函数返回初始行号
     * @param sheet
     * @param cellstartrow
     * @param cellstartcol
     * @return
     */
    private int getCellEndRow(XSSFSheet sheet, int cellstartrow, int cellstartcol) {
        int sheetmergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetmergerCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            if (cellstartrow == ca.getFirstRow()
                    && cellstartcol == ca.getFirstColumn()) {
                return ca.getLastRow();
            }
        }
        return cellstartrow;
    }

    /**
     *给指定的sheet的指定行rowrecord设置汇总数据
     * @param sheet
     * @param rowrecord
     * @param start
     * @param end
     * @param c1  总点数
     * @param c2  合格点数
     * @param c3  不合格点数
     * @param s1 各个子项的列
     * @param s2
     * @param s3
     * @param s4
     * @param wb
     */
    private void setTotalData(XSSFSheet sheet, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4,XSSFWorkbook wb) {
        XSSFCellStyle cellstyle = wb.createCellStyle();
        cellstyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        String H = "";
        String I = "";
        String J = "";
        //5, 7, 8, 9, 10, 11, 12
        for (int index = 0; index < start.size(); index+=10) {
            H += "COUNT(" + start.get(index).getCell(c1).getReference() + ":"
                    + end.get(index).getCell(c1).getReference() + ")+";
            I += "COUNTIF(" + start.get(index).getCell(c2).getReference() + ":"
                    + end.get(index).getCell(c2).getReference() + ",\"√\")+";
            J += "COUNTIF(" + start.get(index).getCell(c3).getReference() + ":"
                    + end.get(index).getCell(c3).getReference() + ",\"×\")+";
        }
        setTotalTitle(sheet.getRow(rowrecord.getRowNum() - 1), cellstyle, s1, s2, s3, s4);
        rowrecord.createCell(s1).setCellFormula(H.substring(0, H.length() - 1));// H=COUNT(E216:E245)
        rowrecord.getCell(s1).setCellStyle(cellstyle);

        rowrecord.createCell(s2).setCellFormula(I.substring(0, I.length() - 1));// I=COUNTIF(F41:F60,"√")
        rowrecord.getCell(s2).setCellStyle(cellstyle);

        rowrecord.createCell(s3).setCellFormula(J.substring(0, J.length() - 1));// J=COUNTIF(G216:G245,"×")
        rowrecord.getCell(s3).setCellStyle(cellstyle);

        rowrecord.createCell(s4).setCellFormula(
                rowrecord.getCell(s2).getReference() + "/"
                        + rowrecord.getCell(s1).getReference() + "*100");// K==I41/H41*100
        rowrecord.getCell(s4).setCellStyle(cellstyle);

    }

    /**
     *根据给定的行rowtitle，为汇总数据设置标题
     * @param rowtitle
     * @param cellstyle
     * @param s1
     * @param s2
     * @param s3
     * @param s4
     */
    private void setTotalTitle(XSSFRow rowtitle, XSSFCellStyle cellstyle, int s1, int s2, int s3, int s4) {
        rowtitle.createCell(s1).setCellStyle(cellstyle);
        rowtitle.getCell(s1).setCellValue("总点数");
        rowtitle.createCell(s2).setCellStyle(cellstyle);
        rowtitle.getCell(s2).setCellValue("合格点数");
        rowtitle.createCell(s3).setCellStyle(cellstyle);
        rowtitle.getCell(s3).setCellValue("不合格点数");
        rowtitle.createCell(s4).setCellStyle(cellstyle);
        rowtitle.getCell(s4).setCellValue("合格率");
    }

    /**
     *桥梁的名称可能后面带有（1），（2）或者(1)，(2),所以要过滤掉括号后面的序号
     * @param name
     * @return
     */
    private String getBridgeName(String name){
        if(name.contains("(")){
            return name.substring(0,name.indexOf("("));
        }
        else if(name.contains("（")){
            return name.substring(0,name.indexOf("（"));
        }
        return name;
    }

    /**
     * 写入数据
     * @param data
     * @param wb
     * @param totalTable
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcQlgcSbJgcc> data, XSSFWorkbook wb,int totalTable) throws ParseException {
        Collections.sort(data, new Comparator<JjgFbgcQlgcSbJgcc>() {
            @Override
            public int compare(JjgFbgcQlgcSbJgcc o1, JjgFbgcQlgcSbJgcc o2) {
                // 名字相同时按照 qdzh 排序
                String zhys1 = o1.getLbh();
                String zhys2 = o2.getLbh();
                String zh = zhys1.substring(0,1);
                String zh1 = zhys2.substring(0,1);
                String qlmc1 = o1.getQlmc();
                String qlmc2 = o2.getQlmc();
                int cmpq = qlmc1.compareTo(qlmc2);
                if (cmpq != 0) {
                    return cmpq;
                }
                if (zh.contains("Z") || zh.contains("Y") || zh1.contains("Z") || zh1.contains("Y") ){
                    int cmp = zh.compareTo(zh1);
                    if (cmp != 0) {
                        return cmp;
                    }
                    String s = StringUtils.substringBetween(zhys1, zh, "-");
                    String s1 = StringUtils.substringBetween(zhys2, zh1, "-");
                    return s.compareTo(s1);
                }else {
                    return zhys1.compareTo(zhys2);
                }
            }
        });
        System.out.println(data);
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFSheet sheet = wb.getSheet("尺寸");
        String name = data.get(0).getQlmc();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(data.get(0).getJcsj());
        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet, tableNum,data.get(0));
        sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
        for(JjgFbgcQlgcSbJgcc row:data){
            if(name.equals(row.getQlmc())){
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if (index != 0 && index % 30 == 0){
                    sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
                    tableNum++;
                    fillTitleCellData(sheet, tableNum, row);//填写表头
                    index = 0;
                }
                fillCommonCellData(sheet, tableNum, index, row);
                index += 1;
            }else {
                name = row.getQlmc();
                tableNum ++;
                index = 0;
                fillTitleCellData(sheet, tableNum, row);
                sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
                fillCommonCellData(sheet, tableNum, index, row);
                index += 1;

            }
            /*if(name.equals(row.getQlmc())){
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(index/10 == 3){
                    sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    if(totalTable < tableNum){
                        return true;
                    }
                    fillTitleCellData(sheet, tableNum,row);
                    index = 0;
                }
                if(index%10 == 0){
                    sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);
                }
                fillCommonCellData(sheet, tableNum, index+5, row);
                index++;
            }
            else{
                name = row.getQlmc();
                tableNum ++;
                index = 0;
                if(index/10 == 0 || index == 10){
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    index = 10;
                }
                else if(index/10 == 1 || index == 20){
                    index = 20;
                }
                else if(index/10 == 2 || index == 30){
                    sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum,row);
                    index = 0;
                }
                sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);
                fillCommonCellData(sheet, tableNum, index+5, row);
                index ++;

            }*/
        }
        if(totalTable < tableNum){
            return true;
        }
        sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
        return true;
    }

    /**
     * 填写数据
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcQlgcSbJgcc row) {
        sheet.getRow(tableNum*35+5+index).getCell(0).setCellValue(row.getQlmc());
        sheet.getRow(tableNum*35+5+index).getCell(1).setCellValue(row.getLbh());
        sheet.getRow(tableNum*35+5+index).getCell(2).setCellValue(row.getBw());
        sheet.getRow(tableNum*35+5+index).getCell(3).setCellValue(row.getLb());
        sheet.getRow(tableNum*35+5+index).getCell(4).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(tableNum*35+5+index).getCell(5).setCellValue(Double.parseDouble(row.getScz()));

        if((""+row.getYxwcz()).equals(row.getYxwcf())){
            sheet.getRow(tableNum*35+5+index).getCell(6).setCellValue("±"+new Double(Double.parseDouble(row.getYxwcz())).intValue());
        }
        else{
            String piancha = "";
            String piancha_ = "";
            if(new Double(Double.parseDouble(row.getYxwcz())).intValue() != 0){
                piancha = "+"+new Double(Double.parseDouble(row.getYxwcz())).intValue();
            }else{
                piancha = "0";
            }
            if(new Double(Double.parseDouble(row.getYxwcf())).intValue() != 0){
                piancha_ += "-"+new Double(Double.parseDouble(row.getYxwcf())).intValue();
            }else{
                piancha_ = "0";
            }
            sheet.getRow(tableNum*35+5+index).getCell(6).setCellValue(piancha+","+piancha_);
        }
    }

    /**
     * 填写标题
     * @param sheet
     * @param tableNum
     * @param row
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum,JjgFbgcQlgcSbJgcc row) {
        sheet.getRow(tableNum*35+1).getCell(2).setCellValue(row.getProname());
        sheet.getRow(tableNum*35+1).getCell(6).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue("桥梁上部");

    }


    private void createTable(int gettableNum,XSSFWorkbook wb) {
        int record = 0;
        record = gettableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "尺寸", "尺寸", 0, 34, i * 35);
        }
        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex("尺寸"), 0, 8, 0, record * 35-1);
    }

    /**
     * 获取页数
     * @param mapList
     * @return
     */
    private int gettableNum(List<JjgFbgcQlgcSbJgcc> mapList) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (JjgFbgcQlgcSbJgcc map : mapList) {
            String name = map.getQlmc();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%30==0){
                num += value/30;
            }else {
                num += value/30+1;
            }
        }
        return num;
       /* int size = 0;
        int n = 0;
        for (Map<String, Object> m : mapList)
        {
            for (String k : m.keySet()){
                String qlnum = m.get(k).toString();
                Integer nums = Integer.valueOf(qlnum);
                if (nums % 30 == 0){
                    size += nums/30;
                }else {
                    size += nums/30+1;
                }
            }
        }
        return size;*/
    }

    /**
     * 查看鉴定结果
     * @param commonInfoVo
     * @return
     */
    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //String fbgc = commonInfoVo.getFbgc();
        String title = "结构（断面）尺寸质量鉴定表";
        String sheetname = "尺寸";

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"30桥梁上部主要结构尺寸.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(6);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(2);//涵洞1
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())){
                slSheet.getRow(2).getCell(9).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(10).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(11).setCellType(CellType.STRING);


                jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(9).getStringCellValue())));
                jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(10).getStringCellValue())));
                jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(2).getCell(11).getStringCellValue())));
                mapList.add(jgmap);
                return mapList;
            }else {
                return null;
            }

        }
    }

    /**
     * 导出模板文件
     * @param response
     */
    @Override
    public void exportqlsbjgcc(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "06桥梁上部结构尺寸实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcSbJgccVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"06桥梁上部结构尺寸实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    /**
     * 导入数据
     * @param file
     * @param commonInfoVo
     */
    @Override
    @Transactional
    public void importqlsbjgcc(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcSbJgccVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcSbJgccVo>(JjgFbgcQlgcSbJgccVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcSbJgccVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcQlgcSbJgccVo sbJgccVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(sbJgccVo.getQlmc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，桥梁名称为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sbJgccVo.getLbh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，梁板号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sbJgccVo.getBw())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，部位值为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sbJgccVo.getLb())) {
                                            throw new JjgysException(20001,"您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，类别为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(sbJgccVo.getSjz()) || StringUtils.isEmpty(sbJgccVo.getSjz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，设计值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(sbJgccVo.getScz()) || StringUtils.isEmpty(sbJgccVo.getScz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，实测值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(sbJgccVo.getYxwcz()) || StringUtils.isEmpty(sbJgccVo.getYxwcz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，允许误差+值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(sbJgccVo.getYxwcf()) || StringUtils.isEmpty(sbJgccVo.getYxwcf())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，第"+rowNumber+"行的数据中，允许误差-值有误，请修改后重新上传");
                                        }
                                        JjgFbgcQlgcSbJgcc qlgcSbJgcc = new JjgFbgcQlgcSbJgcc();
                                        BeanUtils.copyProperties(sbJgccVo,qlgcSbJgcc);
                                        qlgcSbJgcc.setCreatetime(new Date());
                                        qlgcSbJgcc.setUserid(commonInfoVo.getUserid());
                                        qlgcSbJgcc.setProname(commonInfoVo.getProname());
                                        qlgcSbJgcc.setHtd(commonInfoVo.getHtd());
                                        qlgcSbJgcc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcSbJgccMapper.insert(qlgcSbJgcc);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中06桥梁上部结构尺寸实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }


    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String sheetname = "尺寸";
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "30桥梁上部主要结构尺寸.xlsx");
        if (!f.exists()) {
            return null;
        }else {
            List<Map<String, Object>> data = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(f)) {
                Workbook workbook = WorkbookFactory.create(fis);
                Sheet sheet = workbook.getSheet(sheetname); // 假设数据在第一个工作表中

                DataFormatter dataFormatter = new DataFormatter();
                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();


                Iterator<Row> rowIterator = sheet.iterator();
                int rowNum = 0;
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    rowNum++;
                    if (rowNum > 4 && row !=null) {
                        Cell cell0 = row.getCell(0);
                        Cell cell9 = row.getCell(9);
                        Cell cell10 = row.getCell(10);
                        Cell cell6 = row.getCell(6);

                        if (cell0 != null && cell9 != null) {
                            String cellValue0 = cell0.getStringCellValue();
                            cell6.setCellType(CellType.STRING);
                            String cellValue31 = cell6.getStringCellValue();

                            String cellValue34 = dataFormatter.formatCellValue(cell9, formulaEvaluator);
                            String cellValue35 = dataFormatter.formatCellValue(cell10, formulaEvaluator);

                            //String cellValue34 = String.valueOf(decf.format(cell9.getNumericCellValue()));
                            //String cellValue35 = String.valueOf(decf.format(cell10.getNumericCellValue()));

                            if (!cellValue0.isEmpty() && !cellValue34.isEmpty()) {
                                Map map = new HashMap();
                                map.put("qlmc",cellValue0);
                                map.put("zds",cellValue34);
                                map.put("hgds",cellValue35);
                                map.put("sjqd",cellValue31);

                                data.add(map);
                            }


                        }
                    }

                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return data;

        }
    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcQlgcSbJgccMapper.selectnums(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcQlgcSbJgccMapper.selectnumname(proname);
        return selectnum;
    }

    @Override
    public List<Map<String, Object>> getqlname(String proname, String htd) {
        List<Map<String, Object>> list = jjgFbgcQlgcSbJgccMapper.getqlName(proname,htd);
        return list;
    }

    @Override
    public int createMoreRecords(List<JjgFbgcQlgcSbJgcc> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcQlgcSbJgccMapper, userID);
    }
}
