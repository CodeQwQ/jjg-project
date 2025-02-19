package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcSdgcDmpzd;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJathlqdVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcDmpzdVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcDmpzdMapper;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.service.JjgFbgcSdgcDmpzdService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
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
 * @since 2023-03-26
 */
@Service
public class JjgFbgcSdgcDmpzdServiceImpl extends ServiceImpl<JjgFbgcSdgcDmpzdMapper, JjgFbgcSdgcDmpzd> implements JjgFbgcSdgcDmpzdService {

    @Autowired
    private JjgFbgcSdgcDmpzdMapper jjgFbgcSdgcDmpzdMapper;


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
        //获取数据
        QueryWrapper<JjgFbgcSdgcDmpzd> wrapper=new QueryWrapper<>();
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
        wrapper.orderByAsc("sdmc","zh","FIELD(jcbw, '左边墙','右边墙')");
        List<JjgFbgcSdgcDmpzd> data = jjgFbgcSdgcDmpzdMapper.selectList(wrapper);

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"40隧道大面平整度.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
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
                /*File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String name = "隧道大面平整度.xlsx";
                String path = reportPath + File.separator + name;
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/隧道大面平整度.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                createTable(gettableNum(data),wb);
                if(DBtoExcel(data,wb)){
                    calculatePZDSheet(wb);
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
     * @param wb
     */
    private void calculatePZDSheet(XSSFWorkbook wb) {
        String sheetname = "二衬大面平整度";
        XSSFSheet sheet = wb.getSheet(sheetname);
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
            if (!"".equals(row.getCell(4).toString()) && flag) {
                row.getCell(5).setCellFormula("IF("
                        +row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference()+"<=0,\"√\",\"\")");//F6=IF(E6-D6<=0,"√","")
                row.getCell(6).setCellFormula("IF("
                        +row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference()+">0,\"×\",\"\")");//G6=IF(E6-D6>0,"×","")
            }
            // 到下一个表了
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (flag && !"".equals(row.getCell(0).toString()) && !name.equals(row.getCell(0).toString())) {
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(wb,sheetname, rowrecord, start, end, 4, 5, 6,7,8, 9,10);
                }
                start.clear();
                end.clear();
                name = row.getCell(0).toString();
                rowstart = sheet.getRow(i);
                rowrecord = rowstart;

                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
            }
            // 同一座桥，但是在不同的单元格
            else if (flag && !"".equals(row.getCell(0).toString())
                    && name.equals(row.getCell(0).toString())
                    && rowstart != row) {
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
            }
            // 可以计算啦
            if ("隧道名称".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
                rowrecord = rowstart;
                rowend = sheet.getRow(getCellEndRow(sheet, i + 1, 0));
                if (name.equals(rowstart.getCell(0).toString())) {
                    start.add(rowstart);
                    end.add(rowend);
                } else {
                    if (start.size() > 0) {
                        rowrecord = start.get(0);
                        setTotalData(wb,sheetname, rowrecord, start, end, 4, 5, 6,7,8, 9,10);
                    }
                    start.clear();
                    end.clear();
                    name = rowstart.getCell(0).toString();
                    start.add(rowstart);
                    end.add(rowend);
                    bridgename.add(rowstart.getCell(0).toString());
                }
            }
        }
        if (start.size() > 0) {
            rowrecord = start.get(0);
            setTotalData(wb,sheetname, rowrecord, start, end, 4, 5, 6,7,8, 9,10);
        }
    }

    /**
     *
     * @param sheet
     * @param cellstartrow
     * @param cellstartcol
     * @return
     */
    public int getCellEndRow(XSSFSheet sheet, int cellstartrow, int cellstartcol) {
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
     *
     * @param wb
     * @param sheetname
     * @param rowrecord
     * @param start
     * @param end
     * @param c1
     * @param c2
     * @param c3
     * @param s1
     * @param s2
     * @param s3
     * @param s4
     */
    public void setTotalData(XSSFWorkbook wb,String sheetname, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFCellStyle cellstyle = wb.createCellStyle();
        cellstyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        String H = "";
        String I = "";
        String J = "";
        for (int index = 0; index < start.size(); index++) {
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
     *
     * @param rowtitle
     * @param cellstyle
     * @param s1
     * @param s2
     * @param s3
     * @param s4
     */
    public void setTotalTitle(XSSFRow rowtitle, XSSFCellStyle cellstyle, int s1, int s2, int s3, int s4) {
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
     *
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcSdgcDmpzd> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("二衬大面平整度");
        String testtime = simpleDateFormat.format(data.get(0).getJcsj());
        String name = data.get(0).getSdmc();
        String zy = data.get(0).getZh().substring(0,1);
        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet, tableNum, data.get(0));
        for(JjgFbgcSdgcDmpzd row:data){
            if(name.equals(row.getSdmc()) && zy.equals(row.getZh().substring(0,1))){
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(index/10 == 3){
                    sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum, row);
                    index = 0;
                }
                if(index == 0){
                    sheet.getRow(tableNum*35 + 5).getCell(0).setCellValue(name);//setCellValue(name+"("+(ZY.equals("Z")?"左线":"右线")+")");
                }
                fillCommonCellData(sheet, tableNum, index+5, row);
                index++;
            }
            else{
                name = row.getSdmc();
                zy = row.getZh().substring(0,1);
                index = 0;
                sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet, tableNum, row);
                index = 0;
                sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);//setCellValue(name+"("+(ZY.equals("Z")?"左线":"右线")+")");
                fillCommonCellData(sheet, tableNum, index+5, row);
                index ++;
            }
        }
        sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
        return true;

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param row
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum,JjgFbgcSdgcDmpzd row) {
        sheet.getRow(tableNum*35+1).getCell(1).setCellValue(row.getProname());
        sheet.getRow(tableNum*35+1).getCell(5).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(1).setCellValue(row.getFbgc());

    }


    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcSdgcDmpzd row) {
        /*if((index+1)%6 == 0){
            sheet.getRow(tableNum*35+index).getCell(1).setCellValue(row.getZh());
        }*/
        sheet.getRow(tableNum*35+index).getCell(1).setCellValue(row.getZh());
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(row.getJcbw());
        /*if((index+1)%3 == 0 && (index+1)%6 < 3){
            sheet.getRow(tableNum*35+index).getCell(2).setCellValue("左边墙");
        }else if((index+1)%3 == 0 && (index+1)%6 >= 3){
            sheet.getRow(tableNum*35+index).getCell(2).setCellValue("右边墙");
        }*/
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.valueOf(row.getYxps()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.valueOf(row.getScz()));
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
            RowCopy.copyRows(wb, "二衬大面平整度", "二衬大面平整度", 0, 34, i * 35);
        }
        if(record > 1)
            wb.setPrintArea(wb.getSheetIndex("二衬大面平整度"), 0, 6, 0, record * 35-1);
    }

    /**
     *
     * @param data
     * @return
     */
    private int gettableNum(List<JjgFbgcSdgcDmpzd> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (JjgFbgcSdgcDmpzd map : data) {
            String name = map.getSdmc()+map.getZh().substring(0,1);
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
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //String fbgc = commonInfoVo.getFbgc();
        String sheetname = "二衬大面平整度";

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"40隧道大面平整度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(5);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(1);//涵洞
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && htd.equals(htdname.toString())){
                slSheet.getRow(5).getCell(7).setCellType(CellType.STRING);
                slSheet.getRow(5).getCell(8).setCellType(CellType.STRING);
                slSheet.getRow(5).getCell(9).setCellType(CellType.STRING);
                slSheet.getRow(5).getCell(10).setCellType(CellType.STRING);

                jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(5).getCell(7).getStringCellValue())));
                jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(5).getCell(8).getStringCellValue())));
                jgmap.put("不合格点数",decf.format(Double.valueOf(slSheet.getRow(5).getCell(9).getStringCellValue())));
                jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(5).getCell(10).getStringCellValue())));
                mapList.add(jgmap);
                return mapList;
            }else {
                return null;
            }

        }
    }

    @Override
    public void exportsddmpzd(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "03隧道大面平整度实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcDmpzdVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"03隧道大面平整度实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }

    }

    @Override
    @Transactional
    public void importsddmpzd(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcDmpzdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcDmpzdVo>(JjgFbgcSdgcDmpzdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcDmpzdVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcSdgcDmpzdVo sdgcDmpzdVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(sdgcDmpzdVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sdgcDmpzdVo.getSdmc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，第"+rowNumber+"行的数据中，隧道名称为空，请修改后重新上传");
                                        }

                                        if (!StringUtils.isNumeric(sdgcDmpzdVo.getYxps()) || StringUtils.isEmpty(sdgcDmpzdVo.getYxps())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，第"+rowNumber+"行的数据中，允许偏差值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(sdgcDmpzdVo.getScz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，第"+rowNumber+"行的数据中，实测值有误，请修改后重新上传");
                                        }
                                        JjgFbgcSdgcDmpzd fbgcSdgcDmpzd = new JjgFbgcSdgcDmpzd();
                                        BeanUtils.copyProperties(sdgcDmpzdVo,fbgcSdgcDmpzd);
                                        fbgcSdgcDmpzd.setCreatetime(new Date());
                                        fbgcSdgcDmpzd.setUserid(commonInfoVo.getUserid());
                                        fbgcSdgcDmpzd.setProname(commonInfoVo.getProname());
                                        fbgcSdgcDmpzd.setHtd(commonInfoVo.getHtd());
                                        fbgcSdgcDmpzd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcDmpzdMapper.insert(fbgcSdgcDmpzd);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中03隧道大面平整度实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }

    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String sheetname = "二衬大面平整度";
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "40隧道大面平整度.xlsx");
        if (!f.exists()) {
            return null;
        }else {
            List<Map<String, Object>> data = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(f)) {
                Workbook workbook = WorkbookFactory.create(fis);
                Sheet sheet = workbook.getSheet(sheetname); // 假设数据在第一个工作表中

                DataFormatter dataFormatter = new DataFormatter();
                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

                for (int i = 1; i <= sheet.getLastRowNum(); i++) { // 循环每一行，从第3行开始（忽略表头）

                    Row row = sheet.getRow(i);
                    Cell cell0 = getMergedCell(sheet, i, 0); // 获取合并单元格 // 第0列
                    Cell cell13 = row.getCell(3); // 第0列
                    Cell cell8 = row.getCell(7); // 第0列
                    Cell cell9 = row.getCell(8); // 第34列


                    if (cell0.getStringCellValue().equals("隧道名称") && dataFormatter.formatCellValue(cell8, formulaEvaluator).equals("总点数") ) { // 判断是否不为空
                        Cell nextRowCell8 = sheet.getRow(i + 1).getCell(7); // 下一行的第0列
                        Cell nextRowCell9 = sheet.getRow(i + 1).getCell(8); // 下一行的第34列

                        Cell nextRowCell0 = sheet.getRow(i + 1).getCell(0); // 下一行的第34列
                        Cell nextRowCell3 = sheet.getRow(i + 1).getCell(3); // 下一行的第34列

                        String data8 = dataFormatter.formatCellValue(nextRowCell8, formulaEvaluator);
                        String data9 = dataFormatter.formatCellValue(nextRowCell9, formulaEvaluator);

                        String data0 = nextRowCell0.getStringCellValue();
                        String data3 = dataFormatter.formatCellValue(nextRowCell3, formulaEvaluator);

                        Map map = new HashMap();
                        map.put("qlmc",data0);
                        map.put("zds",data8);
                        map.put("hgds",data9);
                        map.put("sjqd",data3);

                        data.add(map);

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
        int selectnum = jjgFbgcSdgcDmpzdMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcSdgcDmpzdMapper.selectnumname(proname);
        return selectnum;
    }

    @Override
    public List<Map<String, Object>> getsdnum(String proname, String htd) {
        List<Map<String,Object>> list = jjgFbgcSdgcDmpzdMapper.getsdnum(proname,htd);
        return list;
    }

    private static Cell getMergedCell(Sheet sheet, int rowIndex, int columnIndex) {
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(rowIndex, columnIndex)) {
                Row mergedRow = sheet.getRow(range.getFirstRow());
                return mergedRow.getCell(range.getFirstColumn());
            }
        }
        return sheet.getRow(rowIndex).getCell(columnIndex);
    }

    @Override
    public int createMoreRecords(List<JjgFbgcSdgcDmpzd> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcSdgcDmpzdMapper, userID);
    }
}
