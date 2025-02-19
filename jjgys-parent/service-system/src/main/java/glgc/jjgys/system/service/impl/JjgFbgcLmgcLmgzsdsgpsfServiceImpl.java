package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLmgzsdsgpsf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLmgzsdsgpsfVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLmgzsdsgpsfMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcLmgzsdsgpsfService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
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
 * @since 2023-05-03
 */
@Service
public class JjgFbgcLmgcLmgzsdsgpsfServiceImpl extends ServiceImpl<JjgFbgcLmgcLmgzsdsgpsfMapper, JjgFbgcLmgcLmgzsdsgpsf> implements JjgFbgcLmgcLmgzsdsgpsfService {

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfMapper jjgFbgcLmgcLmgzsdsgpsfMapper;

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
        QueryWrapper<JjgFbgcLmgcLmgzsdsgpsf> wrapper=new QueryWrapper<>();
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
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcLmgzsdsgpsf> data = jjgFbgcLmgcLmgzsdsgpsfMapper.selectList(wrapper);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"20构造深度手工铺沙法.xlsx");
        if (data == null || data.size()==0){
            return false;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/构造深度手工铺沙法.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                createTable(gettableNum(data.size()),wb);
                if(DBtoExcel(data,wb)){
                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }

                    JjgFbgcCommonUtils.deleteEmptySheets(wb);
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
     * @param data
     * @param wb
     * @return
     */
    private boolean DBtoExcel(List<JjgFbgcLmgcLmgzsdsgpsf> data, XSSFWorkbook wb) {
        if (data.size()>0){
            XSSFSheet sheet = wb.getSheet("混凝土路面");
            int tableNO = 0;
            //首先填写表头
            fillCellTitle(data.get(0), tableNO, sheet);
            //填写数据
            int record = 0;
            String name = data.get(0).getLmlx();
            for(JjgFbgcLmgcLmgzsdsgpsf row : data){
                if(name.equals(row.getLmlx()) && record <= 26-5){
                    fillCellBody(row, tableNO, record, sheet);
                    record ++;
                }
                else{
                    tableNO ++;
                    fillCellTitle(row, tableNO, sheet);
                    name = row.getLmlx();
                    record = 0;
                    fillCellBody(row, tableNO, record, sheet);
                    record ++;
                }
            }

            XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(wb);
            XSSFCellStyle cellstyle_copy = wb.createCellStyle();
            XSSFFont font=wb.createFont();
            font.setFontHeightInPoints((short)10);
            font.setFontName("宋体");
            cellstyle_copy.setFont(font);
            cellstyle_copy.setBorderBottom(BorderStyle.THIN); //下边框
            cellstyle_copy.setBorderLeft(BorderStyle.THIN);//左边框
            cellstyle_copy.setBorderTop(BorderStyle.THIN);//上边框
            cellstyle_copy.setBorderRight(BorderStyle.THIN);//右边框
            cellstyle_copy.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
            cellstyle_copy.setAlignment(HorizontalAlignment.CENTER);//水平

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
            XSSFCellStyle cellstyle = (XSSFCellStyle) cellstyle_copy.clone();
            cellstyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,#0"));
            row.getCell(4).setCellStyle(cellstyle);
            row.getCell(4).setCellValue(Double.valueOf(value).intValue());

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 5, 6));
            row.getCell(5).setCellValue("合格率");
            row.getCell(7).setCellFormula(row.getCell(4).getReference()+"*100/"+row.getCell(1).getReference());
            value = e.evaluate(row.getCell(7)).getNumberValue();
            row.getCell(7).setCellFormula(null);
            cellstyle = (XSSFCellStyle) cellstyle_copy.clone();
            cellstyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.0"));
            row.getCell(7).setCellStyle(cellstyle);
            DecimalFormat df = new DecimalFormat("#0.0");
            row.getCell(7).setCellValue(df.format(value));

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 8, 9));
            row.getCell(8).setCellValue("最大值");
            row.getCell(10).setCellFormula("MAX(L7:L"+curRow+")");
            value = e.evaluate(row.getCell(10)).getNumberValue();
            row.getCell(10).setCellFormula(null);
            cellstyle = (XSSFCellStyle) cellstyle_copy.clone();
            cellstyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
            row.getCell(10).setCellStyle(cellstyle);
            row.getCell(10).setCellValue(value);

            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 11, 12));
            row.getCell(11).setCellValue("最小值");
            row.getCell(13).setCellFormula("MIN(L7:L"+curRow+")");
            value = e.evaluate(row.getCell(13)).getNumberValue();
            row.getCell(13).setCellFormula(null);
            cellstyle = (XSSFCellStyle) cellstyle_copy.clone();
            cellstyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
            row.getCell(13).setCellStyle(cellstyle);
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
    private void fillCellBody(JjgFbgcLmgcLmgzsdsgpsf row, int tableNO, int record, XSSFSheet sheet) {
        XSSFRow currow = sheet.getRow(tableNO * (33-5) + 6 + record);
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
        currow.getCell(12).setCellFormula("IF("+currow.getCell(11).getReference()+">="+row.getSjzxz()+",\"√\",\"\")");

        currow.getCell(13).setCellFormula("IF("+currow.getCell(12).getReference()+"=\"\",\"×\",\"\")");
    }

    /**
     *
     * @param row
     * @param tableNO
     * @param sheet
     */
    private void fillCellTitle(JjgFbgcLmgcLmgzsdsgpsf row, int tableNO, XSSFSheet sheet) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFRow titlerow = sheet.getRow(tableNO * (33-5) + 1);
        titlerow.getCell(1).setCellValue(row.getProname());
        titlerow.getCell(9).setCellValue(row.getHtd());
        titlerow = sheet.getRow(tableNO * (33-5) + 2);
        titlerow.getCell(1).setCellValue("路面面层("+row.getLmlx()+")");
        titlerow.getCell(9).setCellValue("水泥混凝土路面");
        titlerow = sheet.getRow(tableNO * (33-5) + 3);
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
            RowCopy.copyRows(wb, "混凝土路面", "混凝土路面", 0, 32-5, i * (33-5));
        }

        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex("混凝土路面"), 0, 13, 0, record * (33-5)-1);

    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%22 <= 21 ? size/22+1 : size/22+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "混凝土路面构造深度质量鉴定表";
        String sheetname = "混凝土路面";

        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "20构造深度手工铺沙法.xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            //XSSFSheet slSheet = wb.getSheet(sheetname);

            List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
                    XSSFCell xmname = slSheet.getRow(1).getCell(1);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名

                    if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(1).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(3).getCell(1).setCellType(CellType.STRING);//合格率

                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum).getCell(13).setCellType(CellType.STRING);

                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(1).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        jgmap.put("检测点数", zdsz1);
                        jgmap.put("路面类型", wb.getSheetName(j));
                        jgmap.put("规定值", slSheet.getRow(3).getCell(1).getStringCellValue());
                        jgmap.put("合格点数", hgdsz1);
                        jgmap.put("合格率", hglz1);
                        jgmap.put("Min", slSheet.getRow(lastRowNum).getCell(13).getStringCellValue());
                        jgmap.put("Max", slSheet.getRow(lastRowNum).getCell(10).getStringCellValue());
                        mapList.add(jgmap);
                    }
                }
            }
            return mapList;

        }

    }

    @Override
    public void exportlmgzsdsgpsf(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "08路面构造深度手工铺砂法实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLmgzsdsgpsfVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"08路面构造深度手工铺砂法实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }
    }

    @Override
    @Transactional
    public void importlmgzsdsgpsf(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLmgzsdsgpsfVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLmgzsdsgpsfVo>(JjgFbgcLmgcLmgzsdsgpsfVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLmgzsdsgpsfVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcLmgcLmgzsdsgpsfVo lmgzsdsgpsfVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getLmlx())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，路面类型为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getAbm())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，ABM为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getCd())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，车道值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getSjzxz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，设计最小值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getSjzdz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，设计最大值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getCd1d1())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点1D1值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getCd1d2())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点1D2值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getCd2d1())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点2D1值有误，请修改后重新上传");
                                        }
                                        if ( StringUtils.isEmpty(lmgzsdsgpsfVo.getCd2d2())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点2D2值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lmgzsdsgpsfVo.getCd3d1())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点3D1值有误，请修改后重新上传");
                                        }
                                        if ( StringUtils.isEmpty(lmgzsdsgpsfVo.getCd3d2())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，第"+rowNumber+"行的数据中，测点3D2值有误，请修改后重新上传");
                                        }
                                        JjgFbgcLmgcLmgzsdsgpsf fbgcLmgcLmgzsdsgpsf = new JjgFbgcLmgcLmgzsdsgpsf();
                                        BeanUtils.copyProperties(lmgzsdsgpsfVo,fbgcLmgcLmgzsdsgpsf);
                                        fbgcLmgcLmgzsdsgpsf.setCreatetime(new Date());
                                        fbgcLmgcLmgzsdsgpsf.setUserid(commonInfoVo.getUserid());
                                        fbgcLmgcLmgzsdsgpsf.setProname(commonInfoVo.getProname());
                                        fbgcLmgcLmgzsdsgpsf.setHtd(commonInfoVo.getHtd());
                                        fbgcLmgcLmgzsdsgpsf.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLmgzsdsgpsfMapper.insert(fbgcLmgcLmgzsdsgpsf);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中08路面构造深度手工铺砂法实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }

    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLmgcLmgzsdsgpsfMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcLmgcLmgzsdsgpsfMapper.selectnumname(proname);
        return selectnum;
    }

    @Override
    public int createMoreRecords(List<JjgFbgcLmgcLmgzsdsgpsf> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcLmgcLmgzsdsgpsfMapper, userID);
    }
}
