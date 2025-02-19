package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcQmhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcQmhpVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmhpMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcQmhpService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellType;
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
public class JjgFbgcQlgcQmhpServiceImpl extends ServiceImpl<JjgFbgcQlgcQmhpMapper, JjgFbgcQlgcQmhp> implements JjgFbgcQlgcQmhpService {

    @Autowired
    private JjgFbgcQlgcQmhpMapper jjgFbgcQlgcQmhpMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Autowired
    private SysUserService sysUserService;

    @Override
    public boolean generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
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
            idlist.add(Long.valueOf(userid));
        }
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmhpMapper.selectqlmc(proname,htd,fbgc,idlist);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    List<Map<String,Object>> zhlist = jjgFbgcQlgcQmhpMapper.selectzh(proname,htd,fbgc,qlmc,idlist);
                    String zh = zhlist.get(0).get("zh").toString();
                    String substring = zh.substring(0, 1);
                    if (substring.equals("Z") || substring.equals("Y")){
                        DBtoExcelZYql(proname,htd,fbgc,qlmc,idlist);
                    }else {
                        DBtoExcelql(proname,htd,fbgc,qlmc,idlist);
                    }

                }
            }
            return true;
        }else {
            return false;
        }
    }

    /**
     * 桩号是不分左右幅的
     * @param proname
     * @param htd
     * @param fbgc
     * @param qlmc
     * @param idlist
     */
    private void DBtoExcelql(String proname, String htd, String fbgc, String qlmc, List<Long> idlist) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcQlgcQmhp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("qlmc",qlmc);
        wrapper.in("userid",idlist);
        wrapper.orderByAsc("zh");
        List<JjgFbgcQlgcQmhp> data = jjgFbgcQlgcQmhpMapper.selectList(wrapper);
        String lmlx = data.get(0).getLmlx();
        String lx = lmlx.substring(0, 2);
        String sheetname;
        if (lx.equals("水泥")){
            sheetname="混凝土桥面左幅";
        }else {
            sheetname="沥青桥面左幅";

        }
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"35桥面横坡-"+qlmc+".xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (data == null || data.size()==0){
            return;
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
                String name = "横坡.xlsx";
                String path = reportPath + File.separator + name;
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/横坡.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                createTable(gettableNumAll(data.size()),wb,sheetname);
                if(DBtoExcel(data,wb,sheetname,qlmc)){
                    String[] arr = {"混凝土匝道","沥青匝道","沥青路面左幅","沥青路面右幅","沥青桥面左幅","沥青桥面右幅","沥青隧道左幅","沥青隧道右幅","混凝土路面左幅","混凝土路面右幅","混凝土桥面左幅","混凝土桥面右幅","混凝土隧道左幅","混凝土隧道右幅"};
                    for (int i = 0; i < arr.length; i++) {
                        if (shouldBeCalculate(wb.getSheet(arr[i]))) {
                            calculateSheet(wb,wb.getSheet(arr[i]).getSheetName());
                            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet(arr[i]));
                        } else {
                            wb.removeSheetAt(wb.getSheetIndex(arr[i]));
                        }
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
     * @return
     */
    private boolean shouldBeCalculate(XSSFSheet sheet) {
        sheet.getRow(6).getCell(0).setCellType(CellType.STRING);
        if(sheet.getRow(6).getCell(0) == null || "".equals(sheet.getRow(6).getCell(0).getStringCellValue())){
            return false;
        }
        return true;
    }

    /**
     *
     * @param wb
     * @param sheetname
     */
    private void calculateSheet(XSSFWorkbook wb, String sheetname) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFRow row = null;
        boolean flag = false;
        String name = "";
        ArrayList<XSSFRow> startRowList = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> endRowList = new ArrayList<XSSFRow>();
        sheet.getRow(0).getCell(12).setCellValue("合计");
        sheet.getRow(0).getCell(13).setCellValue("总点数");
        sheet.getRow(0).getCell(14).setCellFormula("SUM("
                +sheet.getRow(5).getCell(14).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(14).getReference()+")");
        sheet.getRow(0).getCell(15).setCellValue("合格点数");
        sheet.getRow(0).getCell(16).setCellFormula("SUM("
                +sheet.getRow(5).getCell(16).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(16).getReference()+")");
        sheet.getRow(0).getCell(17).setCellValue("合格率");
        sheet.getRow(0).getCell(18).setCellFormula(
                sheet.getRow(0).getCell(16).getReference()+"/"
                        +sheet.getRow(0).getCell(14).getReference()+"*100");
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (flag && !"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表")) {
                flag = false;
            }
            if(flag && !"".equals(row.getCell(2).toString())){
                row.getCell(5).setCellFormula(row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference());//F=E7-D7
                row.getCell(7).setCellFormula(row.getCell(5).getReference()+"/"
                        +row.getCell(6).getReference()+"*100");//H=F7/G7*100
                row.getCell(10).setCellFormula("IF((("
                        +row.getCell(7).getReference()+"+"
                        +row.getCell(9).getReference()+")>="
                        +row.getCell(8).getReference()+")*AND("
                        +row.getCell(8).getReference()+">=("
                        +row.getCell(7).getReference()+"-"
                        +row.getCell(9).getReference()+")),\"√\",\"\")");//K=IF(((H7+0.3)>=I7)*AND(I7>=(H7-0.3)),"√","")
                row.getCell(11).setCellFormula("IF((("
                        +row.getCell(7).getReference()+"+"
                        +row.getCell(9).getReference()+")>="
                        +row.getCell(8).getReference()+")*AND("
                        +row.getCell(8).getReference()+">=("
                        +row.getCell(7).getReference()+"-"
                        +row.getCell(9).getReference()+")),\"\",\"×\")");//L=IF(((H7+0.3)>=I7)*AND(I7>=(H7-0.3)),"","×")
                row.getCell(12).setCellFormula(row.getCell(7).getReference()+"-"
                        +row.getCell(8).getReference());
            }
            if ("桩  号".equals(row.getCell(0).toString())) {
                i += 1;
                flag = true;
                if(!name.equals(sheet.getRow(i-3).getCell(8).toString()) && !"".equals(name)){
                    String column14 = "";
                    String column16 = "";
                    for(int r = 0;r < startRowList.size(); r++){
                        column14 += startRowList.get(r).getCell(7).getReference()+":"+endRowList.get(r).getCell(7).getReference()+",";
                        column16 += "COUNTIF("+startRowList.get(r).getCell(10).getReference()+":"+endRowList.get(r).getCell(10).getReference()+",\"√\")+";
                    }
                    startRowList.get(0).createCell(14).setCellFormula("COUNT("+column14.substring(0, column14.lastIndexOf(','))+")");
                    startRowList.get(0).createCell(16).setCellFormula(column16.substring(0, column16.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
                    startRowList.get(0).createCell(18).setCellFormula(startRowList.get(0).getCell(16).getReference()+"*100/"
                            +startRowList.get(0).getCell(14).getReference());//=COUNTIF(AH6:AH50,"×")

                    startRowList.clear();
                    endRowList.clear();
                }
                startRowList.add(sheet.getRow(i+1));
                endRowList.add(sheet.getRow(i+29));
                name = sheet.getRow(i-3).getCell(8).toString();
            }
        }
        String column14 = "";
        String column16 = "";
        for(int r = 0;r < startRowList.size(); r++){
            column14 += startRowList.get(r).getCell(7).getReference()+":"+endRowList.get(r).getCell(7).getReference()+",";
            column16 += "COUNTIF("+startRowList.get(r).getCell(10).getReference()+":"+endRowList.get(r).getCell(10).getReference()+",\"√\")+";
        }
        startRowList.get(0).createCell(14).setCellFormula("COUNT("+column14.substring(0, column14.lastIndexOf(','))+")");
        startRowList.get(0).createCell(16).setCellFormula(column16.substring(0, column16.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
        startRowList.get(0).createCell(18).setCellFormula(startRowList.get(0).getCell(16).getReference()+"*100/"
                +startRowList.get(0).getCell(14).getReference());//=COUNTIF(AH6:AH50,"×")

        startRowList.clear();
        endRowList.clear();
    }

    /**
     *
     * @param data
     * @param wb
     * @param sheetname
     * @param qlmc
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcQlgcQmhp> data, XSSFWorkbook wb,String sheetname,String qlmc) throws ParseException {
        if (data.size()>0){
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
            String proname = data.get(0).getProname();
            String htd = data.get(0).getHtd();
            String fbgc = data.get(0).getFbgc();
            XSSFSheet sheet = wb.getSheet(sheetname);
            Date jcsj = data.get(0).getJcsj();
            String testtime = simpleDateFormat.format(jcsj);
            int index = 0;
            int tableNum = 0;
            fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).getLmlx(),qlmc);
            for(JjgFbgcQlgcQmhp row:data){
                //比较检测时间，拿到最新的检测时间
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(index/29 == 1){
                    sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.getLmlx(),qlmc);
                    index = 0;
                }
                //填写中间下方的普通单元格
                fillCommonCellData(sheet, tableNum, index+6, row,cellstyle);
                index++;
            }
            sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
            return true;
        }else {
            return false;
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
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcQlgcQmhp row, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*35+index).getCell(0).setCellValue(row.getZh());
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(row.getWz());
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.parseDouble(row.getQsds()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.parseDouble(row.getHsds()));
        sheet.getRow(tableNum*35+index).getCell(6).setCellValue(Double.parseDouble(row.getLength()));
        sheet.getRow(tableNum*35+index).getCell(8).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(tableNum*35+index).getCell(9).setCellValue(Double.parseDouble(row.getYxps()));
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param fbgc
     * @param lmlx
     * @param qlmc
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String fbgc,String lmlx,String qlmc) {
        sheet.getRow(tableNum*35+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*35+1).getCell(8).setCellValue(htd);
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue("桥梁工程");
        sheet.getRow(tableNum*35+2).getCell(8).setCellValue(fbgc+"("+qlmc+")");
        sheet.getRow(tableNum*35+3).getCell(2).setCellValue(lmlx);

    }

    /**
     *
     * @param tableNum
     * @param wb
     * @param sheetname
     */
    private void createTable(int tableNum, XSSFWorkbook wb,String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 34, i* 35);
        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0,(record) * 35-1);
        }
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNumAll(int size) {
        return size%29 ==0 ? size/29 : size/29+1;
    }

    /**
     * 桩号是分左右幅的
     * @param proname
     * @param htd
     * @param fbgc
     * @param qlmc
     * @param idlist
     * @throws IOException
     */
    private void DBtoExcelZYql(String proname, String htd, String fbgc, String qlmc, List<Long> idlist) throws IOException, ParseException {
        XSSFWorkbook wb = null;

        QueryWrapper<JjgFbgcQlgcQmhp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("qlmc",qlmc);
        wrapper.like("wz","左幅");
        wrapper.in("userid",idlist);
        wrapper.orderByAsc("zh");
        List<JjgFbgcQlgcQmhp> zdata = jjgFbgcQlgcQmhpMapper.selectList(wrapper);

        QueryWrapper<JjgFbgcQlgcQmhp> wrapper2=new QueryWrapper<>();
        wrapper2.like("proname",proname);
        wrapper2.like("htd",htd);
        wrapper2.like("fbgc",fbgc);
        wrapper2.like("qlmc",qlmc);
        wrapper2.like("wz","右幅");
        wrapper2.in("userid",idlist);

        wrapper2.orderByAsc("zh");
        List<JjgFbgcQlgcQmhp> ydata = jjgFbgcQlgcQmhpMapper.selectList(wrapper2);
        String lmlx = zdata.get(0).getLmlx();
        String lx = lmlx.substring(0, 2);
        String shname;
        if (lx.equals("水泥")){
            shname="混凝土桥面";
        }else {
            shname="沥青桥面";

        }
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"35桥面横坡-"+qlmc+".xlsx");
        //存放鉴定表的目录
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        try {
           /* File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "横坡.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));*/
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/横坡.xlsx");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNumAll(zdata.size()),wb,shname+"左幅");
            createTable(gettableNumAll(ydata.size()),wb,shname+"右幅");
            DBtoExcel(zdata,wb,shname+"左幅",qlmc);
            DBtoExcel(ydata,wb,shname+"右幅",qlmc);
            calculateSheet(wb,shname+"左幅");
            calculateSheet(wb,shname+"右幅");

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }
            JjgFbgcCommonUtils.deleteEmptySheets(wb);
            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
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


    /**
     * 页面上调用的
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String userid = commonInfoVo.getUserid();
        String htd = commonInfoVo.getHtd();
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
            idlist.add(Long.valueOf(userid));
        }
        // fbgc = commonInfoVo.getFbgc();
        List<Map<String, Object>> selectqlmc = selectqlmc2(proname, htd);

        List<Map<String, Object>> mapList = new ArrayList<>();

        for (int i=0;i<selectqlmc.size();i++){
            String qlmc = selectqlmc.get(i).get("qlmc").toString();
            File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "35桥面横坡-"+qlmc+".xlsx");
            if (!f.exists()) {
                continue;
            } else {
                List<Map<String,Object>> zhlist = jjgFbgcQlgcQmhpMapper.selectzh2(proname,htd,qlmc,idlist);
                String zh = zhlist.get(0).get("zh").toString();
                String lx = zhlist.get(0).get("lmlx").toString().substring(0,2);
                String lxname;
                if (lx.equals("水泥")){
                    lxname="混凝土桥面";
                }else {
                    lxname="沥青桥面";

                }
                String substring = zh.substring(0, 1);
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));

                if (substring.equals("Z") || substring.equals("Y")){
                    String sheetnameleft = lxname+"左幅";
                    String sheetnameright = lxname+"右幅";
                    XSSFSheet sheetleft = xwb.getSheet(sheetnameleft);
                    XSSFSheet sheetright = xwb.getSheet(sheetnameright);
                    XSSFCell xmnameleft = sheetleft.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetleft.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        if(sheetleft != null){
                            sheetleft.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                            sheetleft.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                            sheetleft.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                            sheetleft.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率

                            Map<String, Object> jgmap = new HashMap<>();
                            jgmap.put("检测项目",qlmc);
                            jgmap.put("zyf",sheetnameleft);
                            jgmap.put("yxpc",sheetleft.getRow(6).getCell(9).getStringCellValue());
                            jgmap.put("总点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(14).getStringCellValue())));
                            jgmap.put("合格点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(16).getStringCellValue())));
                            jgmap.put("合格率", df.format(Double.valueOf(sheetleft.getRow(0).getCell(18).getStringCellValue())));
                            mapList.add(jgmap);
                        }

                        if(sheetright != null){
                            sheetright.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                            sheetright.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                            sheetright.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                            sheetright.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率

                            Map<String, Object> jgmapright = new HashMap<>();
                            jgmapright.put("检测项目",qlmc);
                            jgmapright.put("zyf",sheetnameright);
                            jgmapright.put("yxpc",sheetright.getRow(6).getCell(9).getStringCellValue());
                            jgmapright.put("总点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(14).getStringCellValue())));
                            jgmapright.put("合格点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(16).getStringCellValue())));
                            jgmapright.put("合格率", df.format(Double.valueOf(sheetright.getRow(0).getCell(18).getStringCellValue())));
                            mapList.add(jgmapright);
                        }

                    }
                }else {
                    String sheetname = lxname+"左幅";
                    XSSFSheet sheetl = xwb.getSheet(sheetname);
                    XSSFCell xmnameleft = sheetl.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetl.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        sheetl.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetl.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetl.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        sheetl.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率
                        Map<String, Object> jgmap = new HashMap<>();
                        jgmap.put("检测项目",qlmc);
                        jgmap.put("zyf",lxname);
                        jgmap.put("yxpc",sheetl.getRow(6).getCell(9).getStringCellValue());
                        jgmap.put("总点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(14).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(16).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheetl.getRow(0).getCell(18).getStringCellValue())));
                        mapList.add(jgmap);
                    }
                }
            }

        }
        return mapList;
    }

    /**
     *
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    public List<Map<String, Object>> lookJdbjgdpksh(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String, Object>> selectqlmc = selectqlmc2(proname, htd);

        List<Map<String, Object>> mapList = new ArrayList<>();

        for (int i=0;i<selectqlmc.size();i++){
            String qlmc = selectqlmc.get(i).get("qlmc").toString();
            File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "35桥面横坡-"+qlmc+".xlsx");
            if (!f.exists()) {
                continue;
            } else {
                List<Map<String,Object>> zhlist = jjgFbgcQlgcQmhpMapper.selectzh2noid(proname,htd,qlmc);
                String zh = zhlist.get(0).get("zh").toString();
                String lx = zhlist.get(0).get("lmlx").toString().substring(0,2);
                String lxname;
                if (lx.equals("水泥")){
                    lxname="混凝土桥面";
                }else {
                    lxname="沥青桥面";

                }
                String substring = zh.substring(0, 1);
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));

                if (substring.equals("Z") || substring.equals("Y")){
                    String sheetnameleft = lxname+"左幅";
                    String sheetnameright = lxname+"右幅";
                    XSSFSheet sheetleft = xwb.getSheet(sheetnameleft);
                    XSSFSheet sheetright = xwb.getSheet(sheetnameright);
                    XSSFCell xmnameleft = sheetleft.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetleft.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        sheetleft.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetleft.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetleft.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        sheetleft.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率

                        sheetright.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetright.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetright.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        sheetright.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率

                        Map<String, Object> jgmap = new HashMap<>();
                        jgmap.put("检测项目",qlmc);
                        jgmap.put("zyf",sheetnameleft);
                        jgmap.put("yxpc",sheetleft.getRow(6).getCell(9).getStringCellValue());
                        jgmap.put("总点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(14).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(16).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheetleft.getRow(0).getCell(18).getStringCellValue())));

                        Map<String, Object> jgmapright = new HashMap<>();
                        jgmapright.put("检测项目",qlmc);
                        jgmapright.put("zyf",sheetnameright);
                        jgmapright.put("yxpc",sheetright.getRow(6).getCell(9).getStringCellValue());
                        jgmapright.put("总点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(14).getStringCellValue())));
                        jgmapright.put("合格点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(16).getStringCellValue())));
                        jgmapright.put("合格率", df.format(Double.valueOf(sheetright.getRow(0).getCell(18).getStringCellValue())));
                        mapList.add(jgmap);
                        mapList.add(jgmapright);
                    }
                }else {
                    String sheetname = lxname+"左幅";
                    XSSFSheet sheetl = xwb.getSheet(sheetname);
                    XSSFCell xmnameleft = sheetl.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetl.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        sheetl.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetl.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetl.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        sheetl.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率
                        Map<String, Object> jgmap = new HashMap<>();
                        jgmap.put("检测项目",qlmc);
                        jgmap.put("zyf",lxname);
                        jgmap.put("yxpc",sheetl.getRow(6).getCell(9).getStringCellValue());
                        jgmap.put("总点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(14).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(16).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheetl.getRow(0).getCell(18).getStringCellValue())));
                        mapList.add(jgmap);
                    }
                }
            }

        }
        return mapList;
    }

    @Override
    public List<Map<String, Object>> getqlname(String proname, String htd) {
        List<Map<String, Object>> list = jjgFbgcQlgcQmhpMapper.getqlName(proname,htd);
        return list;
    }

    @Override
    public void export(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "桥面横坡实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcQmhpVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"桥面横坡实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }
    }

    @Override
    @Transactional
    public void importqmhp(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcQmhpVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcQmhpVo>(JjgFbgcQlgcQmhpVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcQmhpVo> dataList) {
                                    int rowNumber=2;
                                    for(JjgFbgcQlgcQmhpVo fbgcQlgcQmhpVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(fbgcQlgcQmhpVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(fbgcQlgcQmhpVo.getQlmc())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，桥梁名称为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(fbgcQlgcQmhpVo.getLmlx())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，路面类型为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(fbgcQlgcQmhpVo.getWz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，位置为空，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(fbgcQlgcQmhpVo.getQsds()) || StringUtils.isEmpty(fbgcQlgcQmhpVo.getQsds())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，前视读数值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(fbgcQlgcQmhpVo.getHsds()) || StringUtils.isEmpty(fbgcQlgcQmhpVo.getHsds())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，后视读数值有误，请修改后重新上传");
                                        }
                                        if (!StringUtils.isNumeric(fbgcQlgcQmhpVo.getLength()) || StringUtils.isEmpty(fbgcQlgcQmhpVo.getLength())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，长值有误，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(fbgcQlgcQmhpVo.getSjz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，设计值为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(fbgcQlgcQmhpVo.getYxps()) || fbgcQlgcQmhpVo.getYxps().contains("-") || fbgcQlgcQmhpVo.getYxps().contains("+") || fbgcQlgcQmhpVo.getYxps().contains("±") ) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，第"+rowNumber+"行的数据中，允许偏差值有误，请修改后重新上传");
                                        }
                                        JjgFbgcQlgcQmhp fbgcQlgcQmhp = new JjgFbgcQlgcQmhp();
                                        BeanUtils.copyProperties(fbgcQlgcQmhpVo,fbgcQlgcQmhp);
                                        fbgcQlgcQmhp.setCreatetime(new Date());
                                        fbgcQlgcQmhp.setProname(commonInfoVo.getProname());
                                        fbgcQlgcQmhp.setUserid(commonInfoVo.getUserid());
                                        fbgcQlgcQmhp.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcQmhp.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcQmhpMapper.insert(fbgcQlgcQmhp);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中桥面横坡实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }

    }

    @Override
    public List<Map<String, Object>> selectqlmc(String proname, String htd, String fbgc,String userid) {
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
            idlist.add(Long.valueOf(userid));
        }
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmhpMapper.selectqlmc(proname,htd,fbgc,idlist);
        return qlmclist;
    }

    @Override
    public List<Map<String, Object>> lookJdb(CommonInfoVo commonInfoVo, String value) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //List<Map<String, Object>> selectqlmc = selectqlmc2(proname, htd);

        List<Map<String, Object>> mapList = new ArrayList<>();

        //for (int i=0;i<selectqlmc.size();i++){
            String qlmc = StringUtils.substringBetween(value, "-", ".");
            File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + value);
            if (!f.exists()) {
                return new ArrayList<>();
            } else {
                List<Map<String,Object>> zhlist = jjgFbgcQlgcQmhpMapper.selectzh2noid(proname,htd,qlmc);

                // 进行判空操作
                if(zhlist.isEmpty()) return new ArrayList<>();
                String zh = zhlist.get(0).get("zh").toString();
                String lx = zhlist.get(0).get("lmlx").toString().substring(0,2);
                String lxname;
                if (lx.equals("水泥")){
                    lxname="混凝土桥面";
                }else {
                    lxname="沥青桥面";

                }
                String substring = zh.substring(0, 1);
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));

                if (substring.equals("Z") || substring.equals("Y")){
                    String sheetnameleft = lxname+"左幅";
                    String sheetnameright = lxname+"右幅";
                    XSSFSheet sheetleft = xwb.getSheet(sheetnameleft);
                    XSSFSheet sheetright = xwb.getSheet(sheetnameright);
                    XSSFCell xmnameleft = sheetleft.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetleft.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        if(sheetleft!=null){
                            sheetleft.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                            sheetleft.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                            sheetleft.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                            sheetleft.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率

                            Map<String, Object> jgmap = new HashMap<>();
                            jgmap.put("检测项目",qlmc);
                            jgmap.put("zyf",sheetnameleft);
                            jgmap.put("yxpc",sheetleft.getRow(6).getCell(9).getStringCellValue());
                            jgmap.put("总点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(14).getStringCellValue())));
                            jgmap.put("合格点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(16).getStringCellValue())));
                            jgmap.put("合格率", df.format(Double.valueOf(sheetleft.getRow(0).getCell(18).getStringCellValue())));
                            mapList.add(jgmap);
                        }
                        if(sheetright!=null){
                            sheetright.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                            sheetright.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                            sheetright.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                            sheetright.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率

                            Map<String, Object> jgmapright = new HashMap<>();
                            jgmapright.put("检测项目",qlmc);
                            jgmapright.put("zyf",sheetnameright);
                            jgmapright.put("yxpc",sheetright.getRow(6).getCell(9).getStringCellValue());
                            jgmapright.put("总点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(14).getStringCellValue())));
                            jgmapright.put("合格点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(16).getStringCellValue())));
                            jgmapright.put("合格率", df.format(Double.valueOf(sheetright.getRow(0).getCell(18).getStringCellValue())));
                            mapList.add(jgmapright);
                        }


                    }
                }else {
                    String sheetname = lxname+"左幅";
                    XSSFSheet sheetl = xwb.getSheet(sheetname);
                    XSSFCell xmnameleft = sheetl.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetl.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        sheetl.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetl.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetl.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        sheetl.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率
                        Map<String, Object> jgmap = new HashMap<>();
                        jgmap.put("检测项目",qlmc);
                        jgmap.put("zyf",lxname);
                        jgmap.put("yxpc",sheetl.getRow(6).getCell(9).getStringCellValue());
                        jgmap.put("总点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(14).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(16).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheetl.getRow(0).getCell(18).getStringCellValue())));
                        mapList.add(jgmap);
                    }
                }
           // }

        }
        return mapList;

    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcQlgcQmhpMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcQlgcQmhpMapper.selectnumname(proname);
        return selectnum;
    }

    public List<Map<String, Object>> selectqlmc2(String proname, String htd) {
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmhpMapper.selectqlmc2(proname,htd);
        return qlmclist;
    }

    @Override
    public int createMoreRecords(List<JjgFbgcQlgcQmhp> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcQlgcQmhpMapper, userID);
    }
}
