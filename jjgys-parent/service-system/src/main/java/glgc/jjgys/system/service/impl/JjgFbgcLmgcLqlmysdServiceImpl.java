package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLqlmysdVo;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLqlmysdMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcLqlmysdService;
import glgc.jjgys.system.service.SysUserService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.ReceiveUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
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
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-04-10
 */
@Service
public class JjgFbgcLmgcLqlmysdServiceImpl extends ServiceImpl<JjgFbgcLmgcLqlmysdMapper, JjgFbgcLmgcLqlmysd> implements JjgFbgcLmgcLqlmysdService {

    @Autowired
    private JjgFbgcLmgcLqlmysdMapper jjgFbgcLmgcLqlmysdMapper;

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
        QueryWrapper<JjgFbgcLmgcLqlmysd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        String userid = commonInfoVo.getUserid();

        //查询用户类型
        QueryWrapper<SysUser> wrapperuser = new QueryWrapper<>();
        wrapperuser.eq("id",userid);
        SysUser one = sysUserService.getOne(wrapperuser);
        String type = one.getType();
        List<Long> idlist = new ArrayList<>();
        if ("2".equals(type) || "4".equals(type)){
            //公司管理员
            Long deptId = one.getDeptId();
            QueryWrapper<SysUser> wrapperid = new QueryWrapper<>();
            wrapperid.eq("dept_id",deptId);
            List<SysUser> list = sysUserService.list(wrapperid);
            //拿到部门下所有用户的id
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
            idlist.add(Long.valueOf(userid));
        }
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcLqlmysd> data = jjgFbgcLmgcLqlmysdMapper.selectList(wrapper);
        List<JjgFbgcLmgcLqlmysd> fldata = new ArrayList<>();
        List<JjgFbgcLmgcLqlmysd> nofldata = new ArrayList<>();
        if (data != null){
            for (JjgFbgcLmgcLqlmysd datum : data) {
                String sffl = datum.getSffl();
                if ("是".equals(sffl)){
                    fldata.add(datum);
                }else {
                    nofldata.add(datum);
                }
            }
        }

        //String sffl = data.get(0).getSffl();

        if (data == null || data.size()==0){
            return false;
        }else {

           /* File directory = new File("service-system/src/main/resources");
            String reportPath = directory.getCanonicalPath();*/
            //if (sffl.equals("否")) {
            if (nofldata != null && nofldata.size()>0) {
                //鉴定表要存放的路径
                File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"12沥青路面压实度-(不分离上面层).xlsx");
                File fdir = new File(filepath + File.separator + proname + File.separator + htd);
                if (!fdir.exists()) {
                    //创建文件根目录
                    fdir.mkdirs();
                }

                try {
                    /*String path = reportPath + "/static/沥青路面压实度.xlsx";
                    Files.copy(Paths.get(path), new FileOutputStream(f));*/
                    InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/沥青路面压实度.xlsx");
                    Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    FileInputStream out = new FileInputStream(f);
                    wb = new XSSFWorkbook(out);

                    List<JjgFbgcLmgcLqlmysd> zxzfdata = new ArrayList<>();
                    List<JjgFbgcLmgcLqlmysd> zxyfdata = new ArrayList<>();
                    List<JjgFbgcLmgcLqlmysd> sdzfdata = new ArrayList<>();
                    List<JjgFbgcLmgcLqlmysd> sdyfdata = new ArrayList<>();
                    for(JjgFbgcLmgcLqlmysd ysd : nofldata){
                        //沥青路面压实度左幅  将所有左幅的数据收集，包含路，桥，隧道，不包含匝道，连接线，连接线隧道，并且也不分左右幅
                        String substring = ysd.getZh().substring(0, 1);
                        if (substring.equals("Z") || substring.equals("Y")){
                            if (ysd.getZh().substring(0,1).equals("Z") && ysd.getLqs().equals("路") || ysd.getLqs().contains("桥")){
                                zxzfdata.add(ysd);
                            }
                            if (ysd.getZh().substring(0,1).equals("Y") && ysd.getLqs().equals("路") || ysd.getLqs().contains("桥") ){
                                zxyfdata.add(ysd);
                            }
                            //隧道和桥的数据 是放在隧道左右幅的工作簿中，分左右幅
                            if (ysd.getLqs().contains("隧") && ysd.getZh().substring(0,1).equals("Z")){
                                sdzfdata.add(ysd);
                            }
                            if (ysd.getLqs().contains("隧") && ysd.getZh().substring(0,1).equals("Y")){
                                sdyfdata.add(ysd);
                            }
                        }else {
                            if (ysd.getLqs().equals("路") || ysd.getLqs().contains("桥")){
                                zxzfdata.add(ysd);
                            }
                            //隧道和桥的数据 是放在隧道左右幅的工作簿中，分左右幅
                            if (ysd.getLqs().contains("隧")){
                                sdzfdata.add(ysd);
                            }
                        }

                    }

                    // 不理解，这里的数据又是搜索的，直接全部都用sql搜索就行了。现在对下面的搜索限制，需要搜索不可分离的数据。下面的搜索只在这里用到，放心改
                    List<JjgFbgcLmgcLqlmysd> zddata = jjgFbgcLmgcLqlmysdMapper.selectzd(proname,htd,fbgc,idlist);
                    List<JjgFbgcLmgcLqlmysd> zdyfdata = jjgFbgcLmgcLqlmysdMapper.selectzdyf(proname,htd,fbgc,idlist);
                    List<JjgFbgcLmgcLqlmysd> ljxdata = jjgFbgcLmgcLqlmysdMapper.selectljx(proname,htd,fbgc,idlist);
                    List<JjgFbgcLmgcLqlmysd> ljxsddata = jjgFbgcLmgcLqlmysdMapper.selectljxsd(proname,htd,fbgc,idlist);

                    // 主线的部分不用改，因为鉴定表里面主线肯定都是一个部分
                    //沥青路面压实度主线左幅
                    lqlmysdzx(wb, zxzfdata, "沥青路面压实度左幅", "路面面层（主线左幅）");
                    //沥青路面压实度主线右幅
                    lqlmysdzx(wb, zxyfdata, "沥青路面压实度右幅", "路面面层（主线右幅）");

                    // 隧道路面需要改，他们用的都是同样的函数，就更加方便了
                    //沥青路面压实度隧道左幅
                    lqlmysdOther(wb, sdzfdata, "隧道左幅", "隧道路面");
                    //沥青路面压实度隧道右幅
                    lqlmysdOther(wb, sdyfdata, "隧道右幅", "隧道路面");
                    //沥青路面压实度匝道
                    lqlmysdOther(wb, zddata, "沥青路面压实度匝道左幅", "匝道路面");
                    lqlmysdOther(wb, zdyfdata, "沥青路面压实度匝道右幅", "匝道路面");
                    //沥青路面压实度连接线
                    lqlmysdOther(wb, ljxdata, "连接线", "连接线");
                    //沥青路面压实度连接线隧道
                    lqlmysdOther(wb, ljxsddata, "连接线隧道", "连接线隧道");
                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        if (shouldBeCalculate(wb.getSheetAt(j))) {
                            if(wb.getSheetAt(j).getSheetName().contains("隧道")) calculateSheetOther(wb.getSheetAt(j));
                            else calculateSheet(wb.getSheetAt(j));
                            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                        }
                    }
                    /*getTunnelTotal(wb.getSheet("隧道右幅"));
                    getTunnelTotal(wb.getSheet("隧道左幅"));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道右幅"));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道左幅"));*/
                    //删除空表
                    JjgFbgcCommonUtils.deleteEmptySheets(wb);

                    FileOutputStream fileOut = new FileOutputStream(f);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();
                    inputStream.close();
                    out.close();
                }catch (Exception e) {
                    if(f.exists()){
                        f.delete();
                    }
                    e.printStackTrace();
                    throw new JjgysException(20001, "生成鉴定表错误，请检查数据的正确性");
                }

            }
            if (fldata != null && fldata.size()>0){
                //鉴定表要存放的路径
                File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"12沥青路面压实度-(分离上面层).xlsx");
                File fdir = new File(filepath + File.separator + proname + File.separator + htd);
                if (!fdir.exists()) {
                    //创建文件根目录
                    fdir.mkdirs();
                }
                try {
                    String sffl = "是";
                    InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/沥青路面压实度-第二版.xlsx");
                    Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    /*String path = reportPath + "/static/沥青路面压实度-第二版.xlsx";
                    Files.copy(Paths.get(path), new FileOutputStream(f));*/
                    FileInputStream out = new FileInputStream(f);
                    wb = new XSSFWorkbook(out);

                    //左幅 路 上面层
                    List<JjgFbgcLmgcLqlmysd> zxzfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxsmc(proname,htd,fbgc,sffl,"上面层",idlist);
                    //右幅 路 上面层
                    List<JjgFbgcLmgcLqlmysd> zxyfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxyfsmc(proname,htd,fbgc,sffl,"上面层",idlist);

                    //左幅 路 中下面层
                    List<JjgFbgcLmgcLqlmysd> zxzfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxzxmc(proname,htd,fbgc,sffl,idlist);
                    //右幅 路 中下面层
                    List<JjgFbgcLmgcLqlmysd> zxyfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxyfzxmc(proname,htd,fbgc,sffl,idlist);

                    //左幅 隧道 上面层
                    List<JjgFbgcLmgcLqlmysd> sdzfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdzfsmc(proname,htd,fbgc,sffl,"上面层",idlist);
                    //右幅 隧道 上面层
                    List<JjgFbgcLmgcLqlmysd> sdyfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdyfsmc(proname,htd,fbgc,sffl,"上面层",idlist);

                    //左幅 隧道 中下面层
                    List<JjgFbgcLmgcLqlmysd> sdzfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdzfzxmc(proname,htd,fbgc,sffl,idlist);
                    //右幅 隧道 中下面层
                    List<JjgFbgcLmgcLqlmysd> sdyfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdyfzxmc(proname,htd,fbgc,sffl,idlist);

                    // 2025.4.10匝道的sql对桩号的限制不符合交控的规范
                    // 分左右幅的就会有Z和Y的字母，但是不分左右幅的，比如太乙宫互通A匝道，如果不分左右幅，就是AK，默认要放在左幅里面。连接线也是，理想路连接线不分左右幅，就是K···
                    //左幅 匝道 上面层
                    List<JjgFbgcLmgcLqlmysd> zdsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzdsmc(proname,htd,fbgc,idlist,sffl);
                    //右幅 匝道 上面层
                    List<JjgFbgcLmgcLqlmysd> zdyfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzdyfsmc(proname,htd,fbgc,idlist,sffl);
                    //左幅 匝道 中下面层
                    List<JjgFbgcLmgcLqlmysd> zdsxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzdsxmc(proname,htd,fbgc,idlist,sffl);
                    //右幅 匝道 中下面层
                    List<JjgFbgcLmgcLqlmysd> zdyfsxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzdyfsxmc(proname,htd,fbgc,idlist,sffl);

                    List<JjgFbgcLmgcLqlmysd> ljxsmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxsmc(proname,htd,fbgc,idlist,sffl);

                    List<JjgFbgcLmgcLqlmysd> ljxzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxzxmc(proname,htd,fbgc,idlist,sffl);

                    List<JjgFbgcLmgcLqlmysd> ljxsdsmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxsdsmc(proname,htd,fbgc,idlist,sffl);

                    List<JjgFbgcLmgcLqlmysd> ljxsdzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxsdzxmc(proname,htd,fbgc,idlist,sffl);

                    //沥青路面压实度主线左幅上面层
                    lqlmysdzxsmc(wb,zxzfsmcdata,"沥青路面压实度左幅-上面层","路面面层（主线左幅）");

                    //沥青路面压实度主线左幅中下面层
                    lqlmysdzxzxmc(wb,zxzfzxmcdata,"沥青路面压实度左幅-中下面层","路面面层（主线左幅）");

                    //沥青路面压实度主线右幅上面层
                    lqlmysdzxsmc(wb,zxyfsmcdata,"沥青路面压实度右幅-上面层","路面面层（主线右幅）");

                    //沥青路面压实度主线右幅中下面层
                    lqlmysdzxzxmc(wb,zxyfzxmcdata,"沥青路面压实度右幅-中下面层","路面面层（主线右幅）");

                    //沥青路面压实度隧道左幅上面层
                    lqlmysdsdsmc2(wb,sdzfsmcdata,"隧道左幅-上面层","隧道路面");

                    //沥青路面压实度隧道左幅中下面层
                    lqlmysdzxmc2(wb,sdzfzxmcdata,"隧道左幅-中下面层","隧道路面");

                    //沥青路面压实度隧道右幅上面层
                    lqlmysdsdsmc2(wb,sdyfsmcdata,"隧道右幅-上面层","隧道路面");

                    //沥青路面压实度隧道右幅中下面层
                    lqlmysdzxmc2(wb,sdyfzxmcdata,"隧道右幅-中下面层","隧道路面");

                    //沥青路面压实度匝道上面层
                    lqlmysdsdsmc2(wb,zdsmcdata,"沥青路面压实度匝道左幅-上面层","匝道路面");

                    //沥青路面压实度匝道右幅上面层
                    lqlmysdsdsmc2(wb,zdyfsmcdata,"沥青路面压实度匝道右幅-上面层","匝道路面");

                    //沥青路面压实度匝道中下面层
                    lqlmysdzxmc2(wb,zdsxmcdata,"沥青路面压实度匝道左幅-中下面层","匝道路面");

                    //沥青路面压实度匝道右幅中下面层
                    lqlmysdzxmc2(wb,zdyfsxmcdata,"沥青路面压实度匝道右幅-中下面层","匝道路面");

                    //沥青路面压实度连接线上面层
                    lqlmysdsdsmc2(wb,ljxsmcdata,"连接线-上面层","连接线");

                    //沥青路面压实度连接线中下面层
                    lqlmysdzxmc2(wb,ljxzxmcdata,"连接线-中下面层","连接线");

                    //沥青路面压实度连接线隧道上面层
                    lqlmysdsdsmc2(wb,ljxsdsmcdata,"连接线隧道-上面层","连接线隧道");

                    //沥青路面压实度连接线隧道中下面层
                    lqlmysdzxmc2(wb,ljxsdzxmcdata,"连接线隧道-中下面层","连接线隧道");

                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                        if (shouldBeCalculate(wb.getSheetAt(j))) {
                            if(wb.getSheetAt(j).getSheetName().contains("隧道")) calculateSheetOther(wb.getSheetAt(j));
                            else calculateSheet(wb.getSheetAt(j));
                            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                        }
                    }
                    /*getTunnelTotal(wb.getSheet("隧道右幅-上面层"));
                    getTunnelTotal(wb.getSheet("隧道右幅-中下面层"));
                    getTunnelTotal(wb.getSheet("隧道左幅-上面层"));
                    getTunnelTotal(wb.getSheet("隧道左幅-中下面层"));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道右幅-上面层"));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道右幅-中下面层"));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道左幅-上面层"));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道左幅-中下面层"));*/
                    //删除空表
                    JjgFbgcCommonUtils.deleteEmptySheets(wb);

                    FileOutputStream fileOut = new FileOutputStream(f);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();
                    inputStream.close();
                    out.close();
                }catch (Exception e) {
                    if(f.exists()){
                        f.delete();
                    }
                    e.printStackTrace();
                    throw new JjgysException(20001, "生成鉴定表错误，请检查数据的正确性");
                }

            }
            wb.close();
            return true;
        }
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    boolean[] flag = new boolean[2];
                    for(int j = 0; j < 2; j++){
                        if(i+j < data.size()){
                            if("中面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index).getCell(2).setCellValue("中面层");
                                fillsdsmcCommonCellData(sheet, index+0, data.get(i+j));
                                flag[0] = true;
                            }
                            else if("下面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index+1).getCell(2).setCellValue("下面层");
                                fillsdsmcCommonCellData(sheet, index+1, data.get(i+j));
                                flag[1] = true;
                            }
                        }
                        else{
                            break;
                        }
                    }
                    for (int j = 0; j < flag.length; j++) {
                        if(!flag[j]){
                            fillCommonCell_Null_Data(sheet, index+j);
                        }
                    }
                    i++;
                    index+=2;
                    sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 0, 0));
                    if (data.get(i).getLqs().contains("隧道")){
                        sheet.getRow(index-2).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                    }else {
                        sheet.getRow(index-2).getCell(0).setCellValue(zh);
                    }

                    sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 1, 1));
                    sheet.getRow(index-2).getCell(1).setCellValue(data.get(i).getQywz());
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }

        }



    }

    /**
     * 隧道上面层：用于匝道、连接线的路面显示（修改后）
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdsdsmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;
            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if("上面层".equals(data.get(i).getCw())){
                        sheet.getRow(index).getCell(2).setCellValue("上面层");
                        fillsdsmcCommonCellData(sheet, index, data.get(i));
                    }
//                    if (data.get(i).getLqs().contains("隧道")){
//                        sheet.getRow(index).getCell(0).setCellValue(data.get(i).getLqs()+zh);
//                    }else {
//                        sheet.getRow(index).getCell(0).setCellValue(zh);
//                    }
                    if (data.get(i).getLqs().contains("隧")){ // 这里可能是因为有的数据命名不规范所以进行补充
                        sheet.getRow(index).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                    }else {
                        sheet.getRow(index).getCell(0).setCellValue(zh);
                    }

                    sheet.getRow(index).getCell(1).setCellValue(data.get(i).getQywz());
                    index++;
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
        }

    }

    /**
     * 隧道上面层：专门用于隧道上面层
     * @param wb
     * @param data1
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdsdsmc2(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data1, String sheetname, String fbgcname) {
        // 先对data进行分类,相同lqsName放在一起
        List<List<JjgFbgcLmgcLqlmysd>> dataList = new ArrayList<>();
        if (data1.size() > 0 && data1 != null){
            String lqsName = data1.get(0).getLqs();
            List<JjgFbgcLmgcLqlmysd> temp = new ArrayList<>();
            for(JjgFbgcLmgcLqlmysd item : data1){
                if(lqsName.equals(item.getLqs())){
                    temp.add(item);
                }else{
                    dataList.add(temp);
                    lqsName = item.getLqs();
                    temp = new ArrayList<>();
                    temp.add(item);
                }
            }
            // 把temp里面剩余的加入到dataList中
            dataList.add(temp);
        }else return;

        int row = 0;
        int tempRow = 0;
        for(List<JjgFbgcLmgcLqlmysd> data : dataList){
            tempRow = createTable(wb,getZXZtableNum(data.size()),sheetname, row);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(row + 1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(row + 1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(row + 2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(row + 2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(row + 3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(row + 3).getCell(9).setCellValue(date);
            int index = 6;
            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if("上面层".equals(data.get(i).getCw())){
                        sheet.getRow(row + index).getCell(2).setCellValue("上面层");
                        fillsdsmcCommonCellData(sheet, row + index, data.get(i));
                    }
//                    if (data.get(i).getLqs().contains("隧道")){
//                        sheet.getRow(index).getCell(0).setCellValue(data.get(i).getLqs()+zh);
//                    }else {
//                        sheet.getRow(index).getCell(0).setCellValue(zh);
//                    }
                    if (data.get(i).getLqs().contains("隧")){ // 这里可能是因为有的数据命名不规范所以进行补充
                        sheet.getRow(row + index).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                    }else {
                        sheet.getRow(row + index).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                    }

                    sheet.getRow(row + index).getCell(1).setCellValue(data.get(i).getQywz());
                    index++;
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
            row = tempRow;
        }

    }

    /**
     * 专门用于隧道的下面层
     * @param wb
     * @param data1
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxmc2(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data1, String sheetname, String fbgcname) {
        // 先对data进行分类,相同lqsName放在一起
        List<List<JjgFbgcLmgcLqlmysd>> dataList = new ArrayList<>();
        if (data1.size() > 0 && data1 != null){
            String lqsName = data1.get(0).getLqs();
            List<JjgFbgcLmgcLqlmysd> temp = new ArrayList<>();
            for(JjgFbgcLmgcLqlmysd item : data1){
                if(lqsName.equals(item.getLqs())){
                    temp.add(item);
                }else{
                    dataList.add(temp);
                    lqsName = item.getLqs();
                    temp = new ArrayList<>();
                    temp.add(item);
                }
            }
            // 把temp里面剩余的加入到dataList中
            dataList.add(temp);
        }else return;

        int row = 0;
        int tempRow = 0;

        for(List<JjgFbgcLmgcLqlmysd> data : dataList){
            tempRow = createTable(wb,getZXZtableNum(data.size()),sheetname, row);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(row + 1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(row + 1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(row + 2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(row + 2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(row + 3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(row + 3).getCell(9).setCellValue(date);
            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    boolean[] flag = new boolean[2];
                    for(int j = 0; j < 2; j++){
                        if(i+j < data.size()){
                            if("中面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(row + index).getCell(2).setCellValue("中面层");
                                fillsdsmcCommonCellData(sheet, row + index+0, data.get(i+j));
                                flag[0] = true;
                            }
                            else if("下面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(row + index+1).getCell(2).setCellValue("下面层");
                                fillsdsmcCommonCellData(sheet, row + index+1, data.get(i+j));
                                flag[1] = true;
                            }
                        }
                        else{
                            break;
                        }
                    }
                    for (int j = 0; j < flag.length; j++) {
                        if(!flag[j]){
                            fillCommonCell_Null_Data(sheet, row + index+j);
                        }
                    }
                    i++;
                    index+=2;
                    sheet.addMergedRegion(new CellRangeAddress(row + index-2, row + index-1, 0, 0));
                    if (data.get(i).getLqs().contains("隧道")){
                        sheet.getRow(row + index-2).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                    }else {
                        sheet.getRow(row + index-2).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                    }

                    sheet.addMergedRegion(new CellRangeAddress(row + index-2, row + index-1, 1, 1));
                    sheet.getRow(row + index-2).getCell(1).setCellValue(data.get(i).getQywz());
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
            row = tempRow;
        }



    }

    /**
     *
     * @param sheet
     * @param index
     * @param row
     */
    private void fillsdsmcCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd row) {
        if (row.getGzsjzl()!=null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(row.getGzsjzl()));
            sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
            sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(row.getZdllmdgdz()));
        }else {
            sheet.getRow(index).getCell(3).setCellValue("-");
        }
        if (row.getSjszzl()!=null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(row.getSjszzl()));
        }else {
            sheet.getRow(index).getCell(4).setCellValue("-");
        }
        if (row.getSjbgzl()!=null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(row.getSjbgzl()));
        }else {
            sheet.getRow(index).getCell(5).setCellValue("-");
        }

        if (row.getSysbzmd()!=null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(row.getSysbzmd()));
        }else {
            sheet.getRow(index).getCell(7).setCellValue("-");
        }

        if (row.getZdllmd()!=null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(row.getZdllmd()));
        }else {
            sheet.getRow(index).getCell(8).setCellValue("-");
        }

        if (row.getSysbzmdgdz()!=null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(row.getSysbzmdgdz())-1);
        }else {
            sheet.getRow(index).getCell(11).setCellValue("-");
        }

        if (row.getZdllmdgdz()!=null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(row.getZdllmdgdz())-1);
        }else {
            sheet.getRow(index).getCell(12).setCellValue("-");
        }

    }


    /**
     * 分离
     * 沥青路面压实度主线左幅中下面层
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxzxmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧道")){
                        i+=1;
                        index += 2;
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 0, 1));
                        sheet.getRow(index-2).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 2, 8));
                        sheet.getRow(index-2).getCell(2).setCellValue("见隧道路面压实度鉴定表");

                    }
                    else{
                        boolean[] flag = new boolean[2];
                        for(int j = 0; j < 2; j++){
                            if(i+j < data.size()){
                                if("中面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index).getCell(2).setCellValue("中面层");
                                    fillzxsmcCommonCellData(sheet, index, data.get(i+j));
                                    flag[0] = true;
                                }else if("下面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+1).getCell(2).setCellValue("下面层");
                                    fillzxsmcCommonCellData(sheet, index+1, data.get(i+j));
                                    flag[1] = true;
                                }
                            }
                            else{
                                break;
                            }
                        }
                        for (int j = 0; j < flag.length; j++) {
                            if(!flag[j]){
                                fillCommonCell_Null_Data(sheet, index+j);
                            }
                        }
                        i++;
                        index+=2;
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 0, 0));
                        sheet.getRow(index-2).getCell(0).setCellValue(zh);
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 1, 1));
                        sheet.getRow(index-2).getCell(1).setCellValue(data.get(i).getQywz());
                    }
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }

            }

        }

    }


    /**
     * 分离
     * 沥青路面压实度主线左幅上面层
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxsmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;


            for(int i =0; i < data.size(); i++){

                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧道") ){
                        //i++;
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 1));
                        sheet.getRow(index).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 2, 8));
                        sheet.getRow(index).getCell(2).setCellValue("见隧道路面压实度鉴定表");
                        index ++;
                    }
                    else{
                        if("上面层".equals(data.get(i).getCw())){
                            sheet.getRow(index).getCell(2).setCellValue("上面层");
                            fillzxsmcCommonCellData(sheet, index, data.get(i));
                        }
                        if (data.get(i).getLqs().contains("桥")){
                            sheet.getRow(index).getCell(0).setCellValue(data.get(i).getLqs()+zh);
                        }else {
                            sheet.getRow(index).getCell(0).setCellValue(zh);
                        }

                        sheet.getRow(index).getCell(1).setCellValue(data.get(i).getQywz());
                        index++;
                    }
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
        }

    }

    /**
     * 分离
     * 上面层
     * @param sheet
     * @param index
     * @param row
     */
    private void fillzxsmcCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd row) {
        if(!"".equals(row.getGzsjzl()) && row.getGzsjzl() != null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(row.getGzsjzl()));
        }else {
            sheet.getRow(index).getCell(3).setCellValue("-");
        }
        if(!"".equals(row.getSjszzl()) && row.getSjszzl() != null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(row.getSjszzl()));
        }else {
            sheet.getRow(index).getCell(4).setCellValue("-");
        }
        if(!"".equals(row.getSjbgzl()) && row.getSjbgzl() != null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(row.getSjbgzl()));
        }else {
            sheet.getRow(index).getCell(5).setCellValue("-");
        }
        if(!"".equals(row.getSysbzmd()) && row.getSysbzmd() != null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(row.getSysbzmd()));
        }else {
            sheet.getRow(index).getCell(7).setCellValue("-");
        }
        if(!"".equals(row.getZdllmd()) && row.getZdllmd() != null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(row.getZdllmd()));
        }else {
            sheet.getRow(index).getCell(8).setCellValue("-");
        }
        if(!"".equals(row.getSysbzmdgdz()) && row.getSysbzmdgdz() != null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(row.getSysbzmdgdz())-1);
        }
        if(!"".equals(row.getZdllmdgdz()) && row.getZdllmdgdz() != null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(row.getZdllmdgdz())-1);
        }
        if(!"".equals(row.getSysbzmdgdz()) && row.getSysbzmdgdz() != null){
            sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
        }
        if(!"".equals(row.getZdllmdgdz()) && row.getZdllmdgdz() != null){
            sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(row.getZdllmdgdz()));
        }

    }

    /**
     * 不分离上面层
     * 将每个隧道的数据汇总
     * @param sheet
     */
    private void getTunnelTotal(XSSFSheet sheet){
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow startrow = null;
        XSSFRow endrow = null;
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && row.getCell(0) != null && !"".equals(row.getCell(0).toString())){
                if(!name.equals(row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(15).setCellFormula("COUNT("
                            +startrow.getCell(9).getReference()+":"
                            +endrow.getCell(9).getReference()+")");//=COUNT(J7:J21)
                    startrow.createCell(16).setCellFormula("COUNTIF("
                            +startrow.getCell(13).getReference()+":"
                            +endrow.getCell(13).getReference()+",\"√\")");//=COUNTIF(N7:N22,"√")
                    startrow.createCell(17).setCellFormula(startrow.getCell(16).getReference()+"*100/"
                            +startrow.getCell(15).getReference());
                    name = row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "");
                    startrow = row;
                }
                if("".equals(name)){
                    name = row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "");
                    startrow = row;
                }

            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
            }
            if ("评定".equals(row.getCell(0).toString())) {
                break;
            }
        }
    }

    /**
     * 不分离上面层
     * 判断此sheet是否需要进行计算
     * @param sheet
     * @return
     */
    public boolean shouldBeCalculate(XSSFSheet sheet) {
        String title = null;
        title = sheet.getRow(0).getCell(0).getStringCellValue();
        if (title.endsWith("鉴定表")) {
            return true;
        }
        return false;
    }

    /**
     * 不分离上面层
     * @param sheet
     */
    private void calculateSheet(XSSFSheet sheet) {
        XSSFRow row = null;

        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 计算表格中间的数据
            fillBodyRow(row);
            // 完成表格尾部数据的计算及填写
            if ("评定".equals(row.getCell(0).toString())) {
                fillBottomConclusion(sheet, i, 7);
                // 文件的所有计算都已完成，跳出循环，结束剩下2行的读取
                break;
            }
        }

        /*
         * 统计最大值最小值，生成报告表格的时候要用
         */
        row = sheet.getRow(sheet.getPhysicalNumberOfRows()-3);
        row.createCell(15).setCellValue("最大值");
        row.createCell(16).setCellFormula("ROUND("+
                "MAX("+sheet.getRow(6).getCell(9).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MAX(J7:J57)
        row.createCell(17).setCellFormula("ROUND("+"MAX("+sheet.getRow(6).getCell(10).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MAX(J7:J57)

        row = sheet.getRow(sheet.getPhysicalNumberOfRows()-2);
        row.createCell(15).setCellValue("最小值");
        row.createCell(16).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(9).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MIN(J7:J57)
        row.createCell(17).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(10).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MIN(J7:J57)

        //if (!sheet.getSheetName().contains("隧道")){
            //总点数
            //XSSFRow wrow = null;
            boolean flag = false;
            XSSFRow startrow = null;
            XSSFRow endrow = sheet.getRow(sheet.getPhysicalNumberOfRows()-4);
            for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
                startrow = sheet.getRow(i);
                if (startrow == null) {
                    continue;
                }
                if(flag && startrow.getCell(0) != null && !"".equals(startrow.getCell(0).toString())){
                    startrow = sheet.getRow(i);
                    startrow.createCell(15).setCellFormula("COUNT("
                            +startrow.getCell(9).getReference()+":"
                            +endrow.getCell(9).getReference()+")");//=COUNT(J7:J21)
                    startrow.createCell(16).setCellFormula("COUNTIF("
                            +startrow.getCell(13).getReference()+":"
                            +endrow.getCell(13).getReference()+",\"√\")");//=COUNTIF(N7:N22,"√")
                    startrow.createCell(17).setCellFormula(startrow.getCell(16).getReference()+"*100/"
                            +startrow.getCell(15).getReference());
                    break;
                }
                if ("桩号".equals(startrow.getCell(0).toString())) {
                    flag = true;
                    i++;

                }
                if ("评定".equals(startrow.getCell(0).toString())) {
                    break;
                }
            }


        //}


    }

    /**
     * 不分离上面层
     * @param sheet
     */
    private void calculateSheetOther(XSSFSheet sheet) {
        XSSFRow row = null;
        int startRow1 = 0; // 记录每个表格高度，有评定行判断所以不用endRow
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 遇到表头也要跳过
            row.getCell(0).setCellType(CellType.STRING);
            if(row.getCell(0).toString().contains("鉴定表")){
                startRow1 = i;
                i += 5; // for循环里面还有一次++,所以只加5
                continue;
            }
            // 计算表格中间的数据
            fillBodyRow(row);
            // 完成表格尾部数据的计算及填写
            if ("评定".equals(row.getCell(0).toString())) {
                fillBottomConclusion(sheet, i, startRow1 + 7); // 因为excel公式里面是以1开始计数的

                /*
                 * 统计最大值最小值，生成报告表格的时候要用
                 */
                row.createCell(15).setCellValue("最大值");
                row.createCell(16).setCellFormula("ROUND("+
                        "MAX("+sheet.getRow(6).getCell(9).getReference()+":"+
                        sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MAX(J7:J57)
                row.createCell(17).setCellFormula("ROUND("+"MAX("+sheet.getRow(6).getCell(10).getReference()+":"+
                        sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MAX(J7:J57)

                row = sheet.getRow(i+1);
                row.createCell(15).setCellValue("最小值");
                row.createCell(16).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(9).getReference()+":"+
                        sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MIN(J7:J57)
                row.createCell(17).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(10).getReference()+":"+
                        sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MIN(J7:J57)

                // 统计总检测点数、合格点数和合格率
                boolean flag = false;
                XSSFRow startrow = null;
                XSSFRow endrow = sheet.getRow(i);
                for (int j = startRow1; j <= i + 2; j++) {
                    startrow = sheet.getRow(j);
                    if (startrow == null) {
                        continue;
                    }
                    if(flag && startrow.getCell(0) != null && !"".equals(startrow.getCell(0).toString())){
                        startrow = sheet.getRow(j);
                        startrow.createCell(15).setCellFormula("COUNT("
                                +startrow.getCell(9).getReference()+":"
                                +endrow.getCell(9).getReference()+")");//=COUNT(J7:J21)
                        startrow.createCell(16).setCellFormula("COUNTIF("
                                +startrow.getCell(13).getReference()+":"
                                +endrow.getCell(13).getReference()+",\"√\")");//=COUNTIF(N7:N22,"√")
                        startrow.createCell(17).setCellFormula(startrow.getCell(16).getReference()+"*100/"
                                +startrow.getCell(15).getReference());
                        break;
                    }
                    if ("桩号".equals(startrow.getCell(0).toString())) {
                        flag = true;
                        j++;

                    }
                    if ("评定".equals(startrow.getCell(0).toString())) {
                        break;
                    }
                }
            }


        }







        //}


    }

    /**
     * 不分离上面层
     * 计算表格尾部汇总结果
     * @param sheet
     * @param i
     */
    public void fillBottomConclusion(XSSFSheet sheet, int i, int startRow) {
        /*
         * 表格尾部有3行，由于数据间有联系，所以不能逐行计算 计算顺序为：平均值，监测点数->标准差->代表值->结论->合格点数->合格率
         */
        // 定义3个局部变量，表示表尾的1,2,3行
        XSSFRow row1, row2, row3;
        row1 = sheet.getRow(i);
        row2 = sheet.getRow(i + 1);
        row3 = sheet.getRow(i + 2);
        XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(sheet.getWorkbook());

        // 试验室标准密度控制->平均值
        row1.getCell(4).setCellType(CellType.NUMERIC);
        row1.getCell(4).setCellFormula(
                "AVERAGE(" + "J" + startRow + ":" + "J" + row1.getRowNum() + ")");// AVERAGE(J7:J24)
        // 试验室标准密度控制->监测点数
        row1.getCell(6).setCellType(CellType.NUMERIC);
        row1.getCell(6).setCellFormula(
                "COUNT(" + "J" + startRow + ":" + "J" + row1.getRowNum() + ")");// COUNT(J7:J24)

        // 最大理论密度控制->平均值
        row1.getCell(10).setCellType(CellType.NUMERIC);
        row1.getCell(10).setCellFormula(
                "AVERAGE(" + "K" + startRow + ":" + "K" + row1.getRowNum() + ")");// AVERAGE(K7:K24)
        // 最大理论密度控制->监测点数
        row1.getCell(12).setCellType(CellType.NUMERIC);
        row1.getCell(12).setCellFormula(
                "COUNT(" + "K" + startRow + ":" + "K" + row1.getRowNum() + ")");// COUNT(K7:K24)

        // 标准差
        row2.getCell(4).setCellType(CellType.NUMERIC);
        row2.getCell(4).setCellFormula(
                "STDEV(" + "J" + startRow + ":" + "J" + (row2.getRowNum() - 1) + ")");// =STDEV(J7:J24)
        // 标准差
        row2.getCell(10).setCellType(CellType.NUMERIC);
        row2.getCell(10).setCellFormula(
                "STDEV(" + "K" + startRow + ":" + "K" + (row2.getRowNum() - 1) + ")");// =STDEV(K7:K24)

        // 代表值
        row3.getCell(4).setCellType(CellType.NUMERIC);
        row3.getCell(4).setCellFormula(
                "ROUND(" + row1.getCell(4).getReference() + "-("
                        + row2.getCell(4).getReference() + "*VLOOKUP("
                        + row1.getCell(6).getReference()
                        + ",保证率系数!$A:$D,3)),1)");// ROUND(E40-(E41*VLOOKUP(G40,保证率系数!$A:$D,3,)),1)
        row3.getCell(10).setCellType(CellType.NUMERIC);
        row3.getCell(10).setCellFormula(
                "ROUND(" + row1.getCell(10).getReference() + "-("
                        + row2.getCell(10).getReference() + "*VLOOKUP("
                        + row1.getCell(12).getReference()
                        + ",保证率系数!$A:$D,3)),1)");// =ROUND(K40-(K41*VLOOKUP(M40,保证率系数!$A:$D,3,)),1)
        // 右下角结论
        row1.getCell(14).setCellType(CellType.NUMERIC);
        row1.getCell(14).setCellFormula(
                "IF((" + row3.getCell(4).getReference() + ">="
                        + row2.getCell(2).getReference() + ")*AND("
                        + row3.getCell(10).getReference() + ">="
                        + row2.getCell(8).getReference() + "),\"合格\",\"不合格\")");// IF((E42>=C41)*AND(K42>=I41),"合格","不合格")
        // 计算右侧结论
        fillRightConclusion(sheet);
        // 合格点数 =COUNTIF(N7:N24,"√")
        row2.getCell(6).setCellType(CellType.NUMERIC);
        row2.getCell(6).setCellFormula(
                "COUNTIF(" + "N" + startRow + ":" + row1.getCell(13).getReference()
                        + ",\"√\")");
        row2.getCell(12).setCellFormula(
                "COUNTIF(" + "N" + startRow + ":" + row1.getCell(13).getReference()
                        + ",\"√\")");

        BigDecimal row5value = new BigDecimal(e.evaluate(row2.getCell(12)).getNumberValue());
        row2.getCell(12).setCellFormula(null);
        row2.getCell(12).setCellValue(row5value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());

        // 合格率 =G41/G40*100
        row3.getCell(6).setCellFormula(
                row2.getCell(6).getReference() + "/"
                        + row1.getCell(6).getReference() + "*100");
        row3.getCell(12).setCellFormula(
                row2.getCell(12).getReference() + "/"
                        + row1.getCell(12).getReference() + "*100");
    }

    /**
     * 不分离上面层
     * =IF(O58="合格",IF((J13>=L13)*(K13>M13),"√",""),"")
     * =IF(J13="","",IF(N13="","×",""))
     * =IF((E60>=C59)*AND(K60>=I59),"合格","不合格")
     * 根据表格的一些汇总数据，完成右边的结论计算
     * @param sheet
     */
    public void fillRightConclusion(XSSFSheet sheet) {
        XSSFRow row = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 判断前面的数据有值时，才进行结论的计算
            if (row.getCell(9).getCellType() == Cell.CELL_TYPE_FORMULA
                    && row.getCell(10).getCellType() == Cell.CELL_TYPE_FORMULA
                    && row.getCell(11).getCellType() == Cell.CELL_TYPE_NUMERIC
                    && row.getCell(12).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                // 计算结论
                row.getCell(13).setCellFormula(
                        "IF(" + getLastCell(sheet).getReference()
                                + "=\"合格\",IF(("
                                + row.getCell(9).getReference() + ">="
                                + row.getCell(11).getReference() + ")*("
                                + row.getCell(10).getReference() + ">"
                                + row.getCell(12).getReference()
                                + "),\"√\",\"\"),\"\")");// =IF($O$40="合格",IF((J7>=L7)*(K7>M7),"√",""),"")
                row.getCell(14).setCellFormula(
                        "IF(" + row.getCell(9).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(13).getReference()
                                + "=\"\",\"×\",\"\"))");// =IF(J13="","",IF(N13="","×",""))
            }
        }
    }

    /**
     * 不分离上面层
     * 得到最后一个单元格，计算右侧结论要使用
     * @param sheet
     * @return
     */
    public XSSFCell getLastCell(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFCell cell = null;
        row = sheet.getRow(sheet.getLastRowNum() - 2);
        cell = row.getCell(14);
        return cell;
    }

    /**
     * 不分离上面层
     *  计算表格中一行的数据
     * @param row
     */
    public void fillBodyRow(XSSFRow row) {
        if (row.getCell(3).getCellTypeEnum() == CellType.NUMERIC && row.getCell(4).getCellTypeEnum() == CellType.NUMERIC && row.getCell(5).getCellTypeEnum() == CellType.NUMERIC) {
            row.getCell(6).setCellType(CellType.NUMERIC);
            row.getCell(6).setCellFormula(
                    row.getCell(3).getReference() + "/" + "("
                            + row.getCell(5).getReference() + "-"
                            + row.getCell(4).getReference() + ")");// D7/(F7-E7)
        }
        if (row.getCell(7).getCellTypeEnum() == CellType.NUMERIC
                && row.getCell(8).getCellTypeEnum() == CellType.NUMERIC) {
            // 试验室
            row.getCell(9).setCellType(CellType.NUMERIC);
            row.getCell(9).setCellFormula("ROUND("+
                    row.getCell(6).getReference() + "/"
                    + row.getCell(7).getReference() + "*" + "100,1)");// G7/H7*100
            // 最大理论
            row.getCell(10).setCellType(CellType.NUMERIC);
            row.getCell(10).setCellFormula("ROUND("+
                    row.getCell(6).getReference() + "/"
                    + row.getCell(8).getReference() + "*" + "100,1)");// G7/I7*100
        }
    }


    /**
     * 不分离上面层
     *隧道 匝道 连接线 连接线隧道
     * @param wb
     * @param data1
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdOther(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data1, String sheetname, String fbgcname) {
        // 函数进来以后，就已经拿到了所有的（比如说）隧道信息，所以在这里就要对隧道名称进行判断，不同类型就要另外创建表格了
        // 先对data进行分类,相同lqsName放在一起
        List<List<JjgFbgcLmgcLqlmysd>> dataList = new ArrayList<>();
        if (data1.size() > 0 && data1 != null){
            String lqsName = data1.get(0).getLqs();
            List<JjgFbgcLmgcLqlmysd> temp = new ArrayList<>();
            for(JjgFbgcLmgcLqlmysd item : data1){
                if(lqsName.equals(item.getLqs())){
                    temp.add(item);
                }else{
                    dataList.add(temp);
                    lqsName = item.getLqs();
                    temp = new ArrayList<>();
                    temp.add(item);
                }
            }
            // 把temp里面剩余的加入到dataList中
            dataList.add(temp);
        }else return;
        int row = 0;
        int tempRow = 0;
        for(List<JjgFbgcLmgcLqlmysd> data : dataList){
            // 由于不止一次进入表格创建函数，所以得维护一个row。重新写一个createTable函数，因为主线也是这样生成的
            tempRow = createTable(wb,getZXZtableNum(data.size()),sheetname, row);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(row + 1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(row + 1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(row + 2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(row + 2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(row + 3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(row + 3).getCell(9).setCellValue(date);
            int index = 6;
            for (int i = 0; i < data.size(); i++) {
                if (zh.equals(data.get(i).getZh())) {
                    int count = 0;
                    for (int j = 0; j < 3; j++) {
                        if (i + j < data.size()) {
                            if ("上面层".equals(data.get(i + j).getCw()) && data.get(i).getZh().equals(data.get(i + j).getZh())) {
                                sheet.getRow(row + index).getCell(2).setCellValue(data.get(i + j).getCw());
                                fillsdCommonCellData(sheet, row + index, data.get(i + j));
                                count++;
                            }
                            if ("中面层".equals(data.get(i + j).getCw()) && data.get(i).getZh().equals(data.get(i + j).getZh())) {
                                sheet.getRow(row + index + 1).getCell(2).setCellValue(data.get(i + j).getCw());
                                fillsdCommonCellData(sheet, row + index + 1, data.get(i + j));
                                count++;
                            }
                            if ("下面层".equals(data.get(i + j).getCw()) && data.get(i).getZh().equals(data.get(i + j).getZh())) {
                                sheet.getRow(row + index + 2).getCell(2).setCellValue(data.get(i + j).getCw());
                                fillsdCommonCellData(sheet, row + index + 2, data.get(i + j));
                                count++;
                            }
                        } else {
                            break;
                        }
                    }
                    i += 2;
                    index+=count;
                    if(count > 1){
                        sheet.addMergedRegion(new CellRangeAddress(row + index-count, row + index-1, 0, 0));
                    }
                    sheet.getRow(row + index-count).getCell(0).setCellValue(type + zh);
                    /*if (data.get(i).getLqs().equals("路") || data.get(i).getLqs().contains("隧") || data.get(i).getLqs().contains("连接线") || data.get(i).getLqs().equals("桥")){
                        sheet.getRow(index-count).getCell(0).setCellValue(zh+type);

                    }else {
                        sheet.getRow(index-count).getCell(0).setCellValue(zh);
                    }*/
                    if(count > 1){
                        sheet.addMergedRegion(new CellRangeAddress(row + index-count, row + index-1, 1, 1));
                    }
                    //sheet.getRow(index-count).getCell(1).setCellValue(data.get(i)[8]);
                } else {
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
            row = tempRow;
        }
    }

    /**
     * 不分离上面层
     * 沥青路面压实度主线
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzx(XSSFWorkbook wb,List<JjgFbgcLmgcLqlmysd> data,String sheetname,String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);

            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧")){
                        i+=2;
                        index += 3;
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 1));
                        sheet.getRow(index-3).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 2, 8));
                        sheet.getRow(index-3).getCell(2).setCellValue("见隧道路面压实度鉴定表");

                    }else {
                        for(int j = 0; j < 3; j++){
                            if(i+j < data.size()){
                                if("上面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index, data.get(i+j));
                                    //flag[0] = true;
                                }
                                if("中面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+1).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index+1, data.get(i+j));
                                    //flag[1] = true;
                                }
                                if("下面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+2).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index+2, data.get(i+j));
                                    //flag[2] = true;
                                }
                            }else {
                                break;
                            }
                        }
                        i+=2;
                        index+=3;
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 0));
                        sheet.getRow(index-3).getCell(0).setCellValue(zh);
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 1, 1));
                    }
                }else {
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }

        }

    }

    /**
     * 不分离上面层
     * 沥青路面压实度左幅的层位没有数据时，以-显示
     * @param sheet
     * @param index
     */
    private void fillCommonCell_Null_Data(XSSFSheet sheet, int index) {
        int[] num = new int[]{2,3,4,5,7,8,9,10,11,12};
        for (int i = 0; i < num.length; i++) {
            if(sheet.getRow(index).getCell(num[i]) == null){
                sheet.getRow(index).createCell(num[i]);
            }
            sheet.getRow(index).getCell(num[i]).setCellValue("-");
        }
    }


    /**
     * 不分离上面层
     * @param sheet
     * @param index
     * @param zfdata
     */
    private void fillsdCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd zfdata) {
        sheet.getRow(index).getCell(0).setCellValue(zfdata.getLqs()+zfdata.getZh());
        sheet.getRow(index).getCell(1).setCellValue(zfdata.getQywz());
        if(!"".equals(zfdata.getGzsjzl()) && zfdata.getGzsjzl() != null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(zfdata.getGzsjzl()));
        }else {
            sheet.getRow(index).getCell(3).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjszzl()) && zfdata.getSjszzl()!=null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(zfdata.getSjszzl()));
        }else {
            sheet.getRow(index).getCell(4).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjbgzl()) && zfdata.getSjbgzl()!=null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(zfdata.getSjbgzl()));
        }else {
            sheet.getRow(index).getCell(5).setCellValue("-");
            sheet.getRow(index).getCell(6).setCellValue("-");
        }
        if(!"".equals(zfdata.getSysbzmd()) && zfdata.getSysbzmd()!=null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(zfdata.getSysbzmd()));
        }else {
            sheet.getRow(index).getCell(7).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmd()) && zfdata.getZdllmd()!=null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmd()));
        }else {
            sheet.getRow(index).getCell(8).setCellValue("-");
            sheet.getRow(index).getCell(9).setCellValue("-");
            sheet.getRow(index).getCell(10).setCellValue("-");
        }

        if(!"".equals(zfdata.getSysbzmdgdz()) && zfdata.getSysbzmdgdz()!=null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz())-1);
        }else {
            sheet.getRow(index).getCell(11).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmdgdz()) && zfdata.getZdllmdgdz()!=null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz())-1);
        }else {
            sheet.getRow(index).getCell(12).setCellValue("-");
            sheet.getRow(index).getCell(13).setCellValue("-");
            sheet.getRow(index).getCell(14).setCellValue("-");
        }
        if(!"".equals(zfdata.getSysbzmdgdz()) && zfdata.getSysbzmdgdz()!=null){
            sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        }
        if(!"".equals(zfdata.getZdllmdgdz()) && zfdata.getZdllmdgdz()!=null){
            sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));

        }
    }


    /**
     * 不分离上面层
     * 沥青路面压实度左幅数据
     * @param sheet
     * @param index
     * @param zfdata
     */
    private void fillLmzfCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd zfdata) {
        if (zfdata.getLmlx().contains("桥")){
            sheet.getRow(index).getCell(0).setCellValue(zfdata.getZh()+zfdata.getLqs());
        }else {
            sheet.getRow(index).getCell(0).setCellValue(zfdata.getZh());
        }

        sheet.getRow(index).getCell(1).setCellValue(zfdata.getQywz());
        //sheet.getRow(index).getCell(2).setCellValue(zfdata.getCw());
        if(!"".equals(zfdata.getGzsjzl()) && zfdata.getGzsjzl() != null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(zfdata.getGzsjzl()));
            sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
            sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
        }else {
            sheet.getRow(index).getCell(3).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjszzl()) && zfdata.getSjszzl()!=null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(zfdata.getSjszzl()));
        }else {
            sheet.getRow(index).getCell(4).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjbgzl()) && zfdata.getSjbgzl()!=null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(zfdata.getSjbgzl()));
        }else {
            sheet.getRow(index).getCell(5).setCellValue("-");
        }
        if(!"".equals(zfdata.getSysbzmd()) && zfdata.getSysbzmd()!=null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(zfdata.getSysbzmd()));
        }else {
            sheet.getRow(index).getCell(7).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmd()) && zfdata.getZdllmd()!=null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmd()));
        }else {
            sheet.getRow(index).getCell(8).setCellValue("-");
        }

        if(!"".equals(zfdata.getSysbzmdgdz()) && zfdata.getSysbzmdgdz()!=null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz())-1);
        }else {
            sheet.getRow(index).getCell(11).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmdgdz()) && zfdata.getZdllmdgdz()!=null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz())-1);
        }else {
            sheet.getRow(index).getCell(12).setCellValue("-");
        }

    }


    /**
     *不分离上面层
     * @param wb
     * @param tableNum
     */
    private void createZXZTable(XSSFWorkbook wb,int tableNum,String sheetname) {
        if(tableNum == 0){
            return;
        }
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 23, (i - 1) * 18 + 24);
            }
            else{ // 最后一个表格，少3行应该是空出来给“评定”的
                RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 20, (i - 1) * 18 + 24);
            }
        }
        if(record == 1){
            wb.getSheet(sheetname).shiftRows(22, 24, -1); //22-24往上挪了一行
        }
        RowCopy.copyRows(wb, "source", sheetname, 0, 3,(record) * 18 + 3);
        //wb.getSheet("沥青路面压实度左幅").getRow((record) * 18 + 4).getCell(2).setCellValue(lab);
        //wb.getSheet("沥青路面压实度左幅").getRow((record) * 18 + 4).getCell(8).setCellValue(max);
        wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 14, 0,(record) * 18 + 5);

    }

    // 该创建表格函数，需要多次进入，需要维护一个row.模板里面，表格是有表头、填写区域的，但是没有评定
    // 可以自己复制自己，虽然填写信息会复制过来，但是后面再覆盖就好了
    private int createTable(XSSFWorkbook wb,int tableNum,String sheetname, int row) {
        if(tableNum == 0){
            return row;
        }
        int record = 0;
        record = tableNum;

        // 第一次进入，本来就有一个初始表格，但是没有“评定”行，注意，这个鉴定表的评定有3行，钻芯法只有两行
        if(row == 0){
            for (int i = 1; i < record; i++) {
                if(i < record-1){
                    RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 23, (i - 1) * 18 + 24);
                }
                else{ // 最后一段表格，少3行应该是空出来给“评定”的
                    RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 20, (i - 1) * 18 + 24);
                }
            }
            if(record == 1){
                wb.getSheet(sheetname).shiftRows(22, 24, -1); //22-24往上挪了一行
            }
            //
            RowCopy.copyRows(wb, "source", sheetname, 0, 3,(record) * 18 + 3);
            //wb.getSheet("沥青路面压实度左幅").getRow((record) * 18 + 4).getCell(2).setCellValue(lab);
            //wb.getSheet("沥青路面压实度左幅").getRow((record) * 18 + 4).getCell(8).setCellValue(max);
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 14, 0,(record) * 18 + 5); //这个设置打印区域把复制来的第四行给截取了
            return record*18 + 5 + 1; // 最后一行加1，看打印区域为最后一行即可
        }else{
            // 拷贝表头
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 5, row);

            // 不再有第一个表格，所以有多少就复制多少
            for (int i = 1; i <= record; i++) { // 不再有第一个表格了，所以有多少就要复制多少，处理也稍微统一一点，不需要一个23一个22了
                // 在“桥梁左幅”里面复制，是因为这个指标没有桥梁数据，但是刚好模板里面有桥梁空表，因此用来复制
                RowCopy.copyRows(wb, "桥梁左幅", sheetname, 6, 23, (i - 1) * 18 + row+6);//计算上表头
            }

            // 拷贝“评定”
            RowCopy.copyRows(wb, "source", sheetname, 0, 2,(record) * 18 + row + 6);
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 14, 0,(record) * 18 + row + 6 + 2);
            return record*18 + row + 6 + 3;
        }


    }

    /**
     *不分离上面层
     * @param size
     * @return
     */
    private int getZXZtableNum(int size) {
        return size%18 <= 15 ? size/18+1 : size/18+2;
    }

    /*
    * 除了检测点数、合格点数，拿一个规定值就行
    * 修改函数思路：
    *   1. 还是分flag==1和2的情况
    *   2. 把只有一个表的情况放到flag==1的逻辑里面
    *   3. 增加一个字段叫做“分部工程名称”， 一个表的默认为主线
    *       分离：也是只有主线是一个表
    *       不分离：只有主线是一个表。其实实际情况只会有一个表，因为现在把分离和不分离的情况混合起来了
    *       干脆所有的工作表都以有很多表来做，之前的逻辑不能重用了，因为鉴定表中不止一个表，所以在后面的加总再区分flag==1和2的情况
    * */
    @Override
    public List<Map<String, Object>> lookJdbjgbgzbg(CommonInfoVo commonInfoVo, int flag) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object> > jgmap = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "12沥青路面压实度-(不分离上面层).xlsx");
        File f1 = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "12沥青路面压实度-(分离上面层).xlsx");
        if (f.exists()) {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        int thisTableStartRow = 0;
                        for(int i = slSheet.getFirstRowNum(); i <= slSheet.getLastRowNum(); i++){
                            if(slSheet.getRow(i) == null){
                                continue;
                            }

                            if(slSheet.getRow(i).getCell(0) != null){
                                slSheet.getRow(i).getCell(0).setCellType(CellType.STRING);
                                if(slSheet.getRow(i).getCell(0).getStringCellValue().contains("鉴定表")){
                                    thisTableStartRow = i + 6;
                                }else if(slSheet.getRow(i).getCell(0).getStringCellValue().contains("评定")){
                                    // 分部工程名称
                                    String fbgcName = RowCopy.processString(slSheet.getRow(thisTableStartRow).getCell(0).getStringCellValue());

                                    // 检测点数、合格点数
                                    slSheet.getRow(thisTableStartRow).getCell(15).setCellType(CellType.STRING);
                                    slSheet.getRow(thisTableStartRow).getCell(16).setCellType(CellType.STRING);
                                    slSheet.getRow(thisTableStartRow).getCell(17).setCellType(CellType.STRING);
                                    slSheet.getRow(i+1).getCell(2).setCellType(CellType.STRING);//规定值
                                    String gdz1 = slSheet.getRow(i+1).getCell(2).getStringCellValue();
                                    Map map2 = new HashMap();
                                    map2.put("路面类型", wb.getSheetName(j));
                                    map2.put("分部工程名称", fbgcName);
                                    map2.put("密度规定值", gdz1);
                                    map2.put("检测点数", slSheet.getRow(6).getCell(15).getStringCellValue());
                                    map2.put("合格点数", slSheet.getRow(6).getCell(16).getStringCellValue());
                                    map2.put("合格率", slSheet.getRow(6).getCell(17).getStringCellValue());
                                    jgmap.add(map2);
                                }
                            }
                        }
                    }
                }
            }
        }
        if (f1.exists()) {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f1));

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名
                    System.out.println(slSheet.getSheetName());

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        int thisTableStartRow = 0;
                        for(int i = slSheet.getFirstRowNum(); i <= slSheet.getLastRowNum(); i++){
                            if(slSheet.getRow(i) == null){
                                continue;
                            }

                            if(slSheet.getRow(i).getCell(0) != null){
                                slSheet.getRow(i).getCell(0).setCellType(CellType.STRING);
                                if(slSheet.getRow(i).getCell(0).getStringCellValue().contains("鉴定表")){
                                    thisTableStartRow = i + 6;
                                }else if(slSheet.getRow(i).getCell(0).getStringCellValue().contains("评定")){
                                    // 分部工程名称
                                    String fbgcName = RowCopy.processString(slSheet.getRow(thisTableStartRow).getCell(0).getStringCellValue());

                                    // 检测点数、合格点数
                                    slSheet.getRow(thisTableStartRow).getCell(15).setCellType(CellType.STRING);
                                    slSheet.getRow(thisTableStartRow).getCell(16).setCellType(CellType.STRING);
                                    slSheet.getRow(thisTableStartRow).getCell(17).setCellType(CellType.STRING);
                                    slSheet.getRow(i+1).getCell(2).setCellType(CellType.STRING);//规定值
                                    String gdz1 = slSheet.getRow(i+1).getCell(2).getStringCellValue();
                                    Map map2 = new HashMap();
                                    map2.put("路面类型", wb.getSheetName(j));
                                    map2.put("分部工程名称", fbgcName);
                                    map2.put("密度规定值", gdz1);
                                    map2.put("检测点数", slSheet.getRow(6).getCell(15).getStringCellValue());
                                    map2.put("合格点数", slSheet.getRow(6).getCell(16).getStringCellValue());
                                    map2.put("合格率", slSheet.getRow(6).getCell(17).getStringCellValue());
                                    jgmap.add(map2);
                                }
                            }
                        }
                    }
                }
            }
        }

        // 拿到数据以后，根据flag来进行不同的组合
        /*
        * flag == 1:
        *   不分离指标：
        *       1. 主线和匝道可以加在一起了，前提是规定之得相同
        *       2. 隧道也可以加在一起，前提是规定值得相同
        *       3. 连接线也是，规定值得相同
        *
        * flag == 2:
        *   不分离指标：左右幅相加
        *       1. 主线直接加，路面表
        *       2. 匝道直接加，路面表
        *       3. 隧道分部工程相同才加，到自己的表。 连接线隧道不分左右幅
        *       4. 连接线直接加，路面表。连接线不分左右幅
        *    分离指标：左右幅相加，和不分离一样，只是相加的条件多出了同面层才加
        *
        * */
        if(flag == 1){
            processFgmap(jgmap);
            secondMerge(jgmap);
        }else if(flag == 2){

            processFgmap(jgmap);

        }
        sortByOrder(jgmap);
        return jgmap;
    }

    // 分离指标判断函数
    private boolean isSeparated(Map<String, Object> m) {
        return m.get("路面类型").toString().contains("-");
    }
    public void processFgmap(List<Map<String, Object>> fgmap) {



        // 处理不分离指标和分离指标
        List<Map<String, Object>> merged = new ArrayList<>(
                Stream.concat(
                        // 处理不分离指标
                        fgmap.stream()
                                .filter(m -> !isSeparated(m))
                                .collect(Collectors.groupingBy(m -> {
                                    String roadType = m.get("路面类型").toString();
                                    String base = roadType
                                            .replaceAll("左幅|右幅", "")
                                            .replaceAll("匝道", "匝道");
                                    return "隧道".equals(base) ? base + "|" + m.get("分部工程名称") : base;
                                }))
                                .values()
                                .stream()
                                .map(group -> mergeMaps(group)),

                        // 处理分离指标
                        fgmap.stream()
                                .filter(m -> isSeparated(m))
                                .collect(Collectors.groupingBy(m -> {
                                    String roadType = m.get("路面类型").toString();
                                    String[] parts = roadType.split("-");
                                    String base = parts[0]
                                            .replaceAll("左幅|右幅", "")
                                            .replaceAll("匝道", "匝道");
                                    String layer = parts[1];
                                    String key = base + "-" + layer;
                                    return "隧道".equals(base) ? key + "|" + m.get("分部工程名称") : key;
                                }))
                                .values()
                                .stream()
                                .map(group -> mergeMaps(group))
                ).collect(Collectors.toList())
        );

        fgmap.clear();
        fgmap.addAll(merged);
    }

    private Map<String, Object> mergeMaps(List<Map<String, Object>> group) {
        Map<String, Object> merged = new HashMap<>(group.get(0));

        int totalPoints = group.stream()
                .mapToInt(m ->  (int) Double.parseDouble(m.get("检测点数").toString()))
                .sum();

        int qualifiedPoints = group.stream()
                .mapToInt(m ->  (int) Double.parseDouble(m.get("合格点数").toString()))
                .sum();

        double qualifiedRate = qualifiedPoints * 100.0 / totalPoints;

        merged.put("检测点数", String.valueOf(totalPoints));
        merged.put("合格点数", String.valueOf(qualifiedPoints));
        merged.put("合格率", String.format("%.2f", qualifiedRate));

        // 处理路面类型名称 - 仅去除"左幅"和"右幅"
        String originalType = merged.get("路面类型").toString();
        String mergedType = originalType.replaceAll("左幅|右幅", "");

        merged.put("路面类型", mergedType);
        return merged;
    }

    // 第二次合并，即flag==1的情况
    public void secondMerge(List<Map<String, Object>> fgmap) {

        // 核心合并逻辑
        List<Map<String, Object>> merged = Stream.concat(
                // 处理不分离指标（沥青和隧道）
                fgmap.stream()
                        .filter(m -> !isSeparated(m))
                        .collect(Collectors.groupingBy(m -> {
                            String type = m.get("路面类型").toString();
                            String density = m.get("密度规定值").toString();

                            if (type.startsWith("沥青路面压实度")) {
                                return "沥青类#" + density; // 沥青和匝道合并
                            }
                            if (type.startsWith("隧道")) {
                                return "隧道#" + density;   // 隧道只需密度相同
                            }
                            return type + "#" + density;
                        }))
                        .values().stream().map(this::mergeGroup),

                // 处理分离指标（带面层的）
                fgmap.stream()
                        .filter(this::isSeparated)
                        .collect(Collectors.groupingBy(m -> {
                            String[] parts = m.get("路面类型").toString().split("-");
                            String base = parts[0];
                            String layer = parts[1];
                            String density = m.get("密度规定值").toString();

                            if (base.startsWith("沥青路面压实度")) {
                                return "沥青类#" + layer + "#" + density;
                            }
                            if (base.startsWith("隧道")) {
                                return "隧道#" + layer + "#" + density; // 隧道只看面层+密度
                            }
                            return base + "#" + layer + "#" + density;
                        }))
                        .values().stream().map(this::mergeGroup)
        ).collect(Collectors.toList());

        fgmap.clear();
        fgmap.addAll(merged);
    }

    // 合并组内元素
    private Map<String, Object> mergeGroup(List<Map<String, Object>> group) {
        Map<String, Object> base = new HashMap<>(group.get(0));

        // 数值累加
        int total = group.stream().mapToInt(m -> (int) Double.parseDouble(m.get("检测点数").toString())).sum();
        int qualified = group.stream().mapToInt(m -> (int) Double.parseDouble(m.get("合格点数").toString())).sum();
        double rate = qualified * 100.0 / total;

        // 更新关键字段
        base.put("检测点数", String.valueOf(total));
        base.put("合格点数", String.valueOf(qualified));
        base.put("合格率", String.format("%.2f", rate));

        // 统一类型名称
        String type = base.get("路面类型").toString();
        if (type.startsWith("沥青路面压实度")) {
            base.put("路面类型", type.replaceAll("左幅|右幅", ""));
        }
        else if (type.startsWith("隧道")) {
            if (type.contains("-")) {
                String[] parts = type.split("-");
                base.put("路面类型", "隧道" + "-" + parts[1]); // 保留面层信息
            } else {
                base.put("路面类型", "隧道"); // 统一非分离指标名称
            }
        }

        return base;
    }

    // 排序函数
    public static void sortByOrder(List<Map<String, Object>> fgmap) {
        List<String> order = Arrays.asList(
                "沥青路面压实度",
                "沥青路面压实度匝道",
                "隧道",
                "连接线",
                "沥青路面压实度-上面层",
                "沥青路面压实度-中下面层",
                "沥青路面压实度匝道-上面层",
                "沥青路面压实度匝道-中下面层",
                "隧道-上面层",
                "隧道-中下面层",
                "连接线-上面层",
                "连接线-中下面层"
        );

        fgmap.sort(Comparator.comparingInt(m -> {
            String type = m.get("路面类型").toString();
            int index = order.indexOf(type);
            return index != -1 ? index : Integer.MAX_VALUE; // 未匹配的排在最后
        }));
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

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object> > jgmap = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "12沥青路面压实度-(不分离上面层).xlsx");
        File f1 = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "12沥青路面压实度-(分离上面层).xlsx");
        if (f.exists()) {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        //标准密度
                        slSheet.getRow(lastRowNum - 2).getCell(6).setCellType(CellType.STRING);//检测点数
                        slSheet.getRow(lastRowNum - 1).getCell(6).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum - 1).getCell(2).setCellType(CellType.STRING);//规定值
                        slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//代表值
                        slSheet.getRow(lastRowNum - 2).getCell(16).setCellType(CellType.STRING);//最大值
                        slSheet.getRow(lastRowNum - 1).getCell(16).setCellType(CellType.STRING);//最小值

                        String max1 = slSheet.getRow(lastRowNum - 2).getCell(16).getStringCellValue();
                        String min1 = slSheet.getRow(lastRowNum - 1).getCell(16).getStringCellValue();
                        String dbz1 = slSheet.getRow(lastRowNum).getCell(4).getStringCellValue();
                        String gdz1 = slSheet.getRow(lastRowNum - 1).getCell(2).getStringCellValue();
                        String jcds1 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(6).getStringCellValue()));
                        String hgds1 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 1).getCell(6).getStringCellValue()));
                        String hgl1 = df.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue()));
                        //最大理论密度
                        slSheet.getRow(lastRowNum - 2).getCell(12).setCellType(CellType.STRING);//检测点数
                        slSheet.getRow(lastRowNum - 1).getCell(12).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(12).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum - 1).getCell(8).setCellType(CellType.STRING);//规定值
                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);//代表值
                        slSheet.getRow(lastRowNum - 2).getCell(17).setCellType(CellType.STRING);//最大值
                        slSheet.getRow(lastRowNum - 1).getCell(17).setCellType(CellType.STRING);//最小值

                        String max2 = slSheet.getRow(lastRowNum - 2).getCell(17).getStringCellValue();
                        String min2 = slSheet.getRow(lastRowNum - 1).getCell(17).getStringCellValue();
                        String dbz2 = slSheet.getRow(lastRowNum).getCell(10).getStringCellValue();
                        String gdz2 = slSheet.getRow(lastRowNum - 1).getCell(8).getStringCellValue();
                        String jcds2 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(12).getStringCellValue()));
                        String hgds2 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 1).getCell(12).getStringCellValue()));
                        String hgl2 = df.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(12).getStringCellValue()));

                        Map map1 = new HashMap();
                        map1.put("路面类型", wb.getSheetName(j));
                        map1.put("标准", "标准密度");
                        map1.put("实测值变化范围", min1 + "~" + max1);
                        map1.put("密度代表值", dbz1);
                        map1.put("密度规定值", gdz1);
                        map1.put("检测点数", jcds1);
                        map1.put("合格点数", hgds1);
                        map1.put("合格率", hgl1);
                        jgmap.add(map1);

                        Map map2 = new HashMap();
                        map2.put("路面类型", wb.getSheetName(j));
                        map2.put("标准", "最大理论密度");
                        map2.put("实测值变化范围", min2 + "~" + max2);
                        map2.put("密度代表值", dbz2);
                        map2.put("密度规定值", gdz2);
                        map2.put("检测点数", jcds2);
                        map2.put("合格点数", hgds2);
                        map2.put("合格率", hgl2);
                        jgmap.add(map2);
                    }
                }
            }
        }
        if (f1.exists()) {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f1));

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名
                    System.out.println(slSheet.getSheetName());

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        //标准密度
                        slSheet.getRow(lastRowNum - 2).getCell(6).setCellType(CellType.STRING);//检测点数
                        slSheet.getRow(lastRowNum - 1).getCell(6).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum - 1).getCell(2).setCellType(CellType.STRING);//规定值
                        slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//代表值
                        slSheet.getRow(lastRowNum - 2).getCell(16).setCellType(CellType.STRING);//最大值
                        slSheet.getRow(lastRowNum - 1).getCell(16).setCellType(CellType.STRING);//最小值

                        String max1 = slSheet.getRow(lastRowNum - 2).getCell(16).getStringCellValue();
                        String min1 = slSheet.getRow(lastRowNum - 1).getCell(16).getStringCellValue();
                        String dbz1 = slSheet.getRow(lastRowNum).getCell(4).getStringCellValue();
                        String gdz1 = slSheet.getRow(lastRowNum - 1).getCell(2).getStringCellValue();
                        String jcds1 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(6).getStringCellValue()));
                        String hgds1 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 1).getCell(6).getStringCellValue()));
                        String hgl1 = df.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue()));
                        //最大理论密度
                        slSheet.getRow(lastRowNum - 2).getCell(12).setCellType(CellType.STRING);//检测点数
                        slSheet.getRow(lastRowNum - 1).getCell(12).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(12).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum - 1).getCell(8).setCellType(CellType.STRING);//规定值
                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);//代表值
                        slSheet.getRow(lastRowNum - 2).getCell(17).setCellType(CellType.STRING);//最大值
                        slSheet.getRow(lastRowNum - 1).getCell(17).setCellType(CellType.STRING);//最小值

                        String max2 = slSheet.getRow(lastRowNum - 2).getCell(17).getStringCellValue();
                        String min2 = slSheet.getRow(lastRowNum - 1).getCell(17).getStringCellValue();
                        String dbz2 = slSheet.getRow(lastRowNum).getCell(10).getStringCellValue();
                        String gdz2 = slSheet.getRow(lastRowNum - 1).getCell(8).getStringCellValue();
                        String jcds2 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(12).getStringCellValue()));
                        String hgds2 = decf.format(Double.valueOf(slSheet.getRow(lastRowNum - 1).getCell(12).getStringCellValue()));
                        String hgl2 = df.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(12).getStringCellValue()));

                        Map map1 = new HashMap();
                        map1.put("路面类型", wb.getSheetName(j));
                        map1.put("标准", "标准密度");
                        map1.put("实测值变化范围", min1 + "~" + max1);
                        map1.put("密度代表值", dbz1);
                        map1.put("密度规定值", gdz1);
                        map1.put("检测点数", jcds1);
                        map1.put("合格点数", hgds1);
                        map1.put("合格率", hgl1);
                        jgmap.add(map1);

                        Map map2 = new HashMap();
                        map2.put("路面类型", wb.getSheetName(j));
                        map2.put("标准", "最大理论密度");
                        map2.put("实测值变化范围", min2 + "~" + max2);
                        map2.put("密度代表值", dbz2);
                        map2.put("密度规定值", gdz2);
                        map2.put("检测点数", jcds2);
                        map2.put("合格点数", hgds2);
                        map2.put("合格率", hgl2);
                        jgmap.add(map2);
                    }
                }
            }
        }
        return jgmap;

    }

    /**
     * 导出模板文件
     * @param response
     */
    @Override
    public void exportlqlmysd(HttpServletResponse response) {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            String fileName = "01沥青路面压实度实测数据";
            String sheetName = "实测数据";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLqlmysdVo()).finish();
            courseFile = directory.getCanonicalPath();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"01沥青路面压实度实测数据.xlsx");
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
    public void importlqlmysd(MultipartFile file, CommonInfoVo commonInfoVo) {
        String htd = commonInfoVo.getHtd();
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLqlmysdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLqlmysdVo>(JjgFbgcLmgcLqlmysdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLqlmysdVo> dataList) {
                                    int size = dataList.size();
                                    if (size % 3 != 0){
                                        throw new JjgysException(20001,"您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，数据量错误，请确保每个桩号有上、中、下三个面层的数据");
                                    }
                                    int rowNumber=2;
                                    for(JjgFbgcLmgcLqlmysdVo lqlmysdVo: dataList)
                                    {
                                        if (StringUtils.isEmpty(lqlmysdVo.getZh())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，第"+rowNumber+"行的数据中，桩号为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lqlmysdVo.getLmlx())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，第"+rowNumber+"行的数据中，路面类型为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lqlmysdVo.getLqs())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，第"+rowNumber+"行的数据中，路桥隧为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lqlmysdVo.getQywz())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，第"+rowNumber+"行的数据中，取样位置为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lqlmysdVo.getCw())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，第"+rowNumber+"行的数据中，层位为空，请修改后重新上传");
                                        }
                                        if (StringUtils.isEmpty(lqlmysdVo.getSffl())) {
                                            throw new JjgysException(20001, "您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，第"+rowNumber+"行的数据中，是否分离为空，请修改后重新上传");
                                        }

                                        JjgFbgcLmgcLqlmysd fbgcLmgcLqlmysd = new JjgFbgcLmgcLqlmysd();
                                        BeanUtils.copyProperties(lqlmysdVo,fbgcLmgcLqlmysd);
                                        fbgcLmgcLqlmysd.setCreatetime(new Date());
                                        fbgcLmgcLqlmysd.setUserid(commonInfoVo.getUserid());
                                        fbgcLmgcLqlmysd.setProname(commonInfoVo.getProname());
                                        fbgcLmgcLqlmysd.setHtd(commonInfoVo.getHtd());
                                        fbgcLmgcLqlmysd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLqlmysdMapper.insert(fbgcLmgcLqlmysd);
                                        rowNumber++;
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，解析excel出错，请传入正确格式的excel");
        }catch (NullPointerException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，请检查数据的正确性或删除文件最后的空数据，然后重新上传");
        }catch (ExcelAnalysisException e) {
            throw new JjgysException(20001,"您上传的"+htd+"合同段中01沥青路面压实度实测数据.xlsx文件，请将检测日期修改为2021/01/01的格式，然后重新上传");
        }


    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLmgcLqlmysdMapper.selectnum(proname, htd);
        return selectnum;
    }

    @Override
    public int selectnumname(String proname) {
        int selectnum = jjgFbgcLmgcLqlmysdMapper.selectnumname(proname);
        return selectnum;
    }

    @Override
    public List<Map<String, Object>> selectsfl(String proname, String htd, List<Long> idlist) {
        List<Map<String, Object>> s = jjgFbgcLmgcLqlmysdMapper.selectsfl(proname,htd,idlist);
        return s;
    }



    @Override
    public int createMoreRecords(List<JjgFbgcLmgcLqlmysd> data, String userID) {
        return ReceiveUtils.createMore(data, jjgFbgcLmgcLqlmysdMapper, userID);
    }

}
