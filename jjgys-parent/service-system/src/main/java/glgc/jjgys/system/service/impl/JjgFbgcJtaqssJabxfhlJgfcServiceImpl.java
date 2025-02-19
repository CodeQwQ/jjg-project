package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhlJgfc;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabxfhlJgfcVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcJtaqssJabxfhlJgfcMapper;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxfhlJgfcService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
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
 * @since 2024-05-20
 */
@Service
public class JjgFbgcJtaqssJabxfhlJgfcServiceImpl extends ServiceImpl<JjgFbgcJtaqssJabxfhlJgfcMapper, JjgFbgcJtaqssJabxfhlJgfc> implements JjgFbgcJtaqssJabxfhlJgfcService {

    @Autowired
    private JjgFbgcJtaqssJabxfhlJgfcMapper jjgFbgcJtaqssJabxfhlJgfcMapper;

    @Value(value = "${jjgys.path.jgfilepath}")
    private String jgfilepath;

    @Override
    public boolean generateJdb(CommonInfoVo commonInfoVo) {
        String proname = commonInfoVo.getProname();
        List<String> htds = jjgFbgcJtaqssJabxfhlJgfcMapper.gethtd(proname);
        if (htds.size() > 0) {
            for (String htd : htds) {
                gethtdjdb(proname, htd);
            }
            return true;
        } else {
            return false;
        }

    }

    /**
     * @param proname
     * @param htd
     */
    private void gethtdjdb(String proname, String htd) {
        XSSFWorkbook wb = null;
        //获取数据
        QueryWrapper<JjgFbgcJtaqssJabxfhlJgfc> wrapper = new QueryWrapper<>();
        wrapper.like("proname", proname);
        wrapper.like("htd", htd);
        wrapper.orderByAsc("wzjlx");
        List<JjgFbgcJtaqssJabxfhlJgfc> data = jjgFbgcJtaqssJabxfhlJgfcMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "58交安钢防护栏.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (data == null || data.size() == 0) {
            return;
        } else {
            //存放鉴定表的目录
            File fdir = new File(jgfilepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            try {
                /*File directory = new File("service-system/src/main/resources/static");
                String reportPath = directory.getCanonicalPath();
                String name = "钢防护栏.xlsx";
                String path = reportPath + File.separator + name;
                Files.copy(Paths.get(path), new FileOutputStream(f));*/
                InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/钢防护栏.xlsx");
                Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);
                createTable(gettableNum(data.size()), wb);
                if (DBtoExcel(data, wb)) {
                    for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }
                    FileOutputStream fileOut = new FileOutputStream(f);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();
                }
                out.close();
                wb.close();
                inputStream.close();
            } catch (Exception e) {
                if (f.exists()) {
                    f.delete();
                }
                throw new JjgysException(20001, "生成鉴定表错误，请检查数据的正确性");
            }
        }

    }

    /**
     * 写入数据
     *
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcJtaqssJabxfhlJgfc> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("防护栏");
        sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
        sheet.getRow(1).getCell(10).setCellValue(data.get(0).getHtd());

        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for (int i = 1; i < data.size(); i++) {
            Date jcsj = data.get(i).getJcsj();
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(jcsj));
        }
        sheet.getRow(3).getCell(10).setCellValue(date);

        String error[] = getError(data);

        sheet.getRow(5).getCell(2).setCellValue(error[0]);
        sheet.getRow(5).getCell(5).setCellValue(error[1]);
        sheet.getRow(5).getCell(8).setCellValue(error[2]);
        sheet.getRow(5).getCell(11).setCellValue(error[3]);

        int index = 0;
        for (int i = 0; i < data.size(); i++) {
            sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 11, 0, 0));  //位置
            sheet.getRow(index + 7).getCell(0).setCellValue(data.get(i).getWzjlx());

            if (data.get(i).getMrsdscz() != null && !"".equals(data.get(i).getMrsdscz())) {
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 4, 11, 11));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 4, 12, 12));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 4, 13, 13));

                sheet.getRow(index + 7).getCell(11).setCellValue(Double.valueOf(data.get(i).getMrsdscz()));

                sheet.getRow(index + 7).getCell(12).setCellFormula("IF(AND("
                        + sheet.getRow(index + 7).getCell(11).getReference() + "-"
                        + data.get(i).getMrsdgdz() + ">=0),\"√\",\"\")");

                sheet.getRow(index + 7).getCell(13).setCellFormula("IF(AND("
                        + sheet.getRow(index + 7).getCell(11).getReference() + "-"
                        + data.get(i).getMrsdgdz() + ">=0),\"\",\"×\")");

            } else {
                sheet.getRow(index + 7).getCell(11).setCellValue("-");
                sheet.getRow(index + 7).getCell(12).setCellValue("-");
                sheet.getRow(index + 7).getCell(13).setCellValue("-");
            }

            sheet.getRow(index + 7).getCell(1).setCellValue(Double.valueOf((1)));  //序号
            sheet.getRow(index + 8).getCell(1).setCellValue(Double.valueOf((2)));  //序号
            sheet.getRow(index + 9).getCell(1).setCellValue(Double.valueOf((3)));  //序号
            sheet.getRow(index + 10).getCell(1).setCellValue(Double.valueOf((4)));  //序号
            sheet.getRow(index + 11).getCell(1).setCellValue(Double.valueOf((5)));  //序号

            //波形梁板基底金属厚度（mm）
            sheet.getRow(index + 7).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz1()));  //实测值

            sheet.getRow(index + 7).getCell(3).setCellFormula("IF(AND("
                    + sheet.getRow(index + 7).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 7).getCell(4).setCellFormula("IF(AND("
                    + sheet.getRow(index + 7).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 8).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz2()));  //实测值

            sheet.getRow(index + 8).getCell(3).setCellFormula("IF(AND("
                    + sheet.getRow(index + 8).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 8).getCell(4).setCellFormula("IF(AND("
                    + sheet.getRow(index + 8).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 9).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz3()));  //实测值
            sheet.getRow(index + 9).getCell(3).setCellFormula("IF(AND("
                    + sheet.getRow(index + 9).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 9).getCell(4).setCellFormula("IF(AND("
                    + sheet.getRow(index + 9).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 10).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz4()));  //实测值

            sheet.getRow(index + 10).getCell(3).setCellFormula("IF(AND("
                    + sheet.getRow(index + 10).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 10).getCell(4).setCellFormula("IF(AND("
                    + sheet.getRow(index + 10).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 11).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz5()));  //实测值
            sheet.getRow(index + 11).getCell(3).setCellFormula("IF(AND("
                    + sheet.getRow(index + 11).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 11).getCell(4).setCellFormula("IF(AND("
                    + sheet.getRow(index + 11).getCell(2).getReference() + "-"
                    + data.get(i).getJdhdgdz() + ">=0),\"\",\"×\")");

            //波形梁钢护栏立柱
            sheet.getRow(index + 7).getCell(5).setCellValue(Double.valueOf(data.get(i).getLzbhscz1()));
            sheet.getRow(index + 7).getCell(6).setCellFormula("IF(AND("
                    + sheet.getRow(index + 7).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 7).getCell(7).setCellFormula("IF(AND("
                    + sheet.getRow(index + 7).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 8).getCell(5).setCellValue(Double.valueOf(data.get(i).getLzbhscz2()));
            sheet.getRow(index + 8).getCell(6).setCellFormula("IF(AND("
                    + sheet.getRow(index + 8).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 8).getCell(7).setCellFormula("IF(AND("
                    + sheet.getRow(index + 8).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 9).getCell(5).setCellValue(Double.valueOf(data.get(i).getLzbhscz3()));
            sheet.getRow(index + 9).getCell(6).setCellFormula("IF(AND("
                    + sheet.getRow(index + 9).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 9).getCell(7).setCellFormula("IF(AND("
                    + sheet.getRow(index + 9).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 10).getCell(5).setCellValue(Double.valueOf(data.get(i).getLzbhscz4()));
            sheet.getRow(index + 10).getCell(6).setCellFormula("IF(AND("
                    + sheet.getRow(index + 10).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 10).getCell(7).setCellFormula("IF(AND("
                    + sheet.getRow(index + 10).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"\",\"×\")");

            sheet.getRow(index + 11).getCell(5).setCellValue(Double.valueOf(data.get(i).getLzbhscz5()));
            sheet.getRow(index + 11).getCell(6).setCellFormula("IF(AND("
                    + sheet.getRow(index + 11).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"√\",\"\")");
            sheet.getRow(index + 11).getCell(7).setCellFormula("IF(AND("
                    + sheet.getRow(index + 11).getCell(5).getReference() + "-"
                    + data.get(i).getLzbhgdz() + ">=0),\"\",\"×\")");

            //波形梁钢护栏横梁中心高度（mm）
            sheet.getRow(index + 7).getCell(8).setCellValue(Double.valueOf(data.get(i).getZxgdscz1()));

            sheet.getRow(index + 7).getCell(9).setCellFormula("IF(AND("
                    + sheet.getRow(index + 7).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 7).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"√\",\"\")");

            sheet.getRow(index + 7).getCell(10).setCellFormula("IF(AND("
                    + sheet.getRow(index + 7).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 7).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"\",\"×\")");


            sheet.getRow(index + 8).getCell(8).setCellValue(Double.valueOf(data.get(i).getZxgdscz2()));
            sheet.getRow(index + 8).getCell(9).setCellFormula("IF(AND("
                    + sheet.getRow(index + 8).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 8).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"√\",\"\")");

            sheet.getRow(index + 8).getCell(10).setCellFormula("IF(AND("
                    + sheet.getRow(index + 8).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 8).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"\",\"×\")");
            sheet.getRow(index + 9).getCell(8).setCellValue(Double.valueOf(data.get(i).getZxgdscz3()));
            sheet.getRow(index + 9).getCell(9).setCellFormula("IF(AND("
                    + sheet.getRow(index + 9).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 9).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"√\",\"\")");

            sheet.getRow(index + 9).getCell(10).setCellFormula("IF(AND("
                    + sheet.getRow(index + 9).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 9).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"\",\"×\")");
            sheet.getRow(index + 10).getCell(8).setCellValue(Double.valueOf(data.get(i).getZxgdscz4()));
            sheet.getRow(index + 10).getCell(9).setCellFormula("IF(AND("
                    + sheet.getRow(index + 10).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 10).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"√\",\"\")");

            sheet.getRow(index + 10).getCell(10).setCellFormula("IF(AND("
                    + sheet.getRow(index + 10).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 10).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"\",\"×\")");
            sheet.getRow(index + 11).getCell(8).setCellValue(Double.valueOf(data.get(i).getZxgdscz5()));
            sheet.getRow(index + 11).getCell(9).setCellFormula("IF(AND("
                    + sheet.getRow(index + 11).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 11).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"√\",\"\")");

            sheet.getRow(index + 11).getCell(10).setCellFormula("IF(AND("
                    + sheet.getRow(index + 11).getCell(8).getReference() + "-"
                    + data.get(i).getZxgdgdz() + "<=" + data.get(i).getZxgdyxpsz() + ","
                    + sheet.getRow(index + 11).getCell(8).getReference() + "-" + data.get(i).getZxgdgdz() + ">=-" + data.get(i).getZxgdyxpsf() + "),\"\",\"×\")");

            index += 5;

            /*
             * 所有数据填写完毕，此处计算合计
             */
            //波形梁板基底金属
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(3).setCellFormula("COUNT(C8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(2).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(3).setCellFormula("COUNTIF(D8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(3).getReference() + ",\"√\")");//=COUNTIF(E8:E302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(3).setCellFormula("COUNTIF(E8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(4).getReference() + ",\"×\")");//=COUNTIF(F8:F302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(3).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(3).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(3).getReference() + "*100,2)");//=D304/D303*100
            //波形梁钢护栏立柱
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(5).setCellFormula("COUNT(F8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(5).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(5).setCellFormula("COUNTIF(G8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(6).getReference() + ",\"√\")");//=COUNTIF(I8:I302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(5).setCellFormula("COUNTIF(H8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(7).getReference() + ",\"×\")");//=COUNTIF(J8:J302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(5).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(5).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(5).getReference() + "*100,2)");//=D304/D303*100
            //波形梁钢护栏横梁中心
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(8).setCellFormula("COUNT(I8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(8).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(8).setCellFormula("COUNTIF(J8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(9).getReference() + ",\"√\")");//=COUNTIF(L8:L302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(8).setCellFormula("COUNTIF(K8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(10).getReference() + ",\"×\")");//=COUNTIF(M8:M302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(8).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(8).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(8).getReference() + "*100,2)");//=D304/D303*100
            //波形梁钢护栏立柱埋入深度（mm）
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(11).setCellFormula("COUNT(L8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(11).getReference() + ")");

            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(11).setCellFormula("COUNTIF(M8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(12).getReference() + ",\"√\")");//=COUNTIF(O8:O302,"√")
           /* if (sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(15) == null){
                sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).createCell(15);
            }*/
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(11).setCellFormula("COUNTIF(N8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(13).getReference() + ",\"×\")");//=COUNTIF(P8:P302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(11).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(11).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(11).getReference() + "*100,2)");//=D304/D303*100

        }
        return true;
    }

    /**
     * 获取偏差值
     *
     * @param data
     * @return
     */
    private String[] getError(List<JjgFbgcJtaqssJabxfhlJgfc> data) {
        String[] error = {"", "", "", ""};
        String temp = "";
        ArrayList<String> errorlist = new ArrayList<String>();
        for (int i = 0; i < data.size(); i++) {
            if (!errorlist.contains(data.get(i).getWzjlx())) {//  按位置 存偏差
                errorlist.add((data.get(i).getWzjlx()));
                if (!"".equals(data.get(i).getJdhdgdz()) && !data.get(i).getJdhdgdz().equals("0")) {
                    /*if(data.get(i).getJdhdgdz().contains("≥")){
                        temp+=data.get(i).getJdhdgdz();
                    }*/
                    temp += data.get(i).getJdhdgdz();
                    if (!error[0].contains(temp)) {
                        error[0] += "≥" + temp;
                    }
                    temp = "";
                }
                if (data.get(i).getWzjlx().contains("圆柱")) {  //位置中有圆柱
                    if (!"".equals(data.get(i).getLzbhgdz()) && !data.get(i).getLzbhgdz().equals("0")) {
                        if (data.get(i).getLzbhgdz().contains("圆柱≥")) {
                            temp += data.get(i).getLzbhgdz();
                        } else if (data.get(i).getLzbhgdz().contains("≥")) {
                            temp += "圆柱" + data.get(i).getLzbhgdz();
                        } else {
                            temp += "圆柱≥" + data.get(i).getLzbhgdz();
                        }
                    }
                } else if (data.get(i).getWzjlx().contains("方柱")) {
                    if (!"".equals(data.get(i).getLzbhgdz()) && !data.get(i).getLzbhgdz().equals("0")) {
                        temp += "方柱≥" + data.get(i).getLzbhgdz();
                    }
                }
                if (!error[1].contains(temp)) {
                    error[1] += temp;
                }
                temp = "";
                if (data.get(i).getWzjlx().contains("两波") || data.get(i).getWzjlx().contains("双层双波形")) {
                    if (!"".equals(data.get(i).getZxgdgdz()) && !data.get(i).getZxgdgdz().equals("0")) {
                        temp += "两波板" + Double.valueOf(data.get(i).getZxgdgdz()).intValue();
                        if (!data.get(i).getZxgdyxpsz().equals("0") && !data.get(i).getZxgdyxpsf().equals("0")) {
                            if (data.get(i).getZxgdyxpsz().equals(data.get(i).getZxgdyxpsf())) {
                                temp += "±" + Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                        } else {
                            if (!data.get(i).getZxgdyxpsz().equals("0")) {
                                temp += "+" + Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                            if (!data.get(i).getZxgdyxpsf().equals("0")) {
                                temp += "-" + Double.valueOf(data.get(i).getZxgdyxpsf()).intValue();
                            }
                        }
                    }
                } else if (data.get(i).getWzjlx().contains("三波")) {
                    if (!"".equals(data.get(i).getZxgdgdz()) && !data.get(i).getZxgdgdz().equals("0")) {
                        temp += "三波板" + Double.valueOf(data.get(i).getZxgdgdz()).intValue();
                        if (!data.get(i).getZxgdyxpsz().equals("0") && !data.get(i).getZxgdyxpsf().equals("0")) {
                            if (data.get(i).getZxgdyxpsz().equals(data.get(i).getZxgdyxpsf())) {
                                temp += "±" + Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                        } else {
                            if (!data.get(i).getZxgdyxpsz().equals("0")) {
                                temp += "+" + Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                            if (!data.get(i).getZxgdyxpsf().equals("0")) {
                                temp += "-" + Double.valueOf(data.get(i).getZxgdyxpsf()).intValue();
                            }
                        }
                    }
                }
                if (!error[2].contains(temp)) {
                    error[2] += temp;
                }
                temp = "";


                if (data.get(i).getWzjlx().contains("圆柱")) {
                    if (!"".equals(data.get(i).getMrsdgdz()) && !data.get(i).getMrsdgdz().equals("0")) {
                        if (data.get(i).getMrsdgdz().contains("≥")) {
                            temp += "圆柱" + data.get(i).getMrsdgdz();
                        } else {
                            temp += "圆柱≥" + data.get(i).getMrsdgdz();
                        }
                    }
                } else if (data.get(i).getWzjlx().contains("方柱")) {
                    if (!"".equals(data.get(i).getMrsdgdz()) && !data.get(i).getMrsdgdz().equals("0")) {
                        if (data.get(i).getMrsdgdz().contains("≥")) {
                            temp += "方柱" + data.get(i).getMrsdgdz();
                        } else {
                            temp += "方柱≥" + data.get(i).getMrsdgdz();
                        }
                    }
                }
                if (!error[3].contains(temp)) {
                    error[3] += temp;
                }
                temp = "";
            }
        }
        return error;
    }

    /**
     * 创建模板页
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;

        for (int i = 1; i < record; i++) {
            if (i < record - 1) {
                //RowCopy.copyRows(wb, "防护栏", "防护栏", 7, 31, (i - 1) * 25 + 32);
                RowCopy.copyRows(wb, "防护栏", "防护栏", 7, 41, (i - 1) * 35 + 42);
            } else {
                RowCopy.copyRows(wb, "防护栏", "防护栏", 7, 38, (i - 1) * 35 + 42);
            }
        }
        if (record >= 1)
            RowCopy.copyRows(wb, "source", "防护栏", 0, 4, (record - 1) * 35 + 38);


        wb.setPrintArea(wb.getSheetIndex("防护栏"), 0, 13, 0, record * 35 + 6);

    }

    /**
     * 模板页数
     *
     * @param tableNum
     * @return
     */
    private int gettableNum(int tableNum) {
        return tableNum % 5 == 0 || tableNum % 5 == 7 ? tableNum / 5 + 2 : tableNum / 5 + 1;
    }

    @Override
    public void importjabxfhl(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        // 将文件流传过来，变成workbook对象
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        //获得文本
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获得行数
        int rows = sheet.getPhysicalNumberOfRows();
        //获得列数
        int columns = 0;
        for (int i = 1; i < rows; i++) {
            XSSFRow row = sheet.getRow(i);
            columns = row.getPhysicalNumberOfCells();
        }
        JjgFbgcJtaqssJabxfhlJgfcVo jjgFbgcJtaqssJabxfhlVo = new JjgFbgcJtaqssJabxfhlJgfcVo();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        List<String> titlelist = new ArrayList();
        for (int n = 0; n < 1; n++) {
            for (int m = 0; m < rows; m++) {
                XSSFRow row = sheet.getRow(m);
                XSSFCell cell = row.getCell(n);
                titlelist.add(cell.toString());
            }
        }
        List<String> checklist = Arrays.asList("合同段名称","检测日期", "位置及类型", "基底厚度规定值(mm)", "基底厚度实测值1(mm)",
                "基底厚度实测值2(mm)", "基底厚度实测值3(mm)", "基底厚度实测值4(mm)", "基底厚度实测值5(mm)", "立柱壁厚规定值(mm)", "立柱壁厚实测值1(mm)", "立柱壁厚实测值2(mm)",
                "立柱壁厚实测值3(mm)", "立柱壁厚实测值4(mm)", "立柱壁厚实测值5(mm)", "中心高度规定值(mm)", "中心高度允许偏差+(mm)",
                "中心高度允许偏差-(mm)", "中心高度实测值1(mm)", "中心高度实测值2(mm)", "中心高度实测值3(mm)", "中心高度实测值4(mm)", "中心高度实测值5(mm)", "埋入深度规定值(mm)", "埋入深度实测值(mm)");
        if (checklist.equals(titlelist)) {
            for (int j = 1; j < columns; j++) {//列
                Map<String, Object> map = new HashMap<>();
                Field[] fields = jjgFbgcJtaqssJabxfhlVo.getClass().getDeclaredFields();
                JjgFbgcJtaqssJabxfhlJgfc jjgFbgcJtaqssJabxfhl = new JjgFbgcJtaqssJabxfhlJgfc();
                for (int k = 0; k < rows; k++) {//行
                    //列是不变的 行增加
                    XSSFRow row = sheet.getRow(k);
                    XSSFCell cell = row.getCell(j);

                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING://String
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(), cell.getStringCellValue());//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN://bealean
                            //cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                            cell.setCellType(CellType.STRING);
                            //map.put(fields[k].getName(),Boolean.valueOf(cell.getBooleanCellValue()).toString());//属性赋值
                            map.put(fields[k].getName(), String.valueOf(cell.getStringCellValue()));//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC://number
                            //默认日期读取出来是数字，判断是否是日期格式的数字
                            if (DateUtil.isCellDateFormatted(cell)) {
                                //读取的数字是日期，转换一下格式
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                Date date = cell.getDateCellValue();
                                System.out.println(date);
                                map.put(fields[k].getName(), dateFormat.format(date));//属性赋值
                            } else {//不是日期直接赋值
                                System.out.println(cell);

                                //map.put(fields[k].getName(),Double.valueOf(cell.getNumericCellValue()).toString());//属性赋值
                                map.put(fields[k].getName(), String.valueOf(cell.getNumericCellValue()));//属性赋值
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BLANK:
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(), "");//属性赋值
                            break;
                        default:
                            System.out.println("未知类型------>" + cell);
                    }
                }
                jjgFbgcJtaqssJabxfhl.setJcsj(simpleDateFormat.parse((String) map.get("jcsj")));
                if (!map.get("wzjlx").toString().contains("圆柱") && !map.get("wzjlx").toString().contains("方柱")) {
                    throw new JjgysException(20001, "03交安波形防护栏实测数据.xlsx,请在数据的位置及类型中，加入圆柱或方柱字样，然后重新上传");
                } else {
                    jjgFbgcJtaqssJabxfhl.setWzjlx((String) map.get("wzjlx"));
                }
                jjgFbgcJtaqssJabxfhl.setJdhdgdz((String) map.get("jdhdgdz"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz1((String) map.get("jdhdscz1"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz2((String) map.get("jdhdscz2"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz3((String) map.get("jdhdscz3"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz4((String) map.get("jdhdscz4"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz5((String) map.get("jdhdscz5"));
                jjgFbgcJtaqssJabxfhl.setLzbhgdz((String) map.get("lzbhgdz"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz1((String) map.get("lzbhscz1"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz2((String) map.get("lzbhscz2"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz3((String) map.get("lzbhscz3"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz4((String) map.get("lzbhscz4"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz5((String) map.get("lzbhscz5"));
                jjgFbgcJtaqssJabxfhl.setZxgdgdz((String) map.get("zxgdgdz"));
                jjgFbgcJtaqssJabxfhl.setZxgdyxpsz((String) map.get("zxgdyxpsz"));
                jjgFbgcJtaqssJabxfhl.setZxgdyxpsf((String) map.get("zxgdyxpsf"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz1((String) map.get("zxgdscz1"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz2((String) map.get("zxgdscz2"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz3((String) map.get("zxgdscz3"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz4((String) map.get("zxgdscz4"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz5((String) map.get("zxgdscz5"));
                jjgFbgcJtaqssJabxfhl.setMrsdgdz((String) map.get("mrsdgdz"));
                jjgFbgcJtaqssJabxfhl.setMrsdscz((String) map.get("mrsdscz"));
                jjgFbgcJtaqssJabxfhl.setProname(commonInfoVo.getProname());
                jjgFbgcJtaqssJabxfhl.setHtd((String) map.get("htd"));
                jjgFbgcJtaqssJabxfhl.setFbgc("防护栏");
                jjgFbgcJtaqssJabxfhl.setCreatetime(new Date());
                jjgFbgcJtaqssJabxfhlJgfcMapper.insert(jjgFbgcJtaqssJabxfhl);
            }
        } else {
            throw new JjgysException(20001, "您上传的" + commonInfoVo.getHtd() + "合同段中03交安波形防护栏实测数据.xlsx,解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public void exportjabxfhl(HttpServletResponse response) throws IOException {
        File directory = new File("");// 参数为空
        String courseFile = null;
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();// 创建一个Excel文件
            XSSFCellStyle columnHeadStyle = JjgFbgcCommonUtils.tsfCellStyle(workbook);

            XSSFSheet sheet = workbook.createSheet("实测数据");// 创建一个Excel的Sheet
            sheet.setColumnWidth(0, 28 * 256);
            List<String> checklist = Arrays.asList("合同段名称","检测日期", "位置及类型", "基底厚度规定值(mm)", "基底厚度实测值1(mm)",
                    "基底厚度实测值2(mm)", "基底厚度实测值3(mm)", "基底厚度实测值4(mm)", "基底厚度实测值5(mm)", "立柱壁厚规定值(mm)", "立柱壁厚实测值1(mm)", "立柱壁厚实测值2(mm)",
                    "立柱壁厚实测值3(mm)", "立柱壁厚实测值4(mm)", "立柱壁厚实测值5(mm)", "中心高度规定值(mm)", "中心高度允许偏差+(mm)",
                    "中心高度允许偏差-(mm)", "中心高度实测值1(mm)", "中心高度实测值2(mm)", "中心高度实测值3(mm)", "中心高度实测值4(mm)", "中心高度实测值5(mm)", "埋入深度规定值(mm)", "埋入深度实测值(mm)");
            for (int i = 0; i < checklist.size(); i++) {
                XSSFRow row = sheet.createRow(i);// 创建第一行
                XSSFCell cell = row.createCell(0);// 创建第一行第一列
                cell.setCellValue(new XSSFRichTextString(checklist.get(i)));
                cell.setCellStyle(columnHeadStyle);
            }
            String filename = "03交安波形防护栏实测数据.xlsx";// 设置下载时客户端Excel的名称
            filename = new String((filename).getBytes("GBK"), "ISO8859_1");
            response.setContentType("application/octet-stream;charset=UTF-8");
            response.addHeader("Content-Disposition", "attachment;filename=" + filename);
            OutputStream ouputStream = response.getOutputStream();
            workbook.write(ouputStream);
            ouputStream.flush();
            ouputStream.close();
            courseFile = directory.getCanonicalPath();
        }catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            File t = new File(courseFile+File.separator+"03交安波形防护栏实测数据.xlsx");
            if (t.exists()){
                t.delete();
            }
        }


    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        List<String> htds = jjgFbgcJtaqssJabxfhlJgfcMapper.gethtd(proname);
        List<Map<String, Object>> mapList = new ArrayList<>();
        if (htds != null && htds.size() > 0) {
            for (String htd : htds) {
                String sheetname = "防护栏";
                //获取鉴定表文件
                File f = new File(jgfilepath + File.separator + proname + File.separator + htd + File.separator + "58交安钢防护栏.xlsx");
                if (!f.exists()) {
                    return null;
                } else {
                    XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
                    //读取工作表
                    XSSFSheet slSheet = xwb.getSheet(sheetname);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(10);//合同段名

                    Map<String, Object> jgmap1 = new HashMap<>();
                    Map<String, Object> jgmap2 = new HashMap<>();
                    Map<String, Object> jgmap3 = new HashMap<>();
                    Map<String, Object> jgmap4 = new HashMap<>();
                    DecimalFormat df = new DecimalFormat("0.00");
                    DecimalFormat decf = new DecimalFormat("0.##");
                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum - 3).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                        slSheet.getRow(lastRowNum - 2).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                        slSheet.getRow(lastRowNum - 1).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                        slSheet.getRow(lastRowNum).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                        slSheet.getRow(5).getCell(2).setCellType(XSSFCell.CELL_TYPE_STRING);
                        double zds1 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(3).getStringCellValue());
                        double hgds1 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(3).getStringCellValue());
                        double bhgds1 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(3).getStringCellValue());
                        double hgl1 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                        String zdsz1 = decf.format(zds1);
                        String hgdsz1 = decf.format(hgds1);
                        String bhgdsz1 = decf.format(bhgds1);
                        String hglz1 = df.format(hgl1);
                        jgmap1.put("htd", htd);
                        jgmap1.put("检测项目", "波形梁板基底金属厚度");
                        jgmap1.put("规定值或允许偏差", slSheet.getRow(5).getCell(2).getStringCellValue());
                        jgmap1.put("总点数", zdsz1);
                        jgmap1.put("合格点数", hgdsz1);
                        jgmap1.put("不合格点数", bhgdsz1);
                        jgmap1.put("合格率", hglz1);

                        slSheet.getRow(lastRowNum - 3).getCell(5).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                        slSheet.getRow(lastRowNum - 2).getCell(5).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                        slSheet.getRow(lastRowNum - 1).getCell(5).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                        slSheet.getRow(lastRowNum).getCell(5).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                        slSheet.getRow(5).getCell(5).setCellType(XSSFCell.CELL_TYPE_STRING);
                        double zds2 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(5).getStringCellValue());
                        double hgds2 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(5).getStringCellValue());
                        double bhgds2 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(5).getStringCellValue());
                        double hgl2 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(5).getStringCellValue());
                        String zdsz2 = decf.format(zds2);
                        String hgdsz2 = decf.format(hgds2);
                        String bhgdsz2 = decf.format(bhgds2);
                        String hglz2 = df.format(hgl2);
                        jgmap2.put("htd", htd);
                        jgmap2.put("检测项目", "波形梁钢护栏立柱壁厚");
                        jgmap2.put("规定值或允许偏差", slSheet.getRow(5).getCell(5).getStringCellValue());
                        jgmap2.put("总点数", zdsz2);
                        jgmap2.put("合格点数", hgdsz2);
                        jgmap2.put("不合格点数", bhgdsz2);
                        jgmap2.put("合格率", hglz2);

                        slSheet.getRow(lastRowNum - 3).getCell(8).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                        slSheet.getRow(lastRowNum - 2).getCell(8).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                        slSheet.getRow(lastRowNum - 1).getCell(8).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                        slSheet.getRow(lastRowNum).getCell(8).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                        slSheet.getRow(5).getCell(8).setCellType(XSSFCell.CELL_TYPE_STRING);
                        double zds3 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(8).getStringCellValue());
                        double hgds3 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(8).getStringCellValue());
                        double bhgds3 = Double.valueOf(slSheet.getRow(lastRowNum - 1).getCell(8).getStringCellValue());
                        double hgl3 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(8).getStringCellValue());
                        String zdsz3 = decf.format(zds3);
                        String hgdsz3 = decf.format(hgds3);
                        String bhgdsz3 = decf.format(bhgds3);
                        String hglz3 = df.format(hgl3);
                        jgmap3.put("检测项目", "波形梁钢护栏横梁中心高度");
                        jgmap3.put("htd", htd);
                        jgmap3.put("规定值或允许偏差", slSheet.getRow(5).getCell(8).getStringCellValue());
                        jgmap3.put("总点数", zdsz3);
                        jgmap3.put("合格点数", hgdsz3);
                        jgmap3.put("不合格点数", bhgdsz3);
                        jgmap3.put("合格率", hglz3);

                        slSheet.getRow(lastRowNum - 3).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                        slSheet.getRow(lastRowNum - 2).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                        slSheet.getRow(lastRowNum - 1).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                        slSheet.getRow(lastRowNum).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                        slSheet.getRow(5).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);
                        double zds4 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(11).getStringCellValue());
                        double hgds4 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(11).getStringCellValue());
                        double bhgds4 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(11).getStringCellValue());
                        double hgl4 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(11).getStringCellValue());
                        String zdsz4 = decf.format(zds4);
                        String hgdsz4 = decf.format(hgds4);
                        String bhgdsz4 = decf.format(bhgds4);
                        String hglz4 = df.format(hgl4);
                        jgmap4.put("htd", htd);
                        jgmap4.put("检测项目", "波形梁钢护栏立柱埋入深度");
                        jgmap4.put("规定值或允许偏差", slSheet.getRow(5).getCell(11).getStringCellValue());
                        jgmap4.put("总点数", zdsz4);
                        jgmap4.put("合格点数", hgdsz4);
                        jgmap4.put("不合格点数", bhgdsz4);
                        jgmap4.put("合格率", hglz4);

                        mapList.add(jgmap1);
                        mapList.add(jgmap2);
                        mapList.add(jgmap3);
                        mapList.add(jgmap4);

                    }
                }
            }
        }
        return mapList;
    }

    @Override
    public List<String> selecthtd(String proname) {
        List<String> htds = jjgFbgcJtaqssJabxfhlJgfcMapper.gethtd(proname);
        return htds;
    }
}


