package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.project.JjgLqsSd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcGenerateTableMapper;
import glgc.jjgys.system.mapper.JjgHtdMapper;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@Service
public class JjgFbgcGenerateTablelServiceImpl extends ServiceImpl<JjgFbgcGenerateTableMapper,Object> implements JjgFbgcGenerateTablelService {

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Autowired
    private JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;

    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

    @Autowired
    private JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;

    @Autowired
    private JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;

    @Autowired
    private JjgFbgcLjgcXqgqdService jjgFbgcLjgcXqgqdService;

    @Autowired
    private JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;

    @Autowired
    private JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;

    @Autowired
    private JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;

    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmqdService jjgFbgcLmgcHntlmqdService;

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfService jjgFbgcLmgcLmgzsdsgpsfService;

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Autowired
    private JjgFbgcLmgcLmssxsService jjgFbgcLmgcLmssxsService;

    @Autowired
    private JjgFbgcLmgcLmwcService jjgFbgcLmgcLmwcService;

    @Autowired
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Autowired
    private JjgFbgcLmgcTlmxlbgcService jjgFbgcLmgcTlmxlbgcService;

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Autowired
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;

    @Autowired
    private JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;

    @Autowired
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;

    @Autowired
    private JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;

    @Autowired
    private JjgFbgcQlgcXbJgccService jjgFbgcQlgcXbJgccService;

    @Autowired
    private JjgFbgcQlgcXbBhchdService jjgFbgcQlgcXbBhchdService;

    @Autowired
    private JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Autowired
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;

    @Autowired
    private JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;

    @Autowired
    private JjgFbgcSdgcCqtqdService jjgFbgcSdgcCqtqdService;

    @Autowired
    private JjgFbgcSdgcDmpzdService jjgFbgcSdgcDmpzdService;

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private JjgFbgcQlgcQmpzdService jjgFbgcQlgcQmpzdService;

    @Autowired
    private JjgFbgcQlgcQmhpService jjgFbgcQlgcQmhpService;

    @Autowired
    private JjgFbgcQlgcQmgzsdService jjgFbgcQlgcQmgzsdService;

    @Autowired
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;

    @Autowired
    private JjgFbgcSdgcSdlqlmysdService jjgFbgcSdgcSdlqlmysdService;

    @Autowired
    private JjgFbgcSdgcLmssxsService jjgFbgcSdgcLmssxsService;

    @Autowired
    private JjgFbgcSdgcHntlmqdService jjgFbgcSdgcHntlmqdService;

    @Autowired
    private JjgFbgcSdgcTlmxlbgcService jjgFbgcSdgcTlmxlbgcService;

    @Autowired
    private JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;

    @Autowired
    private JjgFbgcSdgcSdhntlmhdzxfService jjgFbgcSdgcSdhntlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcGssdlqlmhdzxfService jjgFbgcSdgcGssdlqlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcLmgzsdsgpsfService jjgFbgcSdgcLmgzsdsgpsfService;


    @Autowired
    private JjgHtdService jjgHtdService;

    @Autowired
    private JjgFbgcSdgcJkService jjgFbgcSdgcJkService;


    @Autowired
    private JjgZdhGzsdService jjgZdhGzsdService;

    @Autowired
    private JjgZdhMcxsService jjgZdhMcxsService;

    @Autowired
    private JjgZdhPzdService jjgZdhPzdService;

    @Autowired
    private JjgZdhLdhdService jjgZdhLdhdService;

    @Autowired
    private JjgZdhCzService jjgZdhCzService;

    @Autowired
    private JjgLqsSdService jjgLqsSdService;

    //=========
    @Autowired
    private JjgFbgcQlgcZdhgzsdService jjgFbgcQlgcZdhgzsdService;

    @Autowired
    private JjgFbgcQlgcZdhmcxsService jjgFbgcQlgcZdhmcxsService;

    @Autowired
    private JjgFbgcQlgcZdhpzdService jjgFbgcQlgcZdhpzdService;

    @Autowired
    private JjgFbgcSdgcZdhczService jjgFbgcSdgcZdhczService;

    @Autowired
    private JjgFbgcSdgcZdhgzsdService jjgFbgcSdgcZdhgzsdService;

    @Autowired
    private JjgFbgcSdgcZdhldhdService jjgFbgcSdgcZdhldhdService;

    @Autowired
    private JjgFbgcSdgcZdhmcxsService jjgFbgcSdgcZdhmcxsService;

    @Autowired
    private JjgFbgcSdgcZdhpzdService jjgFbgcSdgcZdhpzdService;

    @Autowired
    private JjgHtdMapper jjgHtdMapper;



    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generatePdb(CommonInfoVo commonInfoVo) throws IOException {

        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");

        /**
         * 路基工程，路面工程，交安工程
         * 桥梁工程和隧道工程是按  桥梁和隧道
         *
         * 或许也可以不用查询这个合同段的类型，根据当前项目和合同段路基下的所有文件. 2025.2.13:师兄，我觉得要查，我查了
         * 只需要把有关路基，路面和交安相关的文件筛选出来。
         */

        /* 查询合同段的合同类型，有两种情况：1. 路面工程 2. 路面工程+其他
            1都是要合起来的
            2分开归属各个部分
        */
        // flag = 1 合起来 ；flag = 2分开归属
        int flag = 0;
        String lx = jjgHtdMapper.selectlx(proname, htd);
        if(lx.contains("路面工程")){
            if(lx.equals("路面工程")){
                flag = 1;
            }else {
                flag = 2;
            }
        }else{
            flag = 1;// 只有桥梁、交安都是要合起来
        }


        String path = filespath+ File.separator+proname+File.separator+htd+File.separator;

        List<Map<String,Object>> resultlist = new ArrayList<>();
        List<String> filteredFiles = filterFiles(path);
        boolean ysd = false; // 由于之前的旧逻辑一下子就处理了两个表，所以用这个变量专门控制压实度是否已经添加过，防止重复添加

        for (String value : filteredFiles) {
            switch (value) {
                case "08路基涵洞砼强度.xlsx":
                    // 路基涵洞
                    List<Map<String, Object>> maps1 = jjgFbgcLjgcHdgqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps1 = new ArrayList<>();

                    List<String> sjqdlist = jjgFbgcLjgcHdgqdService.selectsjqd(proname,htd);
                    for (String s : sjqdlist) {
                        for (Map<String, Object> map : maps1) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《涵洞砼强度质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "*砼强度");
                            newMap.put("yxps", "C"+s);
                            newMap.put("sheetname", "-分部路基");
                            newMap.put("fbgc", "涵洞");
                            newMaps1.add(newMap);
                        }
                    }
                    maps1 = newMaps1;
                    resultlist.addAll(maps1);
                    break;
                case "09路基涵洞结构尺寸.xlsx":
                    // 路基涵洞结构尺寸
                    List<Map<String, Object>> maps2 = jjgFbgcLjgcHdjgccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps2 = new ArrayList<>();

                    List<String> yxpslist = jjgFbgcLjgcHdjgccService.selectyxps(proname,htd);
                    for (String s : yxpslist) {
                        for (Map<String, Object> map : maps2) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《涵洞结构尺寸质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "结构尺寸");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "涵洞");
                            newMaps2.add(newMap);
                        }
                    }
                    maps2 = newMaps2;
                    resultlist.addAll(maps2);
                    break;
                case "03路基边坡.xlsx":
                    List<Map<String, Object>> maps3 = jjgFbgcLjgcLjbpService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps3 = new ArrayList<>();

                    //List<String> yxpslistbp = jjgFbgcLjgcLjbpService.selectyxps(proname,htd);暂时先这样
                    for (Map<String, Object> map : maps3) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基边坡质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap.put("ccname", "边坡");
                        newMap.put("yxps", "不陡于设计");//先写死
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("fbgc", "路基土石方");
                        newMaps3.add(newMap);
                    }
                    maps3 = newMaps3;
                    resultlist.addAll(maps3);
                    break;
                case "01路基沉降量.xlsx":
                    List<Map<String, Object>> maps4 = jjgFbgcLjgcLjcjService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps4 = new ArrayList<>();

                    List<String> yxpslistcj = jjgFbgcLjgcLjcjService.selectyxps(proname,htd);
                    for (String s : yxpslistcj) {
                        for (Map<String, Object> map : maps4) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《路基压实度沉降质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "△沉降");
                            newMap.put("yxps", "≤"+s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "路基土石方");
                            newMaps4.add(newMap);
                        }
                    }
                    maps4 = newMaps4;
                    resultlist.addAll(maps4);
                    break;
                case "01路基土石方压实度.xlsx":
                    List<Map<String, Object>> maps5 = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps5 = new ArrayList<>();
                    for (Map<String, Object> map : maps5) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基压实度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap.put("ccname", "△压实度"+map.get("压实度项目"));
                        newMap.put("yxps", map.get("规定值"));
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("fbgc", "路基土石方");
                        newMaps5.add(newMap);

                    }
                    maps5 = newMaps5;
                    resultlist.addAll(maps5);
                    break;

                case "02路基弯沉(贝克曼梁法).xlsx":
                    List<Map<String, Object>> maps6 = jjgFbgcLjgcLjwcService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps6 = new ArrayList<>();
                    for (Map<String, Object> map : maps6) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基弯沉(贝克曼梁法)质量鉴定表》检测"+map.get("检测单元数")+"个评定单元,合格"+map.get("合格单元数")+"个评定单元");
                        newMap.put("ccname", "△弯沉(贝克曼梁法)");
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("yxps", map.get("规定值"));
                        newMap.put("fbgc", "路基土石方");
                        newMaps6.add(newMap);
                    }
                    maps6 = newMaps6;
                    resultlist.addAll(maps6);
                    break;
                case "02路基弯沉(落锤法).xlsx":
                    List<Map<String, Object>> maps66 = jjgFbgcLjgcLjwcLcfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps66 = new ArrayList<>();
                    for (Map<String, Object> map : maps66) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基弯沉(落锤法)质量鉴定表》检测"+map.get("检测单元数")+"个评定单元,合格"+map.get("合格单元数")+"个评定单元");
                        newMap.put("ccname", "△弯沉(落锤法)");
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("yxps", map.get("规定值"));
                        newMap.put("fbgc", "路基土石方");
                        newMaps66.add(newMap);
                    }
                    maps66 = newMaps66;
                    resultlist.addAll(maps66);
                    break;
                case "04路基排水断面尺寸.xlsx":
                    List<Map<String, Object>> maps7 = jjgFbgcLjgcPsdmccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps7 = new ArrayList<>();
                    List<Map<String,Object>> yxpslistcc = jjgFbgcLjgcPsdmccService.selectyxps(proname,htd);
                    for (Map<String, Object> s : yxpslistcc) {
                        for (Map<String, Object> map : maps7) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            String va="";
                            if (s.get("yxwcz").equals(s.get("yxwcf"))){
                                va = s.get("sjz")+"±"+s.get("yxwcz");
                            }else {
                                va = s.get("sjz")+"+"+s.get("yxwcz")+";"+s.get("sjz")+"-"+s.get("yxwcf");
                            }
                            newMap.put("filename", "详见《结构（断面）尺寸质量鉴定表》检测"+map.get("检测总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "断面尺寸");
                            newMap.put("yxps", va);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "排水工程");
                            newMaps7.add(newMap);
                        }

                    }
                    maps7 = newMaps7;
                    resultlist.addAll(maps7);
                    break;
                case "05路基排水铺砌厚度.xlsx":
                    List<Map<String, Object>> maps8 = jjgFbgcLjgcPspqhdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps8 = new ArrayList<>();
                    List<String> sjzlistpqhd = jjgFbgcLjgcPspqhdService.selectyxps(proname,htd);
                    for (String s : sjzlistpqhd) {
                        for (Map<String, Object> map : maps8) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《排水铺砌厚度质量鉴定表》检测"+map.get("检测总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "铺砌厚度");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "排水工程");
                            newMaps8.add(newMap);
                        }
                    }
                    maps8 = newMaps8;
                    resultlist.addAll(maps8);
                    break;
                case "06路基小桥砼强度.xlsx":
                    List<Map<String, Object>> maps9 = jjgFbgcLjgcXqgqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps9 = new ArrayList<>();
                    List<String> sjzlistsjqd = jjgFbgcLjgcXqgqdService.selectsjqd(proname,htd);
                    for (String s : sjzlistsjqd) {
                        for (Map<String, Object> map : maps9) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《小桥砼强度质量鉴定表》检测"+map.get("检测总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "*砼强度");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "小桥");
                            newMaps9.add(newMap);
                        }
                    }
                    //maps9 = newMaps9;
                    resultlist.addAll(newMaps9);
                    break;
                case "07路基小桥结构尺寸.xlsx":
                    List<Map<String, Object>> maps10 = jjgFbgcLjgcXqjgccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps10 = new ArrayList<>();
                    List<Map<String,Object>> yxpslistxq = jjgFbgcLjgcXqjgccService.selectyxps(proname,htd);
                    for (Map<String, Object> yxpcmap : yxpslistxq) {
                        for (Map<String, Object> jdjgmap : maps10) {
                            Map<String, Object> newMap = new HashMap<>(jdjgmap);
                            String va="";
                            if (yxpcmap.get("yxwcz").equals(yxpcmap.get("yxwcf"))){
                                va = "±"+yxpcmap.get("yxwcz");
                            }else {
                                va = "+"+yxpcmap.get("yxwcz")+";"+"-"+yxpcmap.get("yxwcf");
                            }
                            newMap.put("filename", "详见《小桥结构尺寸质量鉴定表》检测"+jdjgmap.get("检测总点数")+"点,合格"+jdjgmap.get("合格点数")+"点");
                            newMap.put("ccname", "主要结构尺寸");
                            newMap.put("yxps", va);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "小桥");
                            newMaps10.add(newMap);
                        }

                    }
                    maps10 = newMaps10;
                    resultlist.addAll(maps10);
                    break;
                case "11路基支挡断面尺寸.xlsx":
                    List<Map<String, Object>> maps11 = jjgFbgcLjgcZddmccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps11 = new ArrayList<>();
                    List<Map<String,Object>> yxpslistzd = jjgFbgcLjgcZddmccService.selectyxps(proname,htd);
                    for (Map<String, Object> yxps : yxpslistzd) {
                        for (Map<String, Object> jdjg : maps11) {
                            Map<String, Object> newMap = new HashMap<>(jdjg);
                            newMap.put("filename", "详见《支挡工程结构尺寸质量鉴定表》检测"+jdjg.get("检测总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "△断面尺寸");
                            newMap.put("yxps", yxps.get("result"));
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "支挡工程");
                            newMaps11.add(newMap);
                        }
                    }
                    maps11 = newMaps11;
                    resultlist.addAll(maps11);
                    break;
                case "10路基支挡砼强度.xlsx":
                    List<Map<String, Object>> maps12 = jjgFbgcLjgcZdgqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps12 = new ArrayList<>();

                    List<Map<String,Object>> yxpslistzdtqd = jjgFbgcLjgcZdgqdService.selectsjqd(proname,htd);
                    for (Map<String, Object> sjqd : yxpslistzdtqd) {
                        for (Map<String, Object> jdjg : maps12) {

                            Map<String, Object> newMap = new HashMap<>(jdjg);
                            newMap.put("filename", "详见《支挡工程砼强度质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "*砼强度");
                            newMap.put("yxps", sjqd.get("sjqd"));
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "支挡工程");
                            newMaps12.add(newMap);
                        }
                    }
                    maps12 = newMaps12;
                    resultlist.addAll(maps12);
                    break;

                case "56交安标志.xlsx":
                    List<Map<String, Object>> maps13 = jjgFbgcJtaqssJabzService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps13 = new ArrayList<>();
                    double zdsd = 0;
                    double hgdsd = 0;
                    String hgl = "";
                    for (Map<String, Object> jdjg : maps13) {
                        if (jdjg.get("项目").toString().contains("反光膜逆反射系数")){
                            double zds = Double.valueOf(jdjg.get("总点数").toString());
                            double hgds = Double.valueOf(jdjg.get("合格点数").toString());
                            zdsd+=zds;
                            hgdsd+=hgds;
                        }
                    }
                    if (zdsd != 0|| hgdsd !=0 ){
                        hgl = df.format(hgdsd/zdsd*100);
                    }else {
                        hgl = "0";
                    }
                    for (Map<String, Object> result : maps13) {
                        Map<String, Object> newMap = new HashMap<>(result);
                        if (result.get("项目").toString().contains("反光膜逆反射系数")){
                            newMap.put("filename", "详见《交通标志板安装质量鉴定表》检测"+decf.format(zdsd)+"点,合格"+decf.format(hgdsd)+"点");
                            newMap.put("合格率", hgl);
                            newMap.put("总点数", zdsd);
                            newMap.put("合格点数", hgdsd);

                        }else {
                            newMap.put("filename", "详见《交通标志板安装质量鉴定表》检测"+decf.format(result.get("总点数"))+"点,合格"+decf.format(result.get("合格点数"))+"点");
                        }
                        newMap.put("ccname", result.get("项目"));
                        newMap.put("yxps", result.get("规定值或允许偏差"));
                        newMap.put("sheetname", "分部-交安");
                        newMap.put("fbgc", "标志");
                        newMaps13.add(newMap);
                    }
                    System.out.println(newMaps13);
                    resultlist.addAll(newMaps13);
                    break;
                case "57交安标线厚度.xlsx":
                    List<Map<String, Object>> maps14 = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps14 = new ArrayList<>();
                    for (Map<String, Object> jdjg : maps14) {
                        Map<String, Object> newMap = new HashMap<>(jdjg);
                        if (jdjg.get("检测项目").toString().equals("交安标线厚度")){
                            newMap.put("filename", "详见《道路交通标线施工质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "△标线厚度");
                            newMap.put("yxps", jdjg.get("规定值或允许偏差"));
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "标线");
                            newMaps14.add(newMap);
                        }
                        if (jdjg.get("检测项目").toString().contains("逆反射系数")){
                            newMap.put("filename", "详见《道路交通标线施工质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "△反光标线逆反射亮度系数");
                            newMap.put("yxps", jdjg.get("规定值或允许偏差"));
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "标线");
                            newMaps14.add(newMap);
                        }

                    }
                    maps14 = newMaps14;
                    resultlist.addAll(maps14);
                    break;
                case "58交安钢防护栏.xlsx":
                    List<Map<String, Object>> maps15 = jjgFbgcJtaqssJabxfhlService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps15 = new ArrayList<>();
                    for (Map<String, Object> jdjg : maps15) {
                        Map<String, Object> newMap = new HashMap<>(jdjg);
                        newMap.put("filename", "详见《道路防护栏施工质量鉴定表（波形梁钢护栏）》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                        newMap.put("ccname", "*"+jdjg.get("检测项目"));
                        newMap.put("yxps", jdjg.get("规定值或允许偏差"));
                        newMap.put("sheetname", "分部-交安");
                        newMap.put("fbgc", "防护栏");
                        newMaps15.add(newMap);
                    }
                    maps15 = newMaps15;
                    resultlist.addAll(maps15);
                    break;
                case "59交安砼护栏强度.xlsx":
                    List<Map<String, Object>> maps16 = jjgFbgcJtaqssJathlqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps16 = new ArrayList<>();
                    List<String> yxps = jjgFbgcJtaqssJathlqdService.selectsjqd(proname,htd);
                    for (String s : yxps) {
                        for (Map<String, Object> jdjg : maps16) {
                            Map<String, Object> newMap = new HashMap<>(jdjg);
                            newMap.put("filename", "详见《交安工程砼护栏强度质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "*砼护栏强度");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "防护栏");
                            newMaps16.add(newMap);
                        }
                    }
                    maps16 = newMaps16;
                    resultlist.addAll(maps16);
                    break;
                case "60交安砼护栏断面尺寸.xlsx":
                    List<Map<String, Object>> maps17 = jjgFbgcJtaqssJathldmccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps17 = new ArrayList<>();
                    List<Map<String,Object>> yxpsjalist = jjgFbgcJtaqssJathldmccService.selectyxpc(proname,htd);
                    for (Map<String, Object> yxpcmap : yxpsjalist) {
                        for (Map<String, Object> jdjgmap : maps17) {

                            Map<String, Object> newMap = new HashMap<>(jdjgmap);
                            String va="";
                            if (yxpcmap.get("yxwcz").equals(yxpcmap.get("yxwcf"))){
                                va = "±"+yxpcmap.get("yxwcz");
                            }else {
                                va = "+"+yxpcmap.get("yxwcz")+";"+"-"+yxpcmap.get("yxwcf");
                            }
                            newMap.put("filename", "详见《交安砼护栏断面尺寸质量鉴定表》检测"+jdjgmap.get("检测总点数")+"点,合格"+jdjgmap.get("合格点数")+"点");
                            newMap.put("ccname", "△砼护栏断面尺寸");
                            newMap.put("yxps", va);
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "防护栏");
                            newMaps17.add(newMap);
                        }
                    }
                    maps17 = newMaps17;
                    resultlist.addAll(maps17);
                    break;

                /*case "12沥青路面压实度.xlsx":
                    //分工作簿
                    List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultysd = new ArrayList<>();
                    double sdjcds = 0;
                    double sdhgds = 0;
                    double lmjcds = 0;
                    double lmhgds = 0;
                    String sdgdz="";
                    String lmgdz="";
                    for (Map<String, Object> map : list) {
                        if (map.get("路面类型").toString().contains("隧道右幅") || map.get("路面类型").toString().contains("隧道左幅")){
                            sdjcds += Double.valueOf(map.get("检测点数").toString());
                            sdhgds += Double.valueOf(map.get("合格点数").toString());
                            sdgdz = map.get("密度规定值").toString();

                        }else {
                        //if (map.get("路面类型").toString().contains("沥青路面压实度右幅") || map.get("路面类型").toString().contains("沥青路面压实度左幅")){
                            lmjcds += Double.valueOf(map.get("检测点数").toString());
                            lmhgds += Double.valueOf(map.get("合格点数").toString());
                            lmgdz = map.get("密度规定值").toString();

                        }
                    }
                    double gdz1 = 0;
                    if (!sdgdz.equals("")){
                        gdz1 = Double.valueOf(sdgdz);
                    }

                    String gdz2= String.valueOf(gdz1);
                    Map<String, Object> newMap1 = new HashMap<>();
                    newMap1.put("filename","详见《沥青路面压实度质量鉴定表》检测"+sdjcds+"点,合格"+sdhgds+"点");
                    newMap1.put("ccname", "△沥青路面压实度(隧道路面)");
                    newMap1.put("ccname2", "隧道路面");
                    newMap1.put("yxps", gdz2);
                    newMap1.put("sheetname", "分部-路面");
                    newMap1.put("fbgc", "路面面层");
                    newMap1.put("检测点数", sdjcds);
                    newMap1.put("合格点数", sdhgds);
                    newMap1.put("合格率", (sdjcds != 0) ? df.format(sdhgds/sdjcds*100) : "0");

                    double gdz3 = 0;
                    if (!lmgdz.equals("")){
                        gdz3 = Double.valueOf(lmgdz);
                    }
                    String gdz4 = String.valueOf(gdz3);
                    Map<String, Object> newMap2 = new HashMap<>();
                    newMap2.put("filename","详见《沥青路面压实度质量鉴定表》检测"+lmjcds+"点,合格"+lmhgds+"点");
                    newMap2.put("ccname", "△沥青路面压实度(路面面层)");
                    newMap2.put("ccname2", "路面面层");
                    newMap2.put("yxps", gdz4);
                    newMap2.put("sheetname", "分部-路面");
                    newMap2.put("fbgc", "路面面层");
                    newMap2.put("检测点数", lmjcds);
                    newMap2.put("合格点数", lmhgds);
                    newMap2.put("合格率", (lmjcds != 0) ? df.format(lmhgds/lmjcds*100) : "0");
                    resultysd.add(newMap1);
                    resultysd.add(newMap2);
                    resultlist.addAll(resultysd);
                    break;*/
                case "13路面弯沉(贝克曼梁法).xlsx":
                    List<Map<String, Object>> lmwclist = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultwc = new ArrayList<>();
                    if(flag == 1){
                        Map<String, Object> wcMap = new HashMap<>();
                        wcMap.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+lmwclist.get(0).get("检测单元数")+"个评定单元,合格"+lmwclist.get(0).get("合格单元数")+"个评定单元");
                        wcMap.put("ccname","△沥青路面弯沉(贝克曼梁法)");
                        wcMap.put("ccname2","贝克曼梁法");
                        wcMap.put("sheetname","分部-路面");
                        wcMap.put("fbgc","路面面层");
                        wcMap.put("合格率",lmwclist.get(0).get("合格率"));
                        wcMap.put("yxps",lmwclist.get(0).get("规定值"));
                        resultwc.add(wcMap);
                    }else if(flag==2){
                        // 这一part只有路面互通和连接线，没有桥没有隧
                        for (Map<String, Object> map : lmwclist) {
                            Map<String, Object> wclcfMap = new HashMap<>();
                            if (map.get("分部工程名称").toString().equals("路面/互通类")){
                                wclcfMap.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+lmwclist.get(0).get("检测单元数")+"个评定单元,合格"+lmwclist.get(0).get("合格单元数")+"个评定单元");
                                wclcfMap.put("ccname","△沥青路面弯沉(贝克曼梁法)");
                            }else{
                                wclcfMap.put("filename","详见《连接线路面弯沉质量鉴定结果汇总表》检测"+lmwclist.get(0).get("检测单元数")+"个评定单元,合格"+lmwclist.get(0).get("合格单元数")+"个评定单元");
                                wclcfMap.put("ccname","△沥青路面弯沉(贝克曼梁法)(连接线)");
                            }
                            wclcfMap.put("ccname2","贝克曼梁法");
                            wclcfMap.put("sheetname","分部-路面");
                            wclcfMap.put("fbgc","路面面层");
                            wclcfMap.put("合格率",lmwclist.get(0).get("合格率"));
                            wclcfMap.put("yxps",lmwclist.get(0).get("规定值"));
                            resultwc.add(wclcfMap);
                        }
                    }

                    resultlist.addAll(resultwc);
                    break;
                case "13路面弯沉(落锤法).xlsx":
                    List<Map<String, Object>> wclcflist = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultwclcf = new ArrayList<>();
                    if(flag==1){
                        Map<String, Object> wclcfMap = new HashMap<>();
                        wclcfMap.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+wclcflist.get(0).get("检测单元数")+"个评定单元,合格"+wclcflist.get(0).get("合格单元数")+"个评定单元");
                        wclcfMap.put("ccname","△沥青路面弯沉(落锤法)");
                        wclcfMap.put("ccname2","落锤法");
                        wclcfMap.put("sheetname","分部-路面");
                        wclcfMap.put("fbgc","路面面层");
                        wclcfMap.put("合格率",wclcflist.get(0).get("合格率"));
                        wclcfMap.put("yxps",wclcflist.get(0).get("规定值"));
                        resultwclcf.add(wclcfMap);
                    }else if(flag==2){
                        // 这一part只有路面互通和连接线，没有桥没有隧
                        for (Map<String, Object> map : wclcflist) {
                            Map<String, Object> wclcfMap = new HashMap<>();
                            if (map.get("分部工程名称").toString().equals("路面/互通类")){
                                wclcfMap.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+wclcflist.get(0).get("检测单元数")+"个评定单元,合格"+wclcflist.get(0).get("合格单元数")+"个评定单元");
                                wclcfMap.put("ccname","△沥青路面弯沉(落锤法)");
                            }else{
                                wclcfMap.put("filename","详见《连接线路面弯沉质量鉴定结果汇总表》检测"+wclcflist.get(0).get("检测单元数")+"个评定单元,合格"+wclcflist.get(0).get("合格单元数")+"个评定单元");
                                wclcfMap.put("ccname","△沥青路面弯沉(落锤法)(连接线)");
                            }
                            wclcfMap.put("ccname2","落锤法");
                            wclcfMap.put("sheetname","分部-路面");
                            wclcfMap.put("fbgc","路面面层");
                            wclcfMap.put("合格率",wclcflist.get(0).get("合格率"));
                            wclcfMap.put("yxps",wclcflist.get(0).get("规定值"));
                            resultwclcf.add(wclcfMap);
                        }
                    }
                    resultlist.addAll(resultwclcf);
                    break;
                case "15沥青路面渗水系数.xlsx":
                    //分工作簿
                    List<Map<String, Object>> ssxslist = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultss = new ArrayList<>();
                    double sdssjcds = 0; // 隧道渗水检测点数
                    double sdsshgds = 0;
                    double lmssjcds = 0; // 路面渗水检测点数(注意，后面用反了）
                    double lmsshgds = 0;
                    double ljxssjcds = 0; // 连接线渗水检测点数
                    double ljxsshgds = 0;
                    String sdssgdz="";
                    String lmssgdz="";
                    String ljxssgdz="";
                    boolean sda = false,lma = false, ljxa = false;
                    if(flag == 1){
                        for (Map<String, Object> map : ssxslist) {
                            if (map.get("检测项目").toString().contains("沥青路面")){
                                sda = true;
                                sdssjcds += Double.valueOf(map.get("检测点数").toString());
                                sdsshgds += Double.valueOf(map.get("合格点数").toString());
                                sdssgdz = map.get("规定值").toString();

                            }
                            if (map.get("检测项目").toString().contains("隧道路面")){
                                lma = true;

                                lmssjcds += Double.valueOf(map.get("检测点数").toString());
                                lmsshgds += Double.valueOf(map.get("合格点数").toString());
                                lmssgdz = map.get("规定值").toString();
                            }
                            if (map.get("检测项目").toString().contains("连接线路面")){
                                ljxa = true;

                                ljxssjcds += Double.valueOf(map.get("检测点数").toString());
                                ljxsshgds += Double.valueOf(map.get("合格点数").toString());
                                ljxssgdz = map.get("规定值").toString();
                            }
                        }
                        if (sda){
                            double ssgdz1 = Double.valueOf(sdssgdz);
                            String ssgdz2= String.valueOf(ssgdz1);
                            Map<String, Object> newMapss = new HashMap<>();
                            newMapss.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(sdssjcds)+"点,合格"+decf.format(sdsshgds)+"点");
                            newMapss.put("ccname", "沥青路面渗水系数(路面面层)");
                            newMapss.put("ccname2", "路面面层");
                            newMapss.put("yxps", ssgdz2);
                            newMapss.put("sheetname", "分部-路面");
                            newMapss.put("fbgc", "路面面层");
                            newMapss.put("检测点数", decf.format(sdssjcds));
                            newMapss.put("合格点数", decf.format(sdsshgds));
                            newMapss.put("合格率", (sdssjcds != 0) ? df.format(sdsshgds/sdssjcds*100) : "0");
                            resultss.add(newMapss);
                        }
                        if (lma){
                            double ssgdz3 = Double.valueOf(lmssgdz);
                            String ssgdz4 = String.valueOf(ssgdz3);
                            Map<String, Object> newMapss1 = new HashMap<>();
                            newMapss1.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(lmssjcds)+"点,合格"+decf.format(lmsshgds)+"点");
                            newMapss1.put("ccname", "沥青路面渗水系数(隧道路面)");
                            newMapss1.put("ccname2", "隧道路面");
                            newMapss1.put("yxps", ssgdz4);
                            newMapss1.put("sheetname", "分部-路面");
                            newMapss1.put("fbgc", "路面面层");
                            newMapss1.put("检测点数", decf.format(lmssjcds));
                            newMapss1.put("合格点数", decf.format(lmsshgds));
                            newMapss1.put("合格率", (lmssjcds != 0) ? df.format(lmsshgds/lmssjcds*100) : "0");
                            resultss.add(newMapss1);
                        }
                        if (ljxa){
                            double ssgdz5 = Double.valueOf(ljxssgdz);
                            String ssgdz6 = String.valueOf(ssgdz5);
                            Map<String, Object> newMapss1 = new HashMap<>();
                            newMapss1.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(ljxssjcds)+"点,合格"+decf.format(ljxsshgds)+"点");
                            newMapss1.put("ccname", "沥青路面渗水系数(连接线路面)");
                            newMapss1.put("ccname2", "隧道路面");
                            newMapss1.put("yxps", ssgdz6);
                            newMapss1.put("sheetname", "分部-路面");
                            newMapss1.put("fbgc", "路面面层");
                            newMapss1.put("检测点数", decf.format(ljxssjcds));
                            newMapss1.put("合格点数", decf.format(ljxsshgds));
                            newMapss1.put("合格率", (ljxssjcds != 0) ? df.format(ljxsshgds/ljxssjcds*100) : "0");
                            resultss.add(newMapss1);
                        }
                    }else if(flag == 2){
                        ssxslist = processSsxsMapList(ssxslist);
                        for(Map<String, Object> map1 : ssxslist){
                            sdssjcds = 0;
                            sdsshgds = 0;
                            sdssgdz = "";

                            sdssjcds = Double.valueOf(map1.get("检测点数").toString());
                            sdsshgds = Double.valueOf(map1.get("合格点数").toString());
                            sdssgdz = map1.get("规定值").toString();

                            // 填写数据
                            Map<String, Object> map = new HashMap<>();
                            // 分情况

                            if(map1.get("分部工程名称").toString().contains("连接线")) {
                                map.put("filename","详见《连接线沥青路面渗水系数质量鉴定表》检测"+decf.format(sdssjcds)+"点,合格"+decf.format(sdsshgds)+"点");
                                map.put("ccname", "沥青路面渗水系数(连接线)");
                                map.put("ccname2", "路面面层");
                                map.put("sheetname", "分部-路面");
                            }else if(map1.get("分部工程名称").toString().contains("路面")){
                                map.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(sdssjcds)+"点,合格"+decf.format(sdsshgds)+"点");
                                map.put("ccname", "沥青路面渗水系数(路面面层)");
                                map.put("ccname2", "路面面层");
                                map.put("sheetname", "分部-路面");
                            }

                            else if (map1.get("分部工程名称").toString().contains("隧道")){
                                map.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(sdssjcds)+"点,合格"+decf.format(sdsshgds)+"点");
                                map.put("ccname", "沥青路面渗水系数(隧道路面)");
                                map.put("ccname2", "隧道路面");
                                map.put("sheetname", "分部-"+ map1.get("分部工程名称").toString());
                            }



                            //map.put("ccname3", "沥青路面");
                            map.put("yxps", sdssgdz);
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(sdssjcds));
                            map.put("合格点数", decf.format(sdsshgds));
                            map.put("合格率", (sdsshgds != 0) ? df.format(sdsshgds/sdssjcds*100) : "0");
                            resultss.add(map);
                        }
                    }
                    resultlist.addAll(resultss);
                    break;
                case "16混凝土路面强度.xlsx":
                    List<Map<String, Object>> lmqdlist = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultqd = new ArrayList<>();
                    Map<String, Object> qdMap = new HashMap<>();
                    qdMap.put("filename","详见《混凝土路面强度鉴定结果汇总表》检测"+lmqdlist.get(0).get("总点数")+"点,合格"+lmqdlist.get(0).get("合格点数")+"点");
                    qdMap.put("ccname","*砼路面强度");
                    qdMap.put("ccname2","");
                    qdMap.put("sheetname","分部-路面");
                    qdMap.put("fbgc","路面面层");
                    qdMap.put("合格率",lmqdlist.get(0).get("合格率"));
                    qdMap.put("检测点数",lmqdlist.get(0).get("检测点数"));
                    qdMap.put("合格点数",lmqdlist.get(0).get("合格点数"));
                    qdMap.put("yxps",lmqdlist.get(0).get("规定值"));
                    resultqd.add(qdMap);
                    resultlist.addAll(resultqd);
                    break;
                case "17混凝土路面相邻板高差.xlsx":
                    List<Map<String, Object>> list1 = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultgs = new ArrayList<>();
                    Map<String, Object> gsMap = new HashMap<>();
                    gsMap.put("filename","详见《混凝土路面相邻板高差质量鉴定表》检测"+list1.get(0).get("总点数")+"点,合格"+list1.get(0).get("合格点数")+"点");
                    gsMap.put("ccname","砼路面相邻板高差");
                    gsMap.put("ccname2","");
                    gsMap.put("sheetname","分部-路面");
                    gsMap.put("fbgc","路面面层");
                    gsMap.put("合格率",list1.get(0).get("合格率"));
                    gsMap.put("检测点数",list1.get(0).get("总点数"));
                    gsMap.put("合格点数",list1.get(0).get("合格点数"));
                    gsMap.put("yxps",list1.get(0).get("规定值"));
                    resultgs.add(gsMap);
                    resultlist.addAll(resultgs);
                    break;

                case "20构造深度手工铺沙法.xlsx":
                    List<Map<String, Object>> list2 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultgz = new ArrayList<>();
                    Map<String, Object> gzMap = new HashMap<>();
                    gzMap.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+list2.get(0).get("检测点数")+"点,合格"+list2.get(0).get("合格点数")+"点");
                    gzMap.put("ccname","构造深度");
                    gzMap.put("ccname2","");
                    gzMap.put("sheetname","分部-路面");
                    gzMap.put("fbgc","路面面层");
                    gzMap.put("合格率",list2.get(0).get("合格率"));
                    gzMap.put("检测点数",list2.get(0).get("检测点数"));
                    gzMap.put("合格点数",list2.get(0).get("合格点数"));
                    gzMap.put("yxps",list2.get(0).get("规定值"));
                    resultgz.add(gzMap);
                    resultlist.addAll(resultgz);
                    break;
                case "22沥青路面厚度-钻芯法.xlsx":
                    //连接线和路面和隧道放在一起
                    List<Map<String, Object>> list3 = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> listhdzxf = new ArrayList<>();
                    double hdjcds = 0;
                    double hdhgds = 0;
                    String sjz1 = "";
                    String sjz2 = "";

                    double hdjcds2 = 0;
                    double hdhgds2 = 0;
                    String sjz12 = "";
                    String sjz22 = "";

                    if(flag == 1){
                        for (Map<String, Object> map : list3) {
                            if (map.get("路面类型").toString().contains("隧道")){

                                hdjcds += Double.valueOf(map.get("总厚度检测点数").toString())+Double.valueOf(map.get("上面层厚度检测点数").toString());
                                hdhgds += Double.valueOf(map.get("总厚度合格点数").toString())+Double.valueOf(map.get("上面层厚度合格点数").toString());

                                sjz1 = map.get("总厚度设计值").toString();
                                sjz2 = map.get("上面层设计值").toString();


                            }else if (map.get("路面类型").toString().contains("路面") || map.get("路面类型").toString().contains("路面") ){
                                hdjcds2 += Double.valueOf(map.get("总厚度检测点数").toString())+Double.valueOf(map.get("上面层厚度检测点数").toString());
                                hdhgds2 += Double.valueOf(map.get("总厚度合格点数").toString())+Double.valueOf(map.get("上面层厚度合格点数").toString());

                                sjz12 = map.get("总厚度设计值").toString();
                                sjz22 = map.get("上面层设计值").toString();

                            }
                        }
                        Map map1 = new HashMap();
                        Map map2 = new HashMap();
                        map1.put("检测点数",decf.format(hdjcds));
                        map1.put("合格点数",decf.format(hdhgds));
                        map1.put("yxps",sjz1);
                        map1.put("filename","详见《沥青隧道路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds)+"点,合格"+decf.format(hdhgds)+"点");
                        map1.put("ccname","△厚度");
                        map1.put("ccname2","隧道路面");
                        map1.put("ccname3","沥青路面");
                        map1.put("ccname4","钻芯法");
                        map1.put("fbgc","路面面层");
                        map1.put("合格率",(hdjcds != 0) ? df.format(hdhgds/hdjcds*100) : "0");

                        map2.put("检测点数",decf.format(hdjcds));
                        map2.put("合格点数",decf.format(hdhgds));
                        map2.put("yxps",sjz2);
                        map2.put("filename","详见《沥青隧道路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds)+"点,合格"+decf.format(hdhgds)+"点");
                        map2.put("ccname","△厚度");
                        map2.put("ccname2","隧道路面");
                        map2.put("ccname3","沥青路面");
                        map2.put("ccname4","钻芯法");
                        map2.put("fbgc","路面面层");
                        map2.put("合格率",(hdjcds != 0) ? df.format(hdhgds/hdjcds*100) : "0");

                        Map map3 = new HashMap();
                        Map map4 = new HashMap();
                        map3.put("检测点数",decf.format(hdjcds2));
                        map3.put("合格点数",decf.format(hdhgds2));
                        map3.put("yxps",sjz12);
                        map3.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds2)+"点,合格"+decf.format(hdhgds2)+"点");
                        map3.put("ccname","△厚度");
                        map3.put("ccname2","路面面层");
                        map3.put("ccname3","沥青路面");
                        map3.put("ccname4","钻芯法");
                        map3.put("fbgc","路面面层");
                        map3.put("合格率",(hdjcds2 != 0) ? df.format(hdhgds2/hdjcds2*100) : "0");

                        map4.put("检测点数",decf.format(hdjcds2));
                        map4.put("合格点数",decf.format(hdhgds2));
                        map4.put("yxps",sjz22);
                        map4.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds2)+"点,合格"+decf.format(hdhgds2)+"点");
                        map4.put("ccname","△厚度");
                        map4.put("ccname2","路面面层");
                        map4.put("ccname3","沥青路面");
                        map4.put("ccname4","钻芯法");
                        map4.put("fbgc","路面面层");
                        map4.put("合格率",(hdjcds2 != 0) ? df.format(hdhgds2/hdjcds2*100) : "0");
                        listhdzxf.add(map1);
                        listhdzxf.add(map2);
                        listhdzxf.add(map3);
                        listhdzxf.add(map4);
                        resultlist.addAll(listhdzxf);
                        break;
                    }else {
                        // 分开计算
                        if (list3.size() > 0) {
                            // 对 mapList 进行排序
                            Collections.sort(list3, new Comparator<Map<String, Object>>() {
                                @Override
                                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                                    String roadName1 = (String) o1.get("路面名称");
                                    String roadName2 = (String) o2.get("路面名称");
                                    if (roadName1 == null) {
                                        return roadName2 == null ? 0 : -1;
                                    }
                                    if (roadName2 == null) {
                                        return 1;
                                    }
                                    return roadName1.compareTo(roadName2);
                                }
                            });
                        }

                        // 定义统一使用的变量
                        double totalJcds = 0;
                        double totalHgds = 0;
                        String totalSjz1 = "";
                        String totalSjz2 = "";
                        String lx_zxf = list3.get(0).get("路面类型").toString();
                        String roadName =list3.get(0).get("路面名称").toString();

                        for (Map<String, Object> map : list3) {

                            //String roadName = map.get("路面名称").toString();

                            // 累加检测点数和合格点数
                            double jcds = Double.valueOf(map.get("总厚度检测点数").toString()) + Double.valueOf(map.get("上面层厚度检测点数").toString());
                            double hgds = Double.valueOf(map.get("总厚度合格点数").toString()) + Double.valueOf(map.get("上面层厚度合格点数").toString());

                            String sjz1_1 = map.get("总厚度设计值").toString();
                            String sjz2_1 = map.get("上面层设计值").toString();

                            if (roadName.equals(map.get("路面名称").toString())) { // 相同的一类
                                totalJcds += jcds;
                                totalHgds += hgds;

                                totalSjz1 = sjz1_1;
                                totalSjz2 = sjz2_1;
                            } else {
                                // 把之前的加入到map
                                Map<String, Object> resultMap = new HashMap<>();
                                resultMap.put("检测点数", decf.format(totalJcds));
                                resultMap.put("合格点数", decf.format(totalHgds));
                                resultMap.put("yxps", totalSjz1 + "," + totalSjz2); // 先总后上
                                resultMap.put("filename", "详见《" + getFileName(lx_zxf) + "》检测" + decf.format(totalJcds) + "点,合格" + decf.format(totalHgds) + "点");

                                resultMap.put("ccname2", getCcname2(lx_zxf));
                                resultMap.put("ccname3", "沥青路面");
                                resultMap.put("ccname4", "钻芯法");
                                resultMap.put("fbgc", "路面面层");
                                // 这个指标的连接线貌似只有路面，如果不是，后面再改正
                                if(roadName.equals("路") || roadName.contains("连接线")) {
                                    resultMap.put("sheetname", "分部-路面");
                                    resultMap.put("ccname", roadName.contains("连接线")? "厚度(连接线)" : "厚度");
                                }
                                else {
                                    resultMap.put("ccname", "厚度");
                                    resultMap.put("sheetname", "分部-" + roadName);
                                }
                                resultMap.put("合格率", (totalJcds != 0) ? df.format(totalHgds / totalJcds * 100) : "0");
                                listhdzxf.add(resultMap);

                                // 重新初始化
                                totalJcds = jcds;
                                totalHgds = hgds;

                                totalSjz1 = sjz1_1;
                                totalSjz2 = sjz2_1;

                                roadName = map.get("路面名称").toString();
                                lx_zxf = map.get("路面类型").toString();
                            }
                        }

                        // 最后一次循环结束后，还需要将最后的结果加入到list中
                        if (totalJcds > 0 && totalHgds >= 0) {
                            Map<String, Object> resultMap = new HashMap<>();
                            resultMap.put("检测点数", decf.format(totalJcds));
                            resultMap.put("合格点数", decf.format(totalHgds));
                            resultMap.put("yxps", totalSjz1 + "," + totalSjz2); // 先总后上
                            resultMap.put("filename", "详见《" + getFileName(lx_zxf) + "》检测" + decf.format(totalJcds) + "点,合格" + decf.format(totalHgds) + "点");
                            resultMap.put("ccname", "厚度");
                            resultMap.put("ccname2", getCcname2(lx));
                            resultMap.put("ccname3", "沥青路面");
                            resultMap.put("ccname4", "钻芯法");
                            resultMap.put("fbgc", "路面面层");
                            if(roadName.equals("路")|| roadName.contains("连接线")) resultMap.put("sheetname", "分部-路面");
                            else resultMap.put("sheetname", "分部-" + roadName);
                            resultMap.put("合格率", (totalJcds != 0) ? df.format(totalHgds / totalJcds * 100) : "0");
                            listhdzxf.add(resultMap);
                        }



                        resultlist.addAll(listhdzxf);
                        break;
                    }
                case "23混凝土路面厚度-钻芯法.xlsx":
                    List<Map<String, Object>> list4 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resulthdzxf = new ArrayList<>();
                    Map<String, Object> maphdzxf = new HashMap<>();
                    maphdzxf.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+list4.get(0).get("检测点数")+"点,合格"+list4.get(0).get("合格点数")+"点");
                    maphdzxf.put("ccname", "△厚度");
                    maphdzxf.put("ccname2", "路面面层");
                    maphdzxf.put("ccname3", "混凝土路面");
                    maphdzxf.put("ccname4", "钻芯法");
                    maphdzxf.put("yxps", list4.get(0).get("允许偏差"));
                    maphdzxf.put("sheetname", "分部-路面");
                    maphdzxf.put("fbgc", "路面面层");
                    maphdzxf.put("检测点数", list4.get(0).get("检测点数"));
                    maphdzxf.put("合格点数", list4.get(0).get("合格点数"));
                    maphdzxf.put("合格率", list4.get(0).get("合格率"));
                    resulthdzxf.add(maphdzxf);
                    resultlist.addAll(resulthdzxf);
                    break;
                case "25桥梁下部墩台砼强度.xlsx":
                    List<Map<String, Object>> list11 = jjgFbgcQlgcXbTqdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qltqdlist = new ArrayList<>();
                    for (Map<String, Object> map : list11) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","墩台混凝土强度(MPa)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部墩台砼强度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qltqdlist.add(map5);
                    }
                    resultlist.addAll(qltqdlist);
                    break;

                case "26桥梁下部主要结构尺寸.xlsx":
                    List<Map<String, Object>> list12 = jjgFbgcQlgcXbJgccService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qljgcclist = new ArrayList<>();
                    for (Map<String, Object> map : list12) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","主要构件尺寸(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部主要结构尺寸质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qljgcclist.add(map5);
                    }
                    resultlist.addAll(qljgcclist);
                    break;

                case "27桥梁下部保护层厚度.xlsx":
                    List<Map<String, Object>> list13 = jjgFbgcQlgcXbBhchdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlbhclist = new ArrayList<>();
                    for (Map<String, Object> map : list13) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","钢筋保护层厚度(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部钢筋保护层厚度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlbhclist.add(map5);
                    }
                    resultlist.addAll(qlbhclist);
                    break;

                case "28桥梁下部墩台垂直度.xlsx":
                    List<Map<String, Object>> list14 = jjgFbgcQlgcXbSzdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlczdlist = new ArrayList<>();
                    for (Map<String, Object> map : list14) {
                        if (!"".equals(map.get("zds").toString()) && !"".equals(map.get("hgds").toString())){
                            double z = Double.valueOf(map.get("zds").toString());
                            double h = Double.valueOf(map.get("hgds").toString());
                            Map map5 = new HashMap();
                            map5.put("ccname","墩台垂直度(mm)");
                            map5.put("yxps",map.get("sjqd"));
                            map5.put("filename","详见《桥梁下部墩台垂直度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                            map5.put("sheetname", "分部-"+map.get("qlmc"));
                            map5.put("fbgc", "桥梁下部");
                            map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                            qlczdlist.add(map5);
                        }

                    }
                    resultlist.addAll(qlczdlist);
                    break;

                case "29桥梁上部砼强度.xlsx":
                    List<Map<String, Object>> list15 = jjgFbgcQlgcSbTqdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlsbqdlist = new ArrayList<>();
                    for (Map<String, Object> map : list15) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","墩台混凝土强度(MPa)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁上部墩台砼强度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁上部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlsbqdlist.add(map5);
                    }
                    resultlist.addAll(qlsbqdlist);
                    break;

                case "30桥梁上部主要结构尺寸.xlsx":
                    List<Map<String, Object>> list16 = jjgFbgcQlgcSbJgccService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlsbjgcclist = new ArrayList<>();
                    for (Map<String, Object> map : list16) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","主要构件尺寸(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁上部主要结构尺寸质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁上部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlsbjgcclist.add(map5);
                    }
                    resultlist.addAll(qlsbjgcclist);
                    break;
                case "31桥梁上部保护层厚度.xlsx":
                    List<Map<String, Object>> list17 = jjgFbgcQlgcSbBhchdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlsbbhclist = new ArrayList<>();
                    for (Map<String, Object> map : list17) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","钢筋保护层厚度(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁上部钢筋保护层厚度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁上部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlsbbhclist.add(map5);
                    }
                    resultlist.addAll(qlsbbhclist);
                    break;
                case "24路面横坡.xlsx":
                    //工作簿分路面，桥，隧道
                    List<Map<String, String>> list5 = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resulthp = new ArrayList<>();

                    // 按照flag = 1 还是 2分两套代码
                    double lmhphgds = 0;
                    double qmhphgds = 0;
                    double sdhphgds = 0;
                    double hnthphgds = 0;
                    double ljxhgds = 0;

                    double lmhpjcds = 0;
                    double qmhpjcds = 0;
                    double sdhpjcds = 0;
                    double hnthpjcds = 0;
                    double ljxjcds = 0;

                    String lmyxpc = "";
                    String sdyxpc = "";
                    String qmyxpc = "";
                    String hntyxpc = "";
                    String ljxyxpc = "";

                    boolean ljx = false; // 连接线处理逻辑
                    boolean a = false;
                    boolean b = false;
                    boolean c = false;
                    boolean d = false;
                    if (flag == 1) {
                        for (Map<String, String> map : list5) {
                            if (map.get("路面类型").contains("沥青路面") && (map.get("检测项目").equals("主线") || map.get("检测项目").equals("匝道"))) {
                                lmhphgds += Double.valueOf(map.get("合格点数"));
                                lmhpjcds += Double.valueOf(map.get("检测点数"));
                                lmyxpc = map.get("允许偏差");
                                a = true;

                            }else if(map.get("路面类型").contains("沥青路面") && map.get("检测项目").equals("连接线")){
                                ljxhgds += Double.valueOf(map.get("合格点数"));
                                ljxjcds += Double.valueOf(map.get("检测点数"));
                                ljxyxpc = map.get("允许偏差");
                                ljx = true;

                            } else if (map.get("路面类型").contains("沥青桥面")) {
                                qmhphgds += Double.valueOf(map.get("合格点数"));
                                qmhpjcds += Double.valueOf(map.get("检测点数"));
                                qmyxpc = map.get("允许偏差");
                                b = true;

                            } else if (map.get("路面类型").contains("沥青隧道")) {
                                sdhphgds += Double.valueOf(map.get("合格点数"));
                                sdhpjcds += Double.valueOf(map.get("检测点数"));
                                sdyxpc = map.get("允许偏差");
                                c = true;
                            } else if (map.get("路面类型").contains("混凝土路面")) {
                                hnthphgds += Double.valueOf(map.get("合格点数"));
                                hnthpjcds += Double.valueOf(map.get("检测点数"));
                                hntyxpc = map.get("允许偏差");
                                d = true;
                            }

                        }
                        if (a) {
                            Map<String, Object> maphplm = new HashMap<>();
                            maphplm.put("filename", "详见《沥青路面横坡质量鉴定表》检测" + decf.format(lmhpjcds) + "点,合格" + decf.format(lmhphgds) + "点");
                            maphplm.put("ccname", "横坡(沥青路面)");
                            maphplm.put("ccname2", "路面面层");
                            maphplm.put("ccname3", "沥青路面");
                            maphplm.put("yxps", lmyxpc);
                            maphplm.put("sheetname", "分部-路面");
                            maphplm.put("fbgc", "路面面层");
                            maphplm.put("检测点数", decf.format(lmhpjcds));
                            maphplm.put("合格点数", decf.format(lmhphgds));
                            maphplm.put("合格率", (lmhpjcds != 0) ? df.format(lmhphgds / lmhpjcds * 100) : "0");
                            resulthp.add(maphplm);
                        }
                        if (b) {
                            Map<String, Object> maphpqm = new HashMap<>();
                            maphpqm.put("filename", "详见《沥青桥面横坡质量鉴定表》检测" + decf.format(qmhpjcds) + "点,合格" + decf.format(qmhphgds) + "点");
                            maphpqm.put("ccname", "横坡(桥面系)");
                            maphpqm.put("ccname2", "桥面系");
                            maphpqm.put("ccname3", "沥青路面");
                            maphpqm.put("yxps", qmyxpc);
                            maphpqm.put("sheetname", "分部-路面");
                            maphpqm.put("fbgc", "路面面层");
                            maphpqm.put("检测点数", decf.format(qmhpjcds));
                            maphpqm.put("合格点数", decf.format(qmhphgds));
                            maphpqm.put("合格率", (qmhpjcds != 0) ? df.format(qmhphgds / qmhpjcds * 100) : "0");
                            resulthp.add(maphpqm);

                        }
                        if (c) {
                            Map<String, Object> maphpsd = new HashMap<>();
                            maphpsd.put("filename", "详见《沥青桥面横坡质量鉴定表》检测" + decf.format(sdhpjcds) + "点,合格" + decf.format(sdhphgds) + "点");
                            maphpsd.put("ccname", "横坡(隧道路面)");
                            maphpsd.put("ccname2", "隧道路面");
                            maphpsd.put("ccname3", "沥青路面");
                            maphpsd.put("yxps", sdyxpc);
                            maphpsd.put("sheetname", "分部-路面");
                            maphpsd.put("fbgc", "路面面层");
                            maphpsd.put("检测点数", decf.format(sdhpjcds));
                            maphpsd.put("合格点数", decf.format(sdhphgds));
                            maphpsd.put("合格率", (sdhpjcds != 0) ? df.format(sdhphgds / sdhpjcds * 100) : "0");
                            resulthp.add(maphpsd);


                        }
                        if (d) {
                            Map<String, Object> maphphnt = new HashMap<>();
                            maphphnt.put("filename", "详见《混凝土路面横坡质量鉴定表》检测" + decf.format(hnthpjcds) + "点,合格" + decf.format(hnthphgds) + "点");
                            maphphnt.put("ccname", "横坡(水泥路面)");
                            maphphnt.put("ccname2", "路面面层");
                            maphphnt.put("ccname3", "水泥混凝土");
                            maphphnt.put("yxps", hntyxpc);
                            maphphnt.put("sheetname", "分部-路面");
                            maphphnt.put("fbgc", "路面面层");
                            maphphnt.put("检测点数", decf.format(hnthpjcds));
                            maphphnt.put("合格点数", decf.format(hnthphgds));
                            maphphnt.put("合格率", (hnthpjcds != 0) ? df.format(hnthphgds / hnthpjcds * 100) : "0");
                            resulthp.add(maphphnt);

                        }
                        if(ljx){
                            Map<String, Object> maphplm = new HashMap<>();
                            maphplm.put("filename", "详见《沥青路面横坡质量鉴定表》检测" + decf.format(ljxjcds) + "点,合格" + decf.format(ljxhgds) + "点");
                            maphplm.put("ccname", "横坡(连接线)");
                            maphplm.put("ccname2", "路面面层");
                            maphplm.put("ccname3", "沥青路面");
                            maphplm.put("yxps", ljxyxpc);
                            maphplm.put("sheetname", "分部-路面");
                            maphplm.put("fbgc", "路面面层");
                            maphplm.put("检测点数", decf.format(ljxjcds));
                            maphplm.put("合格点数", decf.format(ljxhgds));
                            maphplm.put("合格率", (ljxjcds != 0) ? df.format(ljxhgds / ljxjcds * 100) : "0");
                            resulthp.add(maphplm);
                        }

                    }else {
                        /*
                         * 处理逻辑：先对list内的左右幅先加总，再进行分类
                         * 如何加总： 分部工程名字一样就加起来。 deepseek写的流式传输， 非常高级优雅
                         * */
                        list5 = list5.stream()
                                .collect(Collectors.groupingBy(m -> m.get("分部工程")))
                                .values()
                                .stream()
                                .map(group -> {
                                    Map<String, String> merged = new HashMap<>();
                                    int totalQualified = 0;
                                    int totalTested = 0;

                                    for (Map<String, String> item : group) {
                                        totalQualified += Integer.parseInt(item.get("合格点数"));
                                        totalTested += Integer.parseInt(item.get("检测点数"));
                                    }

                                    String rate = totalTested == 0 ? "0" : df.format((totalQualified * 1.0) / totalTested *100);
                                    merged.put("分部工程", group.get(0).get("分部工程"));
                                    merged.put("路面类型", group.get(0).get("路面类型"));
                                    merged.put("允许偏差", group.get(0).get("允许偏差"));
                                    merged.put("检测项目", group.get(0).get("检测项目"));
                                    merged.put("合格点数", String.valueOf(totalQualified));
                                    merged.put("检测点数", String.valueOf(totalTested));
                                    merged.put("合格率", rate);
                                    return merged;
                                })
                                .collect(Collectors.toList());
                        // 分类之前，还要把主线当中 路面 + 匝道 加起来，其他都不用加
                        int totalQualified = 0;
                        int totalChecked = 0;
                        String yxpc = "", fbgc = "", lmlx = "";
                        Iterator<Map<String, String>> iterator = list5.iterator();
                        while (iterator.hasNext()) {
                            Map<String, String> item = iterator.next();
                            // 检测项目过滤
                            String project = item.get("检测项目");
                            if (!("主线".equals(project) || "匝道".equals(project))) continue;

                            // 分部工程匹配（支持任意括号内容）
                            String pavementType = item.get("分部工程");
                            if (!pavementType.matches("路面面层\\(.*\\)") || pavementType == null ) continue;

                            // 把主线上的信息保留下来
                            if ("主线".equals(project)){
                                yxpc = item.get("允许偏差");
                                fbgc = item.get("分部工程");
                                lmlx = item.get("路面类型");
                            }

                            // 数值解析累加
                            try {
                                int qualified = Integer.parseInt(item.get("合格点数"));
                                int checked = Integer.parseInt(item.get("检测点数"));

                                totalQualified += qualified;
                                totalChecked += checked;
                            } catch (NumberFormatException | NullPointerException e) {
                                // 异常数据处理（根据业务需求处理）
                                continue;
                            }
                            // 把已经加过的map删去,遍历的时候删除得用迭代器
                            iterator.remove();
                        }
                        // 最后把算好的重新加入list5
                        String rate = totalChecked == 0 ? "0" : df.format((totalQualified * 1.0) / totalChecked *100);
                        Map<String, String> result = new HashMap<>();
                        result.put("允许偏差", yxpc);
                        result.put("分部工程", fbgc);
                        result.put("路面类型", lmlx);
                        result.put("检测项目", "主线");
                        result.put("检测点数", String.valueOf(totalChecked));
                        result.put("合格点数", String.valueOf(totalQualified));
                        result.put("合格率", rate);
                        list5.add(result);


                        for (Map<String, String> map : list5) {
                            // 初始化
                            lmhphgds = 0;
                            qmhphgds = 0;
                            sdhphgds = 0;
                            hnthphgds = 0;
                            ljxhgds = 0;

                            lmhpjcds = 0;
                            qmhpjcds = 0;
                            sdhpjcds = 0;
                            hnthpjcds = 0;
                            ljxjcds = 0;

                            lmyxpc = "";
                            sdyxpc = "";
                            qmyxpc = "";
                            hntyxpc = "";
                            ljxyxpc = "";

                            a = false;
                            b = false;
                            c = false;
                            d = false;
                            ljx = false;


                            if (map.get("路面类型").contains("沥青路面") && (map.get("检测项目").equals("主线") || map.get("检测项目").equals("匝道"))) {
                                lmhphgds = Double.valueOf(map.get("合格点数"));
                                lmhpjcds = Double.valueOf(map.get("检测点数"));
                                lmyxpc = map.get("允许偏差");
                                a = true;

                            } else if(map.get("路面类型").contains("沥青路面") && map.get("检测项目").equals("连接线")){
                                ljxhgds = Double.valueOf(map.get("合格点数"));
                                ljxjcds = Double.valueOf(map.get("检测点数"));
                                ljxyxpc = map.get("允许偏差");
                                ljx = true;

                            } else if (map.get("路面类型").contains("沥青桥面")) {
                                qmhphgds = Double.valueOf(map.get("合格点数"));
                                qmhpjcds = Double.valueOf(map.get("检测点数"));
                                qmyxpc = map.get("允许偏差");
                                b = true;

                            } else if (map.get("路面类型").contains("沥青隧道")) {
                                sdhphgds = Double.valueOf(map.get("合格点数"));
                                sdhpjcds = Double.valueOf(map.get("检测点数"));
                                sdyxpc = map.get("允许偏差");
                                c = true;
                            } else if (map.get("路面类型").contains("混凝土路面")) {
                                hnthphgds = Double.valueOf(map.get("合格点数"));
                                hnthpjcds = Double.valueOf(map.get("检测点数"));
                                hntyxpc = map.get("允许偏差");
                                d = true;
                            }


                            if (a) {
                                Map<String, Object> maphplm = new HashMap<>();
                                maphplm.put("filename", "详见《沥青路面横坡质量鉴定表》检测" + decf.format(lmhpjcds) + "点,合格" + decf.format(lmhphgds) + "点");
                                maphplm.put("ccname", "横坡(沥青路面)");
                                maphplm.put("ccname2", "路面面层");
                                maphplm.put("ccname3", "沥青路面");
                                maphplm.put("yxps", lmyxpc);
                                maphplm.put("sheetname", "分部-路面");
                                maphplm.put("fbgc", "路面面层");
                                maphplm.put("检测点数", decf.format(lmhpjcds));
                                maphplm.put("合格点数", decf.format(lmhphgds));
                                maphplm.put("合格率", (lmhpjcds != 0) ? df.format(lmhphgds / lmhpjcds * 100) : "0");
                                resulthp.add(maphplm);
                            }
                            else if (b) {
                                Map<String, Object> maphpqm = new HashMap<>();
                                maphpqm.put("filename", "详见《沥青桥面横坡质量鉴定表》检测" + decf.format(qmhpjcds) + "点,合格" + decf.format(qmhphgds) + "点");
                                maphpqm.put("ccname", "横坡(桥面系)");
                                maphpqm.put("ccname2", "桥面系");
                                maphpqm.put("ccname3", "沥青路面");
                                maphpqm.put("yxps", qmyxpc);
                                maphpqm.put("sheetname", "分部-"+getPlace(map.get("分部工程")));
                                maphpqm.put("fbgc", "路面面层");
                                maphpqm.put("检测点数", decf.format(qmhpjcds));
                                maphpqm.put("合格点数", decf.format(qmhphgds));
                                maphpqm.put("合格率", (qmhpjcds != 0) ? df.format(qmhphgds / qmhpjcds * 100) : "0");
                                resulthp.add(maphpqm);

                            }
                            else if (c) {
                                Map<String, Object> maphpsd = new HashMap<>();
                                maphpsd.put("filename", "详见《沥青桥面横坡质量鉴定表》检测" + decf.format(sdhpjcds) + "点,合格" + decf.format(sdhphgds) + "点");
                                maphpsd.put("ccname", "横坡(隧道路面)");
                                maphpsd.put("ccname2", "隧道路面");
                                maphpsd.put("ccname3", "沥青路面");
                                maphpsd.put("yxps", sdyxpc);
                                maphpsd.put("sheetname", "分部-"+getPlace(map.get("分部工程")));
                                maphpsd.put("fbgc", "路面面层");
                                maphpsd.put("检测点数", decf.format(sdhpjcds));
                                maphpsd.put("合格点数", decf.format(sdhphgds));
                                maphpsd.put("合格率", (sdhpjcds != 0) ? df.format(sdhphgds / sdhpjcds * 100) : "0");
                                resulthp.add(maphpsd);


                            }
                            else if (d) {
                                Map<String, Object> maphphnt = new HashMap<>();
                                maphphnt.put("filename", "详见《混凝土路面横坡质量鉴定表》检测" + decf.format(hnthpjcds) + "点,合格" + decf.format(hnthphgds) + "点");
                                maphphnt.put("ccname", "横坡(水泥路面)");
                                maphphnt.put("ccname2", "路面面层");
                                maphphnt.put("ccname3", "水泥混凝土");
                                maphphnt.put("yxps", hntyxpc);
                                maphphnt.put("sheetname", "分部-路面");
                                maphphnt.put("fbgc", "路面面层");
                                maphphnt.put("检测点数", decf.format(hnthpjcds));
                                maphphnt.put("合格点数", decf.format(hnthphgds));
                                maphphnt.put("合格率", (hnthpjcds != 0) ? df.format(hnthphgds / hnthpjcds * 100) : "0");
                                resulthp.add(maphphnt);

                            }else if(ljx){
                                Map<String, Object> maphplm = new HashMap<>();
                                maphplm.put("filename", "详见《沥青路面横坡质量鉴定表》检测" + decf.format(ljxjcds) + "点,合格" + decf.format(ljxhgds) + "点");
                                maphplm.put("ccname", "横坡(连接线)");
                                maphplm.put("ccname2", "路面面层");
                                maphplm.put("ccname3", "沥青路面");
                                maphplm.put("yxps", ljxyxpc);
                                maphplm.put("sheetname", "分部-路面");
                                maphplm.put("fbgc", "路面面层");
                                maphplm.put("检测点数", decf.format(ljxjcds));
                                maphplm.put("合格点数", decf.format(ljxhgds));
                                maphplm.put("合格率", (ljxjcds != 0) ? df.format(ljxhgds / ljxjcds * 100) : "0");
                                resulthp.add(maphplm);
                            }
                        }
                    }
                    resultlist.addAll(resulthp);
                    break;
                case "20路面构造深度.xlsx":
                    List<Map<String, Object>> list6 = jjgZdhGzsdService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultgzsd = new ArrayList<>();
                    double lmzds = 0;
                    double lmhgs = 0;
                    String lmsjz = "";
                    double sdzds = 0;
                    double sdhgd = 0;
                    String sdsjz = "";
                    double qzds = 0;
                    double qhgds = 0;
                    String qsjz = "";
                    boolean aa = false;
                    boolean bb = false;
                    boolean cc = false;
                    if(flag == 1){
                        for (Map<String, Object> map : list6) {
                            if (map.get("路面类型").toString().contains("路面")){
                                lmzds += Double.valueOf(map.get("总点数").toString());
                                lmhgs += Double.valueOf(map.get("合格点数").toString());
                                lmsjz = map.get("设计值").toString();
                                aa = true;
                            }
                            if (map.get("路面类型").toString().contains("隧道")){
                                sdzds += Double.valueOf(map.get("总点数").toString());
                                sdhgd += Double.valueOf(map.get("合格点数").toString());
                                sdsjz = map.get("设计值").toString();
                                bb =true;
                            }
                            if (map.get("路面类型").toString().contains("桥")){
                                qzds += Double.valueOf(map.get("总点数").toString());
                                qhgds += Double.valueOf(map.get("合格点数").toString());
                                qsjz = map.get("设计值").toString();
                                cc = true;
                            }
                        }
                        if (aa){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《沥青路面构造深度质量鉴定表》检测"+decf.format(lmzds)+"点,合格"+decf.format(lmhgs)+"点");
                            map.put("ccname", "构造深度");
                            map.put("ccname2", "路面面层");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds));
                            map.put("合格点数", decf.format(lmhgs));
                            map.put("合格率", (lmhgs != 0) ? df.format(lmhgs/lmzds*100) : "0");
                            resultgzsd.add(map);
                        }
                        if (bb){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《隧道路面构造深度质量鉴定表》检测"+decf.format(sdzds)+"点,合格"+decf.format(sdhgd)+"点");
                            map.put("ccname", "构造深度");
                            map.put("ccname2", "隧道路面");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", sdsjz);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(sdzds));
                            map.put("合格点数", decf.format(sdhgd));
                            map.put("合格率", (sdhgd != 0) ? df.format(sdhgd/sdzds*100) : "0");
                            resultgzsd.add(map);
                        }
                        if (cc){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《桥面系构造深度质量鉴定表》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                            map.put("ccname", "构造深度");
                            map.put("ccname2", "桥面系");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", qsjz);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(qzds));
                            map.put("合格点数", decf.format(qhgds));
                            map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                            resultgzsd.add(map);
                        }
                    }else if(flag == 2){
                        list6 = processJgMap(list6); // 复用的处理方法，所以导致路面的最大最小值没有，但是没有关系，暂时用不上
                        for(Map<String, Object> map1 : list6){
                            // 初始化
                            lmzds = 0;
                            lmhgs = 0;
                            lmsjz = "";

                            lmzds = Double.valueOf(map1.get("总点数").toString());
                            lmhgs = Double.valueOf(map1.get("合格点数").toString());
                            lmsjz = map1.get("设计值").toString();

                            // 填写数据
                            Map<String, Object> map = new HashMap<>();
                            // 分情况
                            if (map1.get("路面类型").toString().contains("路面")){
                                if(map1.get("检测项目").toString().contains("连接线")) {
                                    map.put("filename","详见《连接线沥青路面构造深度质量鉴定表》检测"+decf.format(lmzds)+"点,合格"+decf.format(lmhgs)+"点");
                                    map.put("ccname", "构造深度（连接线）");
                                }else {
                                    map.put("filename","详见《沥青路面构造深度质量鉴定表》检测"+decf.format(lmzds)+"点,合格"+decf.format(lmhgs)+"点");
                                    map.put("ccname", "构造深度");
                                }
                                map.put("ccname2", "路面面层");
                                map.put("sheetname", "分部-路面");
                            }else if (map1.get("路面类型").toString().contains("隧道")){
                                map.put("filename","详见《隧道路面构造深度质量鉴定表》检测"+decf.format(lmzds)+"点,合格"+decf.format(lmhgs)+"点");
                                map.put("ccname", "构造深度");
                                map.put("ccname2", "隧道路面");
                                map.put("sheetname", "分部-"+ map1.get("分部工程名称").toString());
                            }else if (map1.get("路面类型").toString().contains("桥")){
                                map.put("filename","详见《桥面系构造深度质量鉴定表》检测"+decf.format(lmzds)+"点,合格"+decf.format(lmhgs)+"点");
                                map.put("ccname", "构造深度");
                                map.put("ccname2", "桥面系");
                                map.put("sheetname", "分部-"+ map1.get("分部工程名称").toString());
                            }



                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz);
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds));
                            map.put("合格点数", decf.format(lmhgs));
                            map.put("合格率", (lmhgs != 0) ? df.format(lmhgs/lmzds*100) : "0");
                            resultgzsd.add(map);
                        }
                    }

                    resultlist.addAll(resultgzsd);
                    break;
                case "19路面摩擦系数.xlsx":
                    //分工作簿
                    List<Map<String, Object>> list7 = jjgZdhMcxsService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultmcxs = new ArrayList<>();
                    double lmzds2 = 0;
                    double lmhgs2 = 0;
                    String lmsjz2 = "";
                    double sdzds2 = 0;
                    double sdhgd2 = 0;
                    String sdsjz2 = "";
                    double qzds2 = 0;
                    double qhgds2 = 0;
                    String qsjz2 = "";
                    boolean aaa = false;
                    boolean bbb = false;
                    boolean ccc = false;
                    if(flag == 1){
                        for (Map<String, Object> map : list7) {
                            if (map.get("路面类型").toString().contains("路面")){
                                lmzds2 += Double.valueOf(map.get("总点数").toString());
                                lmhgs2 += Double.valueOf(map.get("合格点数").toString());
                                lmsjz2 = map.get("设计值").toString();
                                aaa = true;
                            }
                            if (map.get("路面类型").toString().contains("隧道")){
                                sdzds2 += Double.valueOf(map.get("总点数").toString());
                                sdhgd2 += Double.valueOf(map.get("合格点数").toString());
                                sdsjz2 = map.get("设计值").toString();
                                bbb =true;
                            }
                            if (map.get("路面类型").toString().contains("桥")){
                                qzds2 += Double.valueOf(map.get("总点数").toString());
                                qhgds2 += Double.valueOf(map.get("合格点数").toString());
                                qsjz2 = map.get("设计值").toString();
                                ccc = true;
                            }
                        }
                        if (aaa){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《沥青路面摩擦系数质量鉴定表》检测"+decf.format(lmzds2)+"点,合格"+decf.format(lmhgs2)+"点");
                            map.put("ccname", "摩擦系数");
                            map.put("ccname2", "路面面层");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz2);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds2));
                            map.put("合格点数", decf.format(lmhgs2));
                            map.put("合格率", (lmhgs2 != 0) ? df.format(lmhgs2/lmzds2*100) : "0");
                            resultmcxs.add(map);
                        }
                        if (bbb){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《隧道路面摩擦系数质量鉴定表》检测"+decf.format(sdzds2)+"点,合格"+decf.format(sdhgd2)+"点");
                            map.put("ccname", "摩擦系数");
                            map.put("ccname2", "隧道路面");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", sdsjz2);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(sdzds2));
                            map.put("合格点数", decf.format(sdhgd2));
                            map.put("合格率", (sdhgd2 != 0) ? df.format(sdhgd2/sdzds2*100) : "0");
                            resultmcxs.add(map);
                        }
                        if (ccc){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《桥面系摩擦系数质量鉴定表》检测"+decf.format(qzds2)+"点,合格"+decf.format(qhgds2)+"点");
                            map.put("ccname", "摩擦系数");
                            map.put("ccname2", "桥面系");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", qsjz2);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(qzds2));
                            map.put("合格点数", decf.format(qhgds2));
                            map.put("合格率", (qhgds2 != 0) ? df.format(qhgds2/qzds2*100) : "0");
                            resultmcxs.add(map);
                        }
                    }else if(flag == 2){
                        // 对list7进行分组，根据路面类型分组
                        list7 = processJgMap(list7);
                        for(Map<String, Object> map1 : list7){
                            // 初始化
                            lmzds2 = 0;
                            lmhgs2 = 0;
                            lmsjz2 = "";

                            lmzds2 = Double.valueOf(map1.get("总点数").toString());
                            lmhgs2 = Double.valueOf(map1.get("合格点数").toString());
                            lmsjz2 = map1.get("设计值").toString();

                            // 填写数据
                            Map<String, Object> map = new HashMap<>();
                            // 分情况
                            if (map1.get("路面类型").toString().contains("路面")){
                                if(map1.get("检测项目").toString().contains("连接线")) {
                                    map.put("filename","详见《沥青路面连接线摩擦系数质量鉴定表》检测"+decf.format(lmzds2)+"点,合格"+decf.format(lmhgs2)+"点");
                                    map.put("ccname", "摩擦系数(连接线)");
                                }else {
                                    map.put("filename","详见《沥青路面摩擦系数质量鉴定表》检测"+decf.format(lmzds2)+"点,合格"+decf.format(lmhgs2)+"点");
                                    map.put("ccname", "摩擦系数");
                                }
                                map.put("ccname2", "路面面层");
                                map.put("sheetname", "分部-路面");
                            }else if (map1.get("路面类型").toString().contains("隧道")){
                                map.put("filename","详见《隧道路面摩擦系数质量鉴定表》检测"+decf.format(lmzds2)+"点,合格"+decf.format(lmhgs2)+"点");
                                map.put("ccname", "摩擦系数");
                                map.put("ccname2", "隧道路面");
                                map.put("sheetname", "分部-"+ map1.get("分部工程名称").toString());
                            }else if (map1.get("路面类型").toString().contains("桥")){
                                map.put("filename","详见《桥面系摩擦系数质量鉴定表》检测"+decf.format(lmzds2)+"点,合格"+decf.format(lmhgs2)+"点");
                                map.put("ccname", "摩擦系数");
                                map.put("ccname2", "桥面系");
                                map.put("sheetname", "分部-"+ map1.get("分部工程名称").toString());
                            }



                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz2);
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds2));
                            map.put("合格点数", decf.format(lmhgs2));
                            map.put("合格率", (lmhgs2 != 0) ? df.format(lmhgs2/lmzds2*100) : "0");
                            resultmcxs.add(map);
                        }
                    }

                    resultlist.addAll(resultmcxs);
                    break;
                case "18路面平整度.xlsx":
                    List<Map<String, Object>> list8 = jjgZdhPzdService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultpzd = new ArrayList<>();
                    double lmzds3 = 0;
                    double lmhgs3 = 0;
                    String lmsjz3 = "";
                    double sdzds3 = 0;
                    double sdhgd3 = 0;
                    String sdsjz3 = "";
                    double qzds3 = 0;
                    double qhgds3 = 0;
                    String qsjz3 = "";
                    boolean aaaa = false;
                    boolean bbbb = false;
                    boolean cccc = false;
                    if(flag == 1){
                        for (Map<String, Object> map : list8) {
                            if (map.get("路面类型").toString().equals("沥青路面") || map.get("路面类型").toString().equals("沥青匝道")){
                                lmzds3 += Double.valueOf(map.get("总点数").toString());
                                lmhgs3 += Double.valueOf(map.get("合格点数").toString());
                                lmsjz3 = map.get("设计值").toString();
                                aaaa = true;
                            }
                            if (map.get("路面类型").toString().contains("隧道")){
                                sdzds3 += Double.valueOf(map.get("总点数").toString());
                                sdhgd3 += Double.valueOf(map.get("合格点数").toString());
                                sdsjz3 = map.get("设计值").toString();
                                bbbb =true;
                            }
                            if (map.get("路面类型").toString().contains("桥")){
                                qzds3 += Double.valueOf(map.get("总点数").toString());
                                qhgds3 += Double.valueOf(map.get("合格点数").toString());
                                qsjz3 = map.get("设计值").toString();
                                cccc = true;
                            }
                        }
                        if (aaaa){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《沥青路面平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                            map.put("ccname", "平整度");
                            map.put("ccname2", "路面面层");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz3);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds3));
                            map.put("合格点数", decf.format(lmhgs3));
                            map.put("合格率", (lmhgs3 != 0) ? df.format(lmhgs3/lmzds3*100) : "0");
                            resultpzd.add(map);
                        }
                        if (bbbb){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《隧道路面平整度质量鉴定表》检测"+decf.format(sdzds3)+"点,合格"+decf.format(sdhgd3)+"点");
                            map.put("ccname", "平整度");
                            map.put("ccname2", "隧道路面");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", sdsjz3);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(sdzds3));
                            map.put("合格点数", decf.format(sdhgd3));
                            map.put("合格率", (sdhgd3 != 0) ? df.format(sdhgd3/sdzds3*100) : "0");
                            resultpzd.add(map);
                        }
                        if (cccc){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《桥面系平整度质量鉴定表》检测"+decf.format(qzds3)+"点,合格"+decf.format(qhgds3)+"点");
                            map.put("ccname", "平整度");
                            map.put("ccname2", "桥面系");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", qsjz3);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(qzds3));
                            map.put("合格点数", decf.format(qhgds3));
                            map.put("合格率", (qhgds3 != 0) ? df.format(qhgds3/qzds3*100) : "0");
                            resultpzd.add(map);
                        }
                    }else if(flag == 2){
                        // 先进行组合
                        list8 = processPZDList(list8);
                        for (Map<String, Object> map8 : list8) {
                            lmzds3 = 0;
                            lmhgs3 = 0;



                            lmzds3 += Double.valueOf(map8.get("总点数").toString());
                            lmhgs3 += Double.valueOf(map8.get("合格点数").toString());
                            lmsjz3 = map8.get("设计值").toString();

                            Map<String, Object> map = new HashMap<>();


                            if (map8.get("路面类型").toString().equals("沥青路面")){

                                if(map8.get("检测项目").toString().contains("连接线")) {
                                    map.put("filename","详见《连接线路面平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                    map.put("ccname", "平整度（连接线）");
                                }
                                else {
                                    map.put("filename","详见《沥青路面平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                    map.put("ccname", "平整度");
                                }
                                map.put("ccname2", "路面面层");
                                map.put("ccname3", "沥青路面");
                                map.put("sheetname", "分部-路面");
                            }else if (map8.get("路面类型").toString().contains("隧道")){
                                if(map8.get("检测项目").toString().contains("连接线")) {
                                    map.put("filename","详见《连接线隧道平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                    map.put("ccname", "平整度（连接线）");
                                }
                                else {
                                    map.put("filename","详见《隧道路面平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                    map.put("ccname", "平整度");
                                }
                                map.put("ccname2", "隧道路面");
                                map.put("ccname3", "沥青路面");
                                map.put("sheetname", "分部-" + map8.get("分部工程名称").toString());
                            }else if (map8.get("路面类型").toString().contains("桥")){
                                if(map8.get("检测项目").toString().contains("连接线")) {
                                    map.put("filename","详见《连接线桥面系平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                    map.put("ccname", "平整度（连接线）");
                                }
                                else {
                                    map.put("filename","详见《桥面系平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                    map.put("ccname", "平整度");
                                }
                                map.put("ccname2", "桥面系");
                                map.put("ccname3", "沥青路面");
                                map.put("sheetname", "分部-" + map8.get("分部工程名称").toString());
                            }else if (map8.get("路面类型").toString().contains("收费站")){ // 该情况下只有收费站有混凝土

                                map.put("filename","详见《混凝土路面平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                                map.put("ccname", "平整度");

                                map.put("ccname2", "路面面层");
                                map.put("ccname3", "混凝土路面");
                                map.put("sheetname", "分部-路面");
                            }


                            map.put("yxps", lmsjz3);
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds3));
                            map.put("合格点数", decf.format(lmhgs3));
                            map.put("合格率", (lmhgs3 != 0) ? df.format(lmhgs3/lmzds3*100) : "0");
                            resultpzd.add(map);
                        }
                    }

                    resultlist.addAll(resultpzd);

                    break;
                case "16路面雷达厚度.xlsx":
                    List<Map<String, Object>> list9 = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultldhd = new ArrayList<>();
                    double lmzds4 = 0;
                    double lmhgs4 = 0;
                    String lmsjz4 = "";
                    double sdzds4 = 0;
                    double sdhgd4 = 0;
                    String sdsjz4 = "";
                    double qzds4 = 0;
                    double qhgds4 = 0;
                    String qsjz4 = "";
                    boolean aaaaaa = false;
                    boolean bbbbbb = false;
                    boolean cccccc = false;

                    // 合并则用原来代码
                    if(flag == 1){
                        for (Map<String, Object> map : list9) {
                            if (map.get("路面类型").toString().contains("右幅") || map.get("路面类型").toString().contains("左幅")){
                                lmzds4 += Double.valueOf(map.get("总点数").toString());
                                lmhgs4 += Double.valueOf(map.get("合格点数").toString());
                                lmsjz4 = map.get("设计值").toString();
                                aaaaaa = true;
                            }
                            if (map.get("路面类型").toString().contains("隧道")){
                                sdzds4 += Double.valueOf(map.get("总点数").toString());
                                sdhgd4 += Double.valueOf(map.get("合格点数").toString());
                                sdsjz4 = map.get("设计值").toString();
                                bbbbbb =true;
                            }
                            if (map.get("路面类型").toString().contains("桥")){
                                qzds4 += Double.valueOf(map.get("总点数").toString());
                                qhgds4 += Double.valueOf(map.get("合格点数").toString());
                                qsjz4 = map.get("设计值").toString();
                                cccccc = true;
                            }
                        }
                        if (aaaaaa){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《路面厚度质量鉴定表（雷达法）》检测"+decf.format(lmzds4)+"点,合格"+decf.format(lmhgs4)+"点");
                            map.put("ccname", "△雷达厚度");
                            map.put("ccname2", "路面面层");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz4);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds4));
                            map.put("合格点数", decf.format(lmhgs4));
                            map.put("合格率", (lmhgs4 != 0) ? df.format(lmhgs4/lmzds4*100) : "0");
                            resultldhd.add(map);
                        }
                        if (bbbbbb){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《隧道路面厚度质量鉴定表（雷达法）》检测"+decf.format(sdzds4)+"点,合格"+decf.format(sdhgd4)+"点");
                            map.put("ccname", "△雷达厚度");
                            map.put("ccname2", "隧道路面");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", sdsjz4);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(sdzds4));
                            map.put("合格点数", decf.format(sdhgd4));
                            map.put("合格率", (sdhgd4 != 0) ? df.format(sdhgd4/sdzds4*100) : "0");
                            resultldhd.add(map);
                        }
                        if (cccccc){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《桥面厚度质量鉴定表（雷达法）》检测"+decf.format(qzds4)+"点,合格"+decf.format(qhgds4)+"点");
                            map.put("ccname", "△雷达厚度");
                            map.put("ccname2", "桥面系");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", qsjz4);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(qzds4));
                            map.put("合格点数", decf.format(qhgds4));
                            map.put("合格率", (qhgds4 != 0) ? df.format(qhgds4/qzds4*100) : "0");
                            resultldhd.add(map);
                        }
                        resultlist.addAll(resultldhd);
                    }else{
                        /* 处理逻辑：在原来的基础上，加上工程名称的进一步排查
                        *
                        *   路面左右幅： 路面类型包含左右幅，就可以加在一起，但是要排除连接线（检测项目不含“连接线”）
                        *   桥：直接拿，不用加匝道
                        *   隧道: 也是直接拿
                        *   连接线：单独处理
                        * */

                        // 分类之前，还要把路面左右幅加起来， 连接线左右幅也单独处理下
                        int totalQualified = 0;
                        int totalChecked = 0;
                        int ljxtotalQualified = 0;
                        int ljxtotalChecked = 0;
                        String yxpc = "", jcxm = "主线", lmlx = "左幅", sjz = ""; // 检测项目 路面类型 直接写死，并且舍弃最大最小值，因为根本没用上
                        String ljxsjz = "";
                        Iterator<Map<String, Object>> iterator = list9.iterator();
                        while (iterator.hasNext()) {
                            Map<String, Object> item = iterator.next();
                            // 路面类型包含左右幅
                            if((item.get("路面类型").toString().contains("右幅") || item.get("路面类型").toString().contains("左幅"))){
                                // 设计值就用主线的
                                if(item.get("检测项目").toString().contains("主线")){
                                    sjz = item.get("设计值").toString();
                                }
                                // 数值解析累加
                                try {
                                    int qualified = Integer.parseInt(item.get("合格点数").toString());
                                    int checked = Integer.parseInt(item.get("总点数").toString());

                                    if(item.get("检测项目").toString().contains("连接线")){
                                        ljxtotalQualified += qualified;
                                        ljxtotalChecked += checked;
                                        ljxsjz = item.get("设计值").toString();
                                    }else{
                                        totalQualified += qualified;
                                        totalChecked += checked;
                                    }

                                } catch (NumberFormatException | NullPointerException e) {
                                    // 异常数据处理（根据业务需求处理）
                                    continue;
                                }
                                // 把已经加过的map删去,遍历的时候删除得用迭代器
                                iterator.remove();

                            }
                        }
                        // 最后把算好的重新加入list9
                        String rate = totalChecked == 0 ? "0" : df.format((totalQualified * 1.0) / totalChecked *100);
                        Map<String, Object> result = new HashMap<>();
                        result.put("设计值", sjz);
                        result.put("路面类型", lmlx);
                        result.put("检测项目", jcxm);
                        result.put("总点数", String.valueOf(totalChecked));
                        result.put("合格点数", String.valueOf(totalQualified));
                        result.put("合格率", rate);
                        list9.add(result);
                        if(ljxtotalChecked != 0){ // 如果也有连接线
                            String ljxrate = ljxtotalChecked == 0 ? "0" : df.format((ljxtotalQualified * 1.0) / ljxtotalChecked *100);
                            Map<String, Object> ljxresult = new HashMap<>();
                            ljxresult.put("设计值", ljxsjz);
                            ljxresult.put("路面类型", "左幅");
                            ljxresult.put("检测项目", "连接线"); //不用管名字
                            ljxresult.put("总点数", String.valueOf(ljxtotalChecked));
                            ljxresult.put("合格点数", String.valueOf(ljxtotalQualified));
                            ljxresult.put("合格率", ljxrate);
                            list9.add(ljxresult);
                        }

                        // 路面
                        for (Map<String, Object> mapf : list9) {
                            // 初始化
                            lmzds4 = 0;
                            lmhgs4 = 0;
                            lmsjz4 = "";
                            sdzds4 = 0;
                            sdhgd4 = 0;
                            sdsjz4 = "";
                            qzds4 = 0;
                            qhgds4 = 0;
                            qsjz4 = "";
                            aaaaaa = false;
                            bbbbbb = false;
                            cccccc = false;

                            if (mapf.get("路面类型").toString().contains("右幅") || mapf.get("路面类型").toString().contains("左幅")){
                                lmzds4 += Double.valueOf(mapf.get("总点数").toString());
                                lmhgs4 += Double.valueOf(mapf.get("合格点数").toString());
                                lmsjz4 = mapf.get("设计值").toString();
                                aaaaaa = true;
                            }
                            if (mapf.get("路面类型").toString().contains("隧道")){
                                sdzds4 += Double.valueOf(mapf.get("总点数").toString());
                                sdhgd4 += Double.valueOf(mapf.get("合格点数").toString());
                                sdsjz4 = mapf.get("设计值").toString();
                                bbbbbb =true;
                            }
                            if (mapf.get("路面类型").toString().contains("桥")){
                                qzds4 += Double.valueOf(mapf.get("总点数").toString());
                                qhgds4 += Double.valueOf(mapf.get("合格点数").toString());
                                qsjz4 = mapf.get("设计值").toString();
                                cccccc = true;
                            }

                            // 在里面放入连接线的处理
                            if (aaaaaa){
                                Map<String, Object> map = new HashMap<>();

                                if(mapf.get("检测项目").toString().contains("连接线")){
                                    map.put("filename","详见《连接线路面厚度质量鉴定表（雷达法）》检测"+decf.format(lmzds4)+"点,合格"+decf.format(lmhgs4)+"点");
                                    map.put("ccname", "△雷达厚度(连接线路面)");
                                }else{
                                    map.put("filename","详见《路面厚度质量鉴定表（雷达法）》检测"+decf.format(lmzds4)+"点,合格"+decf.format(lmhgs4)+"点");
                                    map.put("ccname", "△雷达厚度");
                                }
                                map.put("ccname2", "路面面层");
                                map.put("ccname3", "沥青路面");
                                map.put("yxps", lmsjz4);
                                map.put("sheetname", "分部-路面");
                                map.put("fbgc", "路面面层");
                                map.put("检测点数", decf.format(lmzds4));
                                map.put("合格点数", decf.format(lmhgs4));
                                map.put("合格率", (lmhgs4 != 0) ? df.format(lmhgs4/lmzds4*100) : "0");
                                resultldhd.add(map);
                            }
                            if (bbbbbb){
                                Map<String, Object> map = new HashMap<>();
                                map.put("filename","详见《隧道路面厚度质量鉴定表（雷达法）》检测"+decf.format(sdzds4)+"点,合格"+decf.format(sdhgd4)+"点");
                                map.put("ccname", "△雷达厚度");
                                map.put("ccname2", "隧道路面");
                                map.put("ccname3", "沥青路面");
                                map.put("yxps", sdsjz4);
                                map.put("sheetname", "分部-"+mapf.get("工程名称").toString());
                                map.put("fbgc", "路面面层");
                                map.put("检测点数", decf.format(sdzds4));
                                map.put("合格点数", decf.format(sdhgd4));
                                map.put("合格率", (sdhgd4 != 0) ? df.format(sdhgd4/sdzds4*100) : "0");
                                resultldhd.add(map);
                            }
                            if (cccccc){
                                Map<String, Object> map = new HashMap<>();
                                map.put("filename","详见《桥面厚度质量鉴定表（雷达法）》检测"+decf.format(qzds4)+"点,合格"+decf.format(qhgds4)+"点");
                                map.put("ccname", "△雷达厚度");
                                map.put("ccname2", "桥面系");
                                map.put("ccname3", "沥青路面");
                                map.put("yxps", qsjz4);
                                map.put("sheetname", "分部-"+mapf.get("工程名称").toString());
                                map.put("fbgc", "路面面层");
                                map.put("检测点数", decf.format(qzds4));
                                map.put("合格点数", decf.format(qhgds4));
                                map.put("合格率", (qhgds4 != 0) ? df.format(qhgds4/qzds4*100) : "0");
                                resultldhd.add(map);
                            }
                        }

                        resultlist.addAll(resultldhd);
                    }

                    break;
                case "14路面车辙.xlsx":
                    List<Map<String, Object>> list10 = jjgZdhCzService.lookJdbjg(commonInfoVo, flag);
                    List<Map<String, Object>> resultcz = new ArrayList<>();
                    double lmzds5 = 0;
                    double lmhgs5 = 0;
                    String lmsjz5 = "";
                    double sdzds5 = 0;
                    double sdhgd5 = 0;
                    String sdsjz5 = "";
                    double qzds5 = 0;
                    double qhgds5 = 0;
                    String qsjz5 = "";
                    boolean aaaaa = false;
                    boolean bbbbb = false;
                    boolean ccccc = false;

                    if (flag == 1){
                        for (Map<String, Object> map : list10) {
                            if (map.get("路面类型").toString().contains("路面")){
                                lmzds5 += Double.valueOf(map.get("总点数").toString());
                                lmhgs5 += Double.valueOf(map.get("合格点数").toString());
                                lmsjz5 = map.get("设计值").toString();
                                aaaaa = true;
                            }
                            if (map.get("路面类型").toString().contains("隧道")){
                                sdzds5 += Double.valueOf(map.get("总点数").toString());
                                sdhgd5 += Double.valueOf(map.get("合格点数").toString());
                                sdsjz5 = map.get("设计值").toString();
                                bbbbb =true;
                            }
                            if (map.get("路面类型").toString().contains("桥")){
                                qzds5 += Double.valueOf(map.get("总点数").toString());
                                qhgds5 += Double.valueOf(map.get("合格点数").toString());
                                qsjz5 = map.get("设计值").toString();
                                ccccc = true;
                            }
                        }
                        if (aaaaa){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《沥青路面车辙质量鉴定表》检测"+decf.format(lmzds5)+"点,合格"+decf.format(lmhgs5)+"点");
                            map.put("ccname", "车辙");
                            map.put("ccname2", "路面面层");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", lmsjz5);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(lmzds5));
                            map.put("合格点数", decf.format(lmhgs5));
                            map.put("合格率", (lmhgs5 != 0) ? df.format(lmhgs5/lmzds5*100) : "0");
                            resultcz.add(map);
                        }
                        if (bbbbb){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《隧道路面车辙质量鉴定表》检测"+decf.format(sdzds5)+"点,合格"+decf.format(sdhgd5)+"点");
                            map.put("ccname", "车辙");
                            map.put("ccname2", "隧道路面");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", sdsjz5);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(sdzds5));
                            map.put("合格点数", decf.format(sdhgd5));
                            map.put("合格率", (sdhgd5 != 0) ? df.format(sdhgd5/sdzds5*100) : "0");
                            resultcz.add(map);
                        }
                        if (ccccc){
                            Map<String, Object> map = new HashMap<>();
                            map.put("filename","详见《桥面系车辙质量鉴定表》检测"+decf.format(qzds5)+"点,合格"+decf.format(qhgds5)+"点");
                            map.put("ccname", "车辙");
                            map.put("ccname2", "桥面系");
                            map.put("ccname3", "沥青路面");
                            map.put("yxps", qsjz5);
                            map.put("sheetname", "分部-路面");
                            map.put("fbgc", "路面面层");
                            map.put("检测点数", decf.format(qzds5));
                            map.put("合格点数", decf.format(qhgds5));
                            map.put("合格率", (qhgds5 != 0) ? df.format(qhgds5/qzds5*100) : "0");
                            resultcz.add(map);
                        }
                    }else if (flag == 2){
                        // 对List10先进行组合，按照要求把需要合在一起的合在一起
                        list10 = processJgMap(list10);
                        for (Map<String, Object> mapf : list10) {
                            lmzds5 = 0;
                            lmhgs5 = 0;
                            lmsjz5 = "";
                            sdzds5 = 0;
                            sdhgd5 = 0;
                            sdsjz5 = "";
                            qzds5 = 0;
                            qhgds5 = 0;
                            qsjz5 = "";
                            aaaaa = false;
                            bbbbb = false;
                            ccccc = false;
                            if (mapf.get("路面类型").toString().contains("路面")){
                                lmzds5 += Double.valueOf(mapf.get("总点数").toString());
                                lmhgs5 += Double.valueOf(mapf.get("合格点数").toString());
                                lmsjz5 = mapf.get("设计值").toString();

                                Map<String, Object> map = new HashMap<>();

                                map.put("ccname2", "路面面层");
                                map.put("ccname3", "沥青路面");
                                map.put("yxps", lmsjz5);
                                map.put("sheetname", "分部-路面");
                                map.put("fbgc", "路面面层");
                                map.put("检测点数", decf.format(lmzds5));
                                map.put("合格点数", decf.format(lmhgs5));
                                map.put("合格率", (lmhgs5 != 0) ? df.format(lmhgs5/lmzds5*100) : "0");

                                if(mapf.get("检测项目").toString().contains("连接线")){
                                    map.put("filename","详见《沥青路面连接线车辙质量鉴定表》检测"+decf.format(lmzds5)+"点,合格"+decf.format(lmhgs5)+"点");
                                    map.put("ccname", "车辙（连接线）");
                                }else{
                                    map.put("filename","详见《沥青路面车辙质量鉴定表》检测"+decf.format(lmzds5)+"点,合格"+decf.format(lmhgs5)+"点");
                                    map.put("ccname", "车辙");
                                }
                                resultcz.add(map);

                            }
                            if (mapf.get("路面类型").toString().contains("隧道")){
                                sdzds5 += Double.valueOf(mapf.get("总点数").toString());
                                sdhgd5 += Double.valueOf(mapf.get("合格点数").toString());
                                sdsjz5 = mapf.get("设计值").toString();

                                Map<String, Object> map = new HashMap<>();
                                map.put("filename","详见《隧道路面车辙质量鉴定表》检测"+decf.format(sdzds5)+"点,合格"+decf.format(sdhgd5)+"点");
                                map.put("ccname", "车辙");
                                map.put("ccname2", "隧道路面");
                                map.put("ccname3", "沥青路面");
                                map.put("yxps", sdsjz5);
                                map.put("sheetname", "分部-" + mapf.get("分部工程名称").toString());
                                map.put("fbgc", "路面面层");
                                map.put("检测点数", decf.format(sdzds5));
                                map.put("合格点数", decf.format(sdhgd5));
                                map.put("合格率", (sdhgd5 != 0) ? df.format(sdhgd5/sdzds5*100) : "0");
                                resultcz.add(map);
                            }
                            if (mapf.get("路面类型").toString().contains("桥")){
                                qzds5 += Double.valueOf(mapf.get("总点数").toString());
                                qhgds5 += Double.valueOf(mapf.get("合格点数").toString());
                                qsjz5 = mapf.get("设计值").toString();

                                Map<String, Object> map = new HashMap<>();
                                map.put("filename","详见《桥面系车辙质量鉴定表》检测"+decf.format(qzds5)+"点,合格"+decf.format(qhgds5)+"点");
                                map.put("ccname", "车辙");
                                map.put("ccname2", "桥面系");
                                map.put("ccname3", "沥青路面");
                                map.put("yxps", qsjz5);
                                map.put("sheetname", "分部-" + mapf.get("分部工程名称").toString());
                                map.put("fbgc", "路面面层");
                                map.put("检测点数", decf.format(qzds5));
                                map.put("合格点数", decf.format(qhgds5));
                                map.put("合格率", (qhgds5 != 0) ? df.format(qhgds5/qzds5*100) : "0");
                                resultcz.add(map);
                            }

                        }

                    }

                    resultlist.addAll(resultcz);
                    break;
                case "38隧道衬砌砼强度.xlsx":
                    List<Map<String, Object>> list18 = jjgFbgcSdgcCqtqdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> sdqdlist = new ArrayList<>();
                    for (Map<String, Object> map : list18) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","*衬砌强度");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《隧道衬砌强度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "衬砌");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        sdqdlist.add(map5);
                    }
                    resultlist.addAll(sdqdlist);
                    break;
                case "40隧道大面平整度.xlsx":
                    List<Map<String, Object>> list19 = jjgFbgcSdgcDmpzdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> dmlist = new ArrayList<>();
                    for (Map<String, Object> map : list19) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","大面平整度");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《隧道大面平整度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "衬砌");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        dmlist.add(map5);
                    }
                    resultlist.addAll(dmlist);
                    break;
                case "41隧道总体宽度.xlsx":
                    List<Map<String, Object>> list20 = jjgFbgcSdgcZtkdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> kdlist = new ArrayList<>();
                    for (Map<String, Object> map : list20) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","宽度");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《隧道总体宽度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "总体");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        kdlist.add(map5);
                    }
                    resultlist.addAll(kdlist);
                    break;
                default:
                    // 默认操作
                    break;
            }
            if (value.contains("12沥青路面压实度-")){
                // 防止重复添加
                if(ysd == true) continue;
                else ysd = true;

                //分工作簿
                List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjgbgzbg(commonInfoVo, flag);
                List<Map<String, Object>> resultysd = new ArrayList<>();
                double sdjcds = 0;
                double sdhgds = 0;
                double lmjcds = 0;
                double lmhgds = 0;
                String sdgdz="";
                String lmgdz="";
                 // 和之前不一样的是，这里不用再加了，因为在返回List的函数里面都已经加好了
                // 所以主要就是判断是什么项


                // 判断“路面类型”即可
                for (Map<String, Object> map : list) {
                    Map<String, Object> newMap1 = new HashMap<>();
                    //不分离的
                    String lmlx = map.get("路面类型").toString();
                    String mcxx = ""; //面层信息
                    if (lmlx.contains("-")){
                        mcxx = lmlx.split("-")[1]; //面层信息
                    }
                    if(lmlx.contains("沥青路面压实度") && !lmlx.contains("沥青路面压实度匝道")){
                        newMap1.put("filename","详见《沥青路面压实度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap1.put("ccname", "△沥青路面压实度(主线"+ mcxx +")");
                        newMap1.put("ccname2", "路面面层");
                        newMap1.put("sheetname", "分部-路面");
                    }else if(lmlx.contains("沥青路面压实度匝道")){
                        newMap1.put("filename","详见《沥青路面压实度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap1.put("ccname", "△沥青路面压实度(匝道"+ mcxx +")");
                        newMap1.put("ccname2", "路面面层");
                        newMap1.put("sheetname", "分部-路面");
                    }else if(lmlx.contains("隧道")){
                        newMap1.put("filename","详见《隧道沥青路面压实度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap1.put("ccname", "△沥青路面压实度(隧道"+ mcxx +")");
                        newMap1.put("ccname2", "隧道路面");
                        if(flag == 1)  newMap1.put("sheetname", "分部-路面");
                        else if(flag == 2) newMap1.put("sheetname", "分部-" + map.get("分部工程名称"));
                    }else if(lmlx.contains("连接线")){
                        newMap1.put("filename","详见《连接线沥青路面压实度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap1.put("ccname", "△沥青路面压实度(连接线"+ mcxx +")");
                        newMap1.put("ccname2", "连接线路面");
                        newMap1.put("sheetname", "分部-路面");
                    }
                    newMap1.put("yxps", map.get("密度规定值"));

                    newMap1.put("fbgc", "路面面层");
                    newMap1.put("检测点数", map.get("检测点数"));
                    newMap1.put("合格点数", map.get("合格点数"));
                    newMap1.put("合格率", map.get("合格率"));
                    resultysd.add(newMap1);
                }
                resultlist.addAll(resultysd);
            }
            else if (value.contains("34桥面平整度3米直尺法-")){
                List<Map<String, Object>> list = jjgFbgcQlgcQmpzdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> pzdlist = new ArrayList<>();
                for (Map<String, Object> map : list) {
                    Map map5 = new HashMap();
                    map5.put("ccname","平整度");
                    map5.put("ccname1","沥青路面");
                    map5.put("yxps",map.get("yxpc"));
                    map5.put("filename","详见《桥面系平整度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                    map5.put("sheetname", "分部-"+map.get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", map.get("合格率"));
                    pzdlist.add(map5);
                }
                resultlist.addAll(pzdlist);

            }else if (value.contains("35桥面横坡-")){
                List<Map<String, Object>> list = jjgFbgcQlgcQmhpService.lookJdb(commonInfoVo,value);
                if(!list.isEmpty()){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","横坡");
                    map5.put("ccname1","砼路面");
                    map5.put("yxps",list.get(0).get("yxpc"));
                    map5.put("filename","详见《桥面系横坡质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }


            }/*else if (value.contains("37桥面构造深度-")){
                List<Map<String, Object>> list =  jjgFbgcQlgcZdhgzsdService.lookJdb(commonInfoVo,value);
                if (list!=null){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","构造深度");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《桥面系构造深度质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }

            }else if (value.contains("36桥面摩擦系数-")){
                List<Map<String, Object>> list =  jjgFbgcQlgcZdhmcxsService.lookJdb(commonInfoVo,value);
                if (list!=null){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","摩擦系数");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《桥面系摩擦系数质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }

            }else if (value.contains("33桥面平整度-")){
                List<Map<String, Object>> list = jjgFbgcQlgcZdhpzdService.lookJdb(commonInfoVo,value);
                if (list != null && list.size()>0){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","平整度");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《桥面平整度质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }

            } */else if (value.contains("37构造深度手工铺沙法-")){
                List<Map<String, Object>> list = jjgFbgcQlgcQmgzsdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> qgzsdlist = new ArrayList<>();
                for (Map<String, Object> map : list) {
                    Map map5 = new HashMap();
                    map5.put("ccname","抗滑");
                    map5.put("ccname1","砼路面");
                    map5.put("ccname2","构造深度");
                    map5.put("yxps",map.get("yxpc"));
                    map5.put("filename","详见《桥面系构造深度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                    map5.put("sheetname", "分部-"+map.get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", map.get("合格率"));
                    qgzsdlist.add(map5);
                }
                resultlist.addAll(qgzsdlist);

            }else if (value.contains("39隧道衬砌厚度-")){
                List<Map<String, Object>> list = jjgFbgcSdgcCqhdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> cqhdlist = new ArrayList<>();
                //查询设计厚度
                String s = StringUtils.substringBetween(value, "-", ".");
                List<Map<String, Object>> sjhd = jjgFbgcSdgcCqhdService.selectsjhd(commonInfoVo,s);
                for (Map<String, Object> map : sjhd) {
                    for (Map<String, Object> stringObjectMap : list) {
                        //if (map.get("sdmc").equals(stringObjectMap.get("检测项目"))){
                            Map map5 = new HashMap();
                            map5.put("ccname","衬砌厚度");
                            map5.put("ccname1","衬砌厚度（mm）");
                            map5.put("ccname2","衬砌厚度（mm）");
                            map5.put("yxps",map.get("sjhd"));
                            map5.put("filename","详见《隧道衬砌厚度质量鉴定表》检测"+stringObjectMap.get("检测总点数")+"点,合格"+stringObjectMap.get("合格点数")+"点");
                            map5.put("sheetname", "分部-"+stringObjectMap.get("检测项目"));
                            map5.put("fbgc", "衬砌");
                            map5.put("合格率", stringObjectMap.get("合格率"));
                            cqhdlist.add(map5);
                        //}
                    }
                }
                resultlist.addAll(cqhdlist);
            }else if (value.contains("43隧道沥青路面压实度-")){
                List<Map<String, Object>> list = jjgFbgcSdgcSdlqlmysdService.lookJdbjgbgzbg(commonInfoVo,value);
                List<Map<String, Object>> sdysdlist = new ArrayList<>();
                double sdhgds = 0;
                double sdjcds = 0;
                double lmhgds = 0;
                double lmjcds = 0;
                boolean a = false;
                boolean b = false;
                String jcxm1 = "";
                String jcxm2 = "";
                String sdgdz = "";
                String lmgdz = "";
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        sdhgds += Double.valueOf(map.get("合格点数").toString());
                        sdjcds += Double.valueOf(map.get("检测点数").toString());
                        sdgdz = map.get("规定值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;
                    }else {
                        //沥青路面
                        lmhgds += Double.valueOf(map.get("合格点数").toString());
                        lmjcds += Double.valueOf(map.get("检测点数").toString());
                        lmgdz = map.get("规定值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b = true;
                    }
                }
                if (a){
                    double gdz1 = Double.valueOf(sdgdz)+1;
                    String gdz2= String.valueOf(gdz1);
                    Map map = new HashMap();
                    map.put("ccname","沥青路面压实度(%)");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",gdz2);
                    map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+decf.format(sdjcds)+"点,合格"+decf.format(sdhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (sdjcds != 0) ? df.format(sdhgds/sdjcds*100) : "0");
                    sdysdlist.add(map);
                }
                if (b){
                    double gdz3 = Double.valueOf(lmgdz)+1;
                    String gdz4 = String.valueOf(gdz3);
                    Map map = new HashMap();
                    map.put("ccname","沥青路面压实度(%)");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",gdz4);
                    map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+decf.format(lmjcds)+"点,合格"+decf.format(lmhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (lmjcds != 0) ? df.format(lmhgds/lmjcds*100) : "0");
                    sdysdlist.add(map);
                }
                resultlist.addAll(sdysdlist);

            }else if (value.contains("46隧道沥青路面渗水系数-")){
                List<Map<String, Object>> list = jjgFbgcSdgcLmssxsService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> ssxslist = new ArrayList<>();
                String jcxm = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm2="";
                boolean b= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("规定值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("规定值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","沥青路面渗水系数");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",sgdz);
                    map.put("filename","详见《路面渗水系数质量鉴定表》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (szds != 0) ? df.format(shgds/szds*100) : "0");
                    ssxslist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","沥青路面渗水系数");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",mgdz);
                    map.put("filename","详见《路面渗水系数质量鉴定表》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mzds != 0) ? df.format(mhgds/mzds*100) : "0");
                    ssxslist.add(map);
                }
                resultlist.addAll(ssxslist);

            }else if (value.contains("47隧道混凝土路面强度-")){
                List<Map<String, Object>> list = jjgFbgcSdgcHntlmqdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                Map map = new HashMap();
                map.put("ccname","混凝土路面强度");
                map.put("ccname1","路面面层");
                map.put("ccname2","");
                map.put("yxps",list.get(0).get("sjqd"));
                map.put("filename","详见《混凝土路面弯拉强度鉴定表》检测"+list.get(0).get("检测点数")+"点,合格"+list.get(0).get("合格点数")+"点");
                map.put("sheetname", "分部-"+s);
                map.put("fbgc", "路面面层");
                map.put("合格率", list.get(0).get("合格率"));
                relist.add(map);
                resultlist.addAll(relist);

            }else if (value.contains("48隧道混凝土路面相邻板高差-")){
                List<Map<String, Object>> list = jjgFbgcSdgcTlmxlbgcService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                Map map = new HashMap();
                map.put("ccname","混凝土路面相邻板高差");
                map.put("ccname1","路面面层");
                map.put("ccname2","");
                map.put("yxps",list.get(0).get("规定值"));
                map.put("filename","详见《混凝土路面相邻板高差质量鉴定表》检测"+list.get(0).get("检测点数")+"点,合格"+list.get(0).get("合格点数")+"点");
                map.put("sheetname", "分部-"+s);
                map.put("fbgc", "路面面层");
                map.put("合格率", list.get(0).get("合格率"));
                relist.add(map);
                resultlist.addAll(relist);

            }else if (value.contains("51构造深度手工铺沙法-")){
                List<Map<String, Object>> list = jjgFbgcSdgcLmgzsdsgpsfService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                Map map = new HashMap();
                map.put("ccname","构造深度手工铺沙法");
                map.put("ccname1","路面面层");
                map.put("ccname2","");
                map.put("yxps",list.get(0).get("设计值"));
                map.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+list.get(0).get("检测点数")+"点,合格"+list.get(0).get("合格点数")+"点");
                map.put("sheetname", "分部-"+s);
                map.put("fbgc", "路面面层");
                map.put("合格率", list.get(0).get("合格率"));
                relist.add(map);
                resultlist.addAll(relist);

            }else if (value.contains("54隧道混凝土路面厚度-钻芯法-")){
                List<Map<String, Object>> list = jjgFbgcSdgcSdhntlmhdzxfService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double qzds = 0;
                double qhgds = 0;
                String qgdz="";
                String jcxm2="";
                boolean b= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm3="";
                boolean c= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("设计值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else if (map.get("路面类型").toString().contains("桥面")){
                        qzds += Double.valueOf(map.get("检测点数").toString());
                        qhgds += Double.valueOf(map.get("合格点数").toString());
                        qgdz = map.get("设计值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    } else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("设计值").toString();
                        jcxm3 = map.get("检测项目").toString();
                        c= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",sgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率",(szds != 0) ? df.format(shgds/szds*100) : "0");
                    relist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","桥面系");
                    map.put("ccname2","");
                    map.put("yxps",qgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                    relist.add(map);
                }
                if (c){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",mgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm3);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mhgds != 0) ? df.format(mhgds/mzds*100) : "0");
                    relist.add(map);
                }
                resultlist.addAll(relist);

            }else if (value.contains("55隧道横坡-")){
                List<Map<String, Object>> list = jjgFbgcSdgcSdhpService.lookjg(commonInfoVo,value);
                String ccname3 ="";
                if (list.get(0).get("路面类型").toString().contains("沥青")){
                    ccname3 = "沥青路面";
                }else if(list.get(0).get("路面类型").toString().contains("混凝土")) {
                    ccname3 = "混凝土路面";
                }

                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double qzds = 0;
                double qhgds = 0;
                String qgdz="";
                String jcxm2="";
                boolean b= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm3="";
                boolean c= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("设计值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else if (map.get("路面类型").toString().contains("桥面")){
                        qzds += Double.valueOf(map.get("检测点数").toString());
                        qhgds += Double.valueOf(map.get("合格点数").toString());
                        qgdz = map.get("设计值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    } else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("设计值").toString();
                        jcxm3 = map.get("检测项目").toString();
                        c= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","路面横坡");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2",ccname3);
                    map.put("yxps",sgdz);
                    map.put("filename","详见《"+ccname3+"横坡质量鉴定表》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率",(szds != 0) ? df.format(shgds/szds*100) : "0");
                    relist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","路面横坡");
                    map.put("ccname1","桥面系");
                    map.put("ccname2",ccname3);
                    map.put("yxps",qgdz);
                    map.put("filename","详见《"+ccname3+"横坡质量鉴定表》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                    relist.add(map);
                }
                if (c){
                    Map map = new HashMap();
                    map.put("ccname","路面横坡");
                    map.put("ccname1","路面面层");
                    map.put("ccname2",ccname3);
                    map.put("yxps",mgdz);
                    map.put("filename","详见《"+ccname3+"横坡质量鉴定表》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm3);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mhgds != 0) ? df.format(mhgds/mzds*100) : "0");
                    relist.add(map);
                }
                resultlist.addAll(relist);


            }else if (value.contains("53隧道沥青路面厚度-钻芯法-")){
                List<Map<String, Object>> list = jjgFbgcSdgcGssdlqlmhdzxfService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double qzds = 0;
                double qhgds = 0;
                String qgdz="";
                String jcxm2="";
                boolean b= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm3="";
                boolean c= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("设计值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else if (map.get("路面类型").toString().contains("桥面")){
                        qzds += Double.valueOf(map.get("检测点数").toString());
                        qhgds += Double.valueOf(map.get("合格点数").toString());
                        qgdz = map.get("设计值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    } else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("设计值").toString();
                        jcxm3 = map.get("检测项目").toString();
                        c= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",sgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率",(szds != 0) ? df.format(shgds/szds*100) : "0");
                    relist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","桥面系");
                    map.put("ccname2","");
                    map.put("yxps",qgdz);
                    map.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                    relist.add(map);
                }
                if (c){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",mgdz);
                    map.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm3);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mhgds != 0) ? df.format(mhgds/mzds*100) : "0");
                    relist.add(map);
                }
                resultlist.addAll(relist);

            }else if (value.contains("42隧道断面测量坐标表-")){
                List<Map<String, Object>> list = jjgFbgcSdgcJkService.lookJdbjg(commonInfoVo);
                if (list!=null && list.size()>0){

                    Map<String, List<Map<String, Object>>> groupByData = list.stream()
                            .collect(Collectors.groupingBy(m -> m.get("sdname").toString()));
                    groupByData.forEach((group, grouphtdData) -> {
                        double zds = 0,hgds = 0;
                        for (Map<String, Object> grouphtdDatum : grouphtdData) {
                            zds += Double.valueOf(grouphtdDatum.get("总点数").toString());
                            hgds += Double.valueOf(grouphtdDatum.get("合格点数").toString());
                        }
                        Map maps = new HashMap();
                        maps.put("ccname","净空");
                        maps.put("ccname1","隧道路面");
                        maps.put("ccname2","");
                        maps.put("yxps","偏差值≥0");
                        maps.put("filename","详见《隧道断面测量坐标表》检测"+decf.format(zds)+"个总断面,合格"+decf.format(hgds)+"个断面");
                        maps.put("sheetname", "分部-"+group);
                        maps.put("fbgc", "总体");
                        maps.put("合格率",(zds != 0) ? df.format(hgds/zds*100) : "0");
                        resultlist.add(maps);
                    });


                   /* for (Map<String, Object> map : list) {
                        double zds = Double.valueOf(map.get("总点数").toString());
                        double hgds = Double.valueOf(map.get("合格点数").toString());
                        String s = map.get("sdname").toString();

                        Map maps = new HashMap();
                        maps.put("ccname","净空");
                        maps.put("ccname1","隧道路面");
                        maps.put("ccname2","");
                        maps.put("yxps","偏差值≥0");
                        maps.put("filename","详见《隧道断面测量坐标表》检测"+decf.format(zds)+"个总断面,合格"+decf.format(hgds)+"个断面");
                        maps.put("sheetname", "分部-"+s);
                        maps.put("fbgc", "总体");
                        maps.put("合格率",(zds != 0) ? df.format(hgds/zds*100) : "0");
                        resultlist.add(maps);

                    }*/

                }

            }/*else if (value.contains("45隧道路面车辙-")){
                List<Map<String,Object>> list = jjgFbgcSdgcZdhczService.lookJdb(commonInfoVo,value);
                if (list!=null && list.size()>0){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","车辙");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《隧道路面车辙质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "路面面层");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }

            }else if (value.contains("51隧道路面构造深度-")){
                List<Map<String,Object>> list = jjgFbgcSdgcZdhgzsdService.lookJdb(commonInfoVo,value);
                if (list!=null && list.size()>0){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","构造深度");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《隧道构造深度质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "路面面层");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }

            }else if (value.contains("52隧道雷达厚度-")){
                List<Map<String,Object>> list = jjgFbgcSdgcZdhldhdService.lookJdb(commonInfoVo,value);
                if (list!=null && list.size()>0){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","雷达厚度");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《隧道雷达厚度质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "路面面层");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }


            }else if (value.contains("50隧道路面摩擦系数-")){
                List<Map<String,Object>> list = jjgFbgcSdgcZdhmcxsService.lookJdb(commonInfoVo,value);
                if (list!=null && list.size()>0){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","摩擦系数");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《隧道摩擦系数质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "路面面层");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }

            }else if (value.contains("49隧道路面平整度-")){
                List<Map<String,Object>> list = jjgFbgcSdgcZdhpzdService.lookJdb(commonInfoVo,value);
                if (list!=null && list.size()>0){
                    double zds = 0;
                    double hgds = 0;
                    for (Map<String, Object> map : list) {
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                    List<Map<String, Object>> qhplist = new ArrayList<>();
                    Map map5 = new HashMap();
                    map5.put("ccname","平整度");
                    map5.put("ccname1","");
                    map5.put("yxps",list.get(0).get("设计值"));
                    map5.put("filename","详见《隧道路面平整度质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                    map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                    map5.put("fbgc", "路面面层");
                    map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                    qhplist.add(map5);
                    resultlist.addAll(qhplist);
                }
            }*/

        }

        for (Map<String, Object> map : resultlist) {
            if (map.containsKey("sheetname")){
                String s = map.get("sheetname").toString();
                if (s.contains("钻芯法-")){
                    String substring = s.substring(0, 3);
                    map.put("sheetname",substring);
                }
            }
        }
        Map<String, List<Map<String, Object>>> groupedData = resultlist.stream()
                .filter(map -> map.get("sheetname") != null) // 添加非空判断
                .collect(Collectors.groupingBy(map -> (String) map.get("sheetname")));
        /*Map<String, List<Map<String, Object>>> groupedData = resultlist.stream()
                .filter(map -> map.get("sheetname") != null)
                .sorted(Comparator.comparing(map -> (String) ((Map<String, Object>)map).get("sheetname"))
                        .thenComparing(map -> (String) ((Map<String, Object>)map).get("ccname")))
                .collect(Collectors.groupingBy(map -> (String) ((Map<String, Object>)map).get("sheetname")));*/

        //按分组后的数据，每个分组写一个工作簿
        DBwriteToExcel(groupedData,proname,htd);

        // 对路面面层的数据进行抽取出来
    }

    // 把括号内地内容匹配出来，并且可以兼容英文和中文括号
    private String getPlace(String s){
        if(s.contains("(")) { // 英文括号
            return s.replaceAll(".*?\\((.*?)\\).*?", "$1");
        }else { // 中文括号
            return s.replaceAll(".*?\\（(.*?)\\）.*?", "$1");
        }
    }

    @Override
    public void generateJSZLPdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String path = filespath+ File.separator+proname+File.separator;
        List<String> filteredFiles = filterFiles(path);

        List<Map<String,Object>> mapList = new ArrayList<>();

        for (String filteredFile : filteredFiles) {
            String pdnpath = path+filteredFile+ File.separator;
            XSSFWorkbook wb = null;
            File f = new File(pdnpath+"00评定表.xlsx");
            if (!f.exists()){
                continue;
            }else {
                FileInputStream in = new FileInputStream(f);
                wb = new XSSFWorkbook(in);
                XSSFSheet sheet = wb.getSheet("合同段");

                for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) {
                    Map map = new HashMap();
                    XSSFRow row = sheet.getRow(i);

                    if (row.getCell(0).getStringCellValue().equals("合同段质量等级")){
                        String djValue = row.getCell(1).getStringCellValue();
                        map.put("htdValue",sheet.getRow(1).getCell(1).getStringCellValue());
                        map.put("djValue",djValue);
                        mapList.add(map);
                    }

                }
            }
        }

        DBwriteJSZLToExcel(mapList,commonInfoVo);




    }



    /**
     *
     * @param mapList
     * @param commonInfoVo
     * @throws IOException
     */
    private void DBwriteJSZLToExcel(List<Map<String, Object>> mapList, CommonInfoVo commonInfoVo) throws IOException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String proname = commonInfoVo.getProname();
        XSSFWorkbook wb = null;
        File f = new File(filepath + File.separator + proname + File.separator + "建设项目质量评定表.xlsx");
        File fdir = new File(filepath + File.separator + proname);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        try {
            /*File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "建设项目质量评定表.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));*/
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/建设项目质量评定表.xlsx");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            XSSFSheet sheet = wb.getSheet("建设项目");
            createdJSXMTable(wb,getJSXMNum(mapList.size()));

            sheet.getRow(1).getCell(2).setCellValue(proname);
            sheet.getRow(1).getCell(6).setCellValue(proname);
            sheet.getRow(2).getCell(2).setCellValue(getAllzh(proname));
            sheet.getRow(2).getCell(6).setCellValue(simpleDateFormat.format(new Date()));
            int index = 0;
            for (Map<String, Object> map : mapList) {
                sheet.getRow(index + 4).getCell(0).setCellValue(map.get("htdValue").toString());
                sheet.getRow(index + 4).getCell(3).setCellValue(map.get("djValue").toString());
                index++;
            }

            calculateJSZLSheet(sheet);
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }

            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();
        }catch (Exception e) {
            if(f.exists()){
                f.delete();
            }
            throw new JjgysException(20001, "生成建设项目质量评定表错误，请检查数据的正确性");
        }



    }

    /**
     *
     * @param sheet
     */
    private void calculateJSZLSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("建设项目质量等级".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-24);
                rowend = sheet.getRow(i-1);
                /*row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference() +":"+rowend.getCell(2).getReference()
                        +",\"<>合格\")=0,\"合格\", \"不合格\")");*///=IF(COUNTIF(C64:C81,"<>合格")=0, "合格", "不合格")
                row.getCell(3).setCellFormula("IF(COUNTIF("+rowstart.getCell(3).getReference()
                        +":"+rowend.getCell(3).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(2).getReference()
                        +":"+rowend.getCell(3).getReference()+"),\"合格\", \"不合格\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")

            }
        }


    }

    /**
     *
     * @param proname
     * @return
     */
    private String getAllzh(String proname) {
        String zh = jjgHtdService.getAllzh(proname);
        return zh;
    }

    /**
     *
     * @param wb
     * @param record
     */
    private void createdJSXMTable(XSSFWorkbook wb, int record) {
        //复制表格的
        for(int i = 1; i < record; i++){
            RowCopy.copyRows(wb, "建设项目", "建设项目", 4, 29, (i - 1) * 26 + 30);
        }
        XSSFSheet sheet = wb.getSheet("建设项目");

        //取消后两行的合并单元格 然后复制source中的内容
        int lastRowNum = sheet.getLastRowNum();
        int numMergedRegions = sheet.getNumMergedRegions();

        for (int i = numMergedRegions - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();

            if (lastRow >= lastRowNum - 1 && lastRow <= lastRowNum) {
                sheet.removeMergedRegion(i);
            }
        }
        RowCopy.copyRows(wb, "source", "建设项目", 0, 1,(record) * 26+2);
        wb.setPrintArea(wb.getSheetIndex("建设项目"), 0, 7, 0, (record) * 26+3);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getJSXMNum(int size) {

        return size%26 ==0 ? size/26 : size/26+1;
    }

    /**
     *
     * @param groupedData
     */
    private void DBwriteToExcel(Map<String, List<Map<String, Object>>> groupedData,String proname,String htd) throws IOException {
        XSSFWorkbook wb = null;
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        //try {
           /* File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "合同段评定表.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));*/
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/合同段评定表.xlsx");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            // 对key进行排序
            List<String> sortedKeys = groupedData.keySet().stream()
                    .sorted(new Comparator<String>() {
                        @Override
                        public int compare(String o1, String o2) {
                            // 判断包含关键字的顺序
                            List<String> keywords = Arrays.asList("分部-路面", "分部-交安","分部-路基", "桥", "隧道");
                            for (String keyword : keywords) {
                                if (o1.contains(keyword) && !o2.contains(keyword)) {
                                    return -1; // o1包含关键字，o2不包含，则o1排在前面
                                } else if (!o1.contains(keyword) && o2.contains(keyword)) {
                                    return 1; // o1不包含关键字，o2包含，则o2排在前面
                                }
                            }
                            // 如果没有关键字，按字典序排列
                            return o1.compareTo(o2);
                        }
                    })
                    .collect(Collectors.toList());

            // 按关键字排序后的数据
            Map<String, List<Map<String, Object>>> sortedData = new LinkedHashMap<>();

            // 按排序后的key遍历原始数据
            for (String key : sortedKeys) {
                List<Map<String, Object>> dataList = groupedData.get(key);
                sortedData.put(key, dataList);
            }
            for (String key : sortedData.keySet()) {
                List<Map<String, Object>> list = groupedData.get(key);

                if (key.equals("分部-路基")){
                    writeLJData(wb,list,key,proname,htd);
                }else if (key.equals("分部-路面")){
                    writeLJData(wb,list,key,proname,htd);
                }else if (key.equals("分部-交安")){
                    writeLJData(wb,list,key,proname,htd);
                }else {
                    //桥梁和隧道
                    writeLJData(wb,list,key,proname,htd);
                }
            }
            //单位工程
            List<Map<String, Object>> dwgclist = new ArrayList<>();
            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.startsWith("分部-")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
            writedwgcData(wb,dwgclist);

            //合同段
            List<Map<String, Object>> list = processhtdSheet(wb);
            writedhtdData(wb,list);

            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();

        /*}catch (Exception e) {
            if(f.exists()){
                f.delete();
            }
            throw new JjgysException(20001, "生成评定表错误，请检查数据的正确性");
        }*/


    }


    /**
     *
     * @param wb
     * @param list
     */
    private void writedhtdData(XSSFWorkbook wb, List<Map<String, Object>> list) {
        XSSFSheet sheet = wb.getSheet("合同段");
        createdHtdTable(wb,getNum(list.size()));
        int index = 0;
        int tableNum = 0;
        fillTitleHtdData(sheet, tableNum, list.get(0));
        for (Map<String, Object> datum : list) {
            if (index > 20){
                tableNum++;
                index = 0;
            }
            fillCommonHtdData(sheet,tableNum,index, datum);
            index++;
        }
        calculateHtdSheet(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }
    }

    /**
     *
     * @param sheet
     */
    private void calculateHtdSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合同段质量等级".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-21);
                rowend = sheet.getRow(i-1);
                /*row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference() +":"+rowend.getCell(2).getReference()
                        +",\"<>合格\")=0,\"合格\", \"不合格\")");*///=IF(COUNTIF(C64:C81,"<>合格")=0, "合格", "不合格")
                row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(4).getReference()
                        +":"+rowend.getCell(4).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(4).getReference()
                        +":"+rowend.getCell(4).getReference()+"),\"合格\", \"不合格\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")

            }
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param datum
     */
    private void fillTitleHtdData(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 29 +1).getCell(1).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 29 + 1).getCell(5).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(1).setCellValue(datum.get("sgdw").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(5).setCellValue(datum.get("jldw").toString());

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonHtdData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(index + 5).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(index + 5).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(index + 5).getCell(4).setCellValue(datum.get("zldj").toString());

    }

    /**
     *
     * @param wb
     * @param gettableNum
     */
    private void createdHtdTable(XSSFWorkbook wb, int gettableNum) {
        int record = 0;
        record = gettableNum;
        for(int i = 1; i < record; i++){
            if(i < record-1){
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 28, (i - 1) * 24 + 29);
            }
            else{
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 26, (i - 1) * 24 + 29);
            }
        }
        XSSFSheet sheet = wb.getSheet("合同段");
        int lastRowNum = sheet.getLastRowNum();
        int numMergedRegions = sheet.getNumMergedRegions();

        for (int i = numMergedRegions - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();

            if (lastRow >= lastRowNum - 2 && lastRow <= lastRowNum) {
                sheet.removeMergedRegion(i);
            }
        }
        RowCopy.copyRows(wb, "source", "合同段", 0, 2,(record) * 24+2);
        wb.setPrintArea(wb.getSheetIndex("合同段"), 0, 7, 0, (record) * 24+4);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getNum(int size) {
        return size%24 ==0 ? size/24 : size/24+1;
    }

    /**
     *
     * @param wb
     * @return
     */
    private List<Map<String,Object>> processhtdSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("单位工程");
        List<Map<String,Object>> list = new ArrayList<>();
        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("单位工程名称：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("dwgc",row.getCell(1).getStringCellValue());
                map.put("jsxm",row.getCell(3).getStringCellValue());
                map.put("htd",sheet.getRow(i+6).getCell(0).getStringCellValue());
                map.put("sgdw",sheet.getRow(i+2).getCell(1).getStringCellValue());
                map.put("jldw",sheet.getRow(i+2).getCell(3).getStringCellValue());
                map.put("zldj",sheet.getRow(i+24).getCell(1).getStringCellValue());
                list.add(map);
            }
        }
        System.out.println(list);
        return list;

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void writedwgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        if (data.size()>0&&data!=null){
            XSSFSheet sheet = wb.getSheet("单位工程");
            createdwgcTable(wb,getdwgcNum(data));

            int index = 0;
            int tableNum = 0;
            String fbgc = data.get(0).get("dwgc").toString();
            for (Map<String, Object> datum : data) {
                if (fbgc.equals(datum.get("dwgc"))){
                    fillTitleDwgcData(sheet,tableNum,datum);

                    fillCommonDwgcData(sheet,tableNum,index,datum);
                    index++;
                }else {
                    fbgc = datum.get("dwgc").toString();
                    tableNum ++;
                    index = 0;
                    fillTitleDwgcData(sheet,tableNum,datum);


                    fillCommonDwgcData(sheet,tableNum,index,datum);
                    index += 1;
                }

            }


            calculateDwgcSheet(sheet);
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }
        }


    }

    /**
     * 合并单元格
     * @param rowAndcol
     * @return
     */
    private List<Map<String, Object>> mergeCells(List<Map<String, Object>> rowAndcol) {
        List<Map<String, Object>> result = new ArrayList<>();
        int currentEndRow = -1;
        int currentStartRow = -1;
        int currentStartCol = -1;
        int currentEndCol = -1;
        String currentName = null;
        int currentTableNum = -1;
        for (Map<String, Object> row : rowAndcol) {
            int tableNum = (int) row.get("tableNum");
            int startRow = (int) row.get("startRow");
            int endRow = (int) row.get("endRow");
            int startCol = (int) row.get("startCol");
            int endCol = (int) row.get("endCol");
            String name = (String) row.get("name");
            if (currentName == null || !currentName.equals(name) || currentStartCol != startCol || currentEndCol != endCol || currentTableNum != tableNum) {
                if (currentStartRow != -1) {
                    for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                        Map<String, Object> newRow = new HashMap<>();
                        newRow.put("name", currentName);
                        newRow.put("startRow", currentStartRow);
                        newRow.put("endRow", currentEndRow);
                        newRow.put("startCol", currentStartCol);
                        newRow.put("endCol", currentEndCol);
                        newRow.put("tableNum", currentTableNum);
                        newRow.putAll(result.get(i));
                        result.set(i, newRow);
                    }
                }
                currentName = name;
                currentStartCol = startCol;
                currentEndCol = endCol;
                currentTableNum = tableNum;
                currentStartRow = startRow;
                currentEndRow = endRow;
                result.add(row);
            } else {
                Map<String, Object> lastRow = result.get(result.size() - 1);
                lastRow.put("endRow", endRow);
                currentEndRow = endRow;
            }
        }
        if (currentStartRow != -1) {
            for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                Map<String, Object> newRow = new HashMap<>();
                newRow.put("name", currentName);
                newRow.put("startRow", currentStartRow);
                newRow.put("endRow", currentEndRow);
                newRow.put("startCol", currentStartCol);
                newRow.put("endCol", currentEndCol);
                newRow.put("tableNum", currentTableNum);
                newRow.putAll(result.get(i));
                result.set(i, newRow);
            }
        }
        return result;
    }

    /**
     *
     * @param sheet
     */
    private void calculateDwgcSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("单位工程质量等级".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-18);
                rowend = sheet.getRow(i-1);
                //=IF(COUNTIF(C36:C53,"合格")>0,"合格","不合格")
                /*row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference() +":"+rowend.getCell(2).getReference()
                        +",\"<>合格\")=0,\"合格\", \"不合格\")");*///=IF(COUNTIF(C64:C81,"<>合格")=0, "合格", "不合格")
                row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference()
                                +":"+rowend.getCell(2).getReference()+",\"不合格\")>0,\"不合格\",\"合格\")");
//                row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference()
//                        +":"+rowend.getCell(2).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(2).getReference()
//                        +":"+rowend.getCell(2).getReference()+"),\"合格\", \"不合格\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")

            }
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param datum
     */
    private void fillTitleDwgcData(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 +1).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(tableNum * 28 +1).getCell(3).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 28 +2).getCell(1).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 28 +2).getCell(3).setCellValue(datum.get("gcbw").toString());
        sheet.getRow(tableNum * 28 +3).getCell(1).setCellValue(datum.get("sgbw").toString());
        sheet.getRow(tableNum * 28 +3).getCell(3).setCellValue(datum.get("jldw").toString());


    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonDwgcData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 + index + 7).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(1).setCellValue(datum.get("fbgc").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(2).setCellValue(datum.get("sfhg").toString());

    }

    /**
     * 单位工程
     * @param wb
     * @param gettableNum
     */
    private void createdwgcTable(XSSFWorkbook wb, int gettableNum) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, "单位工程", "单位工程", 0, 27, i*28);
        }
        if(gettableNum > 1){
            wb.setPrintArea(wb.getSheetIndex("单位工程"), 0, 5, 0, gettableNum*28-1);
        }

    }

    /**
     *
     * @param data
     * @return
     */
    private int getdwgcNum(List<Map<String, Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("dwgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%18==0){
                num += value/18;
            }else {
                num += value/18+1;
            }
        }
        return num;

    }

    /**
     *
     * @param sheet
     * @return
     */
    private List<Map<String,Object>> processSheet(Sheet sheet) {
        List<Map<String,Object>> list = new ArrayList<>();

        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("合同段：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("htd",row.getCell(2).getStringCellValue());
                map.put("fbgc",row.getCell(8).getStringCellValue());
                map.put("jsxm",row.getCell(15).getStringCellValue());

                map.put("gcbw",sheet.getRow(i+1).getCell(2).getStringCellValue());
                map.put("sgbw",sheet.getRow(i+1).getCell(8).getStringCellValue());
                map.put("jldw",sheet.getRow(i+1).getCell(15).getStringCellValue());
                if (sheet.getRow(i+20).getCell(16).getStringCellValue().equals("√")){
                    map.put("sfhg","合格");
                }else {
                    map.put("sfhg","不合格");
                }

                map.put("dwgc",sheet.getSheetName().substring(sheet.getSheetName().indexOf("-")+1));
                list.add(map);
            }
        }
        return list;
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void writeLJData(XSSFWorkbook wb, List<Map<String, Object>> data,String sheetname,String proname,String htd) {
        Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> map1, Map<String, Object> map2) {
                String fbgc1 = (String) map1.get("fbgc");
                String fbgc2 = (String) map2.get("fbgc");
                int result = fbgc1.compareTo(fbgc2);

                // 如果 fbgc 相同，则比较 ccname
                if (result == 0) {
                    String ccname1 = (String) map1.get("ccname"); // 假设ccname是另一个键
                    String ccname2 = (String) map2.get("ccname");
                    result = ccname1.compareTo(ccname2);
                }

                return result;
            }
        });
        copySheet(wb,sheetname);
        XSSFPrintSetup ps = wb.getSheet(sheetname).getPrintSetup();
        ps.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)

        JjgHtd htdlist = jjgHtdService.selectInfo(proname,htd);
        XSSFSheet sheet = wb.getSheet(sheetname);
        createTable(wb,gettableNum(data),sheetname);

        int index = 0;
        int tableNum = 0;
        int startRow = -1, endRow = -1, startCol = -1, endCol = -1, startCols = -1, endCols = -1, startColhgl = -1, endColhgl = -1, startColzl = -1, endColzl = -1;
        List<Map<String, Object>> rowAndcol = new ArrayList<>();
        List<Map<String, Object>> rowAndcol1 = new ArrayList<>();
        List<Map<String, Object>> rowAndcolhgl = new ArrayList<>();
        List<Map<String, Object>> rowAndcolzl = new ArrayList<>();
        String ccname = data.get(0).get("ccname").toString();
        String fbgc = data.get(0).get("fbgc").toString();
        String filename = data.get(0).get("filename").toString();
        for (Map<String, Object> datum : data) {
            if (index % 15 == 0 && index!=0){
                tableNum++;
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                index = 0;
            }
            if (fbgc.equals(datum.get("fbgc"))){
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                if (ccname.equals(datum.get("ccname").toString())){

                    startRow = tableNum * 22 + 6 + index % 16 ;
                    endRow = tableNum * 22 + 6 + index % 16 ;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",ccname);
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 17;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 20;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);


                }else {
                    ccname = datum.get("ccname").toString();
                    startRow = tableNum * 22 + 6 + index % 16 ;
                    endRow = tableNum * 22 + 6 + index % 16 ;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",ccname);
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 17;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 20;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);
                }
                fillCommonData(sheet,tableNum,index,datum);
                index++;
            }else {
                fbgc = datum.get("fbgc").toString();
                tableNum ++;
                index = 0;
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                ccname = datum.get("ccname").toString();
                if (ccname.equals(datum.get("ccname").toString())) {
                    startRow = tableNum * 22 + 6 + index % 16;
                    endRow = tableNum * 22 + 6 + index % 16;

                    startCol = 2;
                    endCol = 5;

                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow", startRow);
                    map.put("endRow", endRow);
                    map.put("startCol", startCol);
                    map.put("endCol", endCol);
                    map.put("name", ccname);
                    map.put("tableNum", tableNum);
                    rowAndcol.add(map);

                    startCols = 7;
                    endCols = 16;
                    Map<String, Object> map1 = new HashMap<>();
                    map1.put("startRow",startRow);
                    map1.put("endRow",endRow);
                    map1.put("startCol",startCols);
                    map1.put("endCol",endCols);
                    map1.put("name",ccname);
                    map1.put("tableNum",tableNum);
                    rowAndcol1.add(map1);

                    startColhgl = 17;
                    endColhgl = 18;
                    Map<String, Object> map2 = new HashMap<>();
                    map2.put("startRow",startRow);
                    map2.put("endRow",endRow);
                    map2.put("startCol",startColhgl);
                    map2.put("endCol",endColhgl);
                    map2.put("name",ccname);
                    map2.put("tableNum",tableNum);
                    rowAndcolhgl.add(map2);

                    startColzl = 19;
                    endColzl = 20;
                    Map<String, Object> map3 = new HashMap<>();
                    map3.put("startRow",startRow);
                    map3.put("endRow",endRow);
                    map3.put("startCol",startColzl);
                    map3.put("endCol",endColzl);
                    map3.put("name",ccname);
                    map3.put("tableNum",tableNum);
                    rowAndcolzl.add(map3);




                }


                fillCommonData(sheet,tableNum,index,datum);
                index += 1;
            }
            ccname = datum.get("ccname").toString();


        }
        List<Map<String, Object>> maps = mergeCells(rowAndcol);
        List<Map<String, Object>> mapss = mergeCells(rowAndcol1);
        List<Map<String, Object>> maphgl = mergeCells(rowAndcolhgl);
        List<Map<String, Object>> mapzl = mergeCells(rowAndcolzl);
        for (Map<String, Object> map : maps) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        if ("分部-交安".equals(sheetname) || "分部-路基".equals(sheetname)){
            for (Map<String, Object> map : mapss) {
                int startRow1 = Integer.parseInt(map.get("startRow").toString());
                int endRow1 = Integer.parseInt(map.get("endRow").toString());
                int startCol1 = Integer.parseInt(map.get("startCol").toString());
                int endCol1 = Integer.parseInt(map.get("endCol").toString());
                CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
                // 检查是否存在重叠的合并区域
                List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                    CellRangeAddress mergedRegion = mergedRegions.get(i);
                    if (mergedRegion.intersects(newRegion)){
                        sheet.removeMergedRegion(i);
                    }
                }
                sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
            }
            for (Map<String, Object> map : maphgl) {
                int startRow1 = Integer.parseInt(map.get("startRow").toString());
                int endRow1 = Integer.parseInt(map.get("endRow").toString());
                int startCol1 = Integer.parseInt(map.get("startCol").toString());
                int endCol1 = Integer.parseInt(map.get("endCol").toString());
                CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
                // 检查是否存在重叠的合并区域
                List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                    CellRangeAddress mergedRegion = mergedRegions.get(i);
                    if (mergedRegion.intersects(newRegion)){
                        sheet.removeMergedRegion(i);
                    }
                }
                sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
            }
            for (Map<String, Object> map : mapzl) {
                int startRow1 = Integer.parseInt(map.get("startRow").toString());
                int endRow1 = Integer.parseInt(map.get("endRow").toString());
                int startCol1 = Integer.parseInt(map.get("startCol").toString());
                int endCol1 = Integer.parseInt(map.get("endCol").toString());
                CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
                // 检查是否存在重叠的合并区域
                List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
                for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                    CellRangeAddress mergedRegion = mergedRegions.get(i);
                    if (mergedRegion.intersects(newRegion)){
                        sheet.removeMergedRegion(i);
                    }
                }
                sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
            }
        }


        /*for (Map<String, Object> map : maps) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : mapss) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }
        for (Map<String, Object> map : maphgl) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }

        for (Map<String, Object> map : mapzl) {
            int startRow1 = Integer.parseInt(map.get("startRow").toString());
            int endRow1 = Integer.parseInt(map.get("endRow").toString());
            int startCol1 = Integer.parseInt(map.get("startCol").toString());
            int endCol1 = Integer.parseInt(map.get("endCol").toString());
            CellRangeAddress newRegion = new CellRangeAddress(startRow1, endRow1, startCol1, endCol1);
            // 检查是否存在重叠的合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = mergedRegions.size() - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                if (mergedRegion.intersects(newRegion)){
                    sheet.removeMergedRegion(i);
                }
            }
            sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
        }*/

        //数据写完当前工作簿的后，就要插入公式计算了
        calculateFbgcSheet(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }

    }

    /**
     * 计算分部工程的结果
     * @param sheet
     */
    private void calculateFbgcSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("实测项目是否全部合格".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-15);
                rowend = sheet.getRow(i-1);
                //=IF(COUNTIF(T29:T43,"不合格")>0,"","√")
                //row.getCell(8).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"√\", \"\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")
                row.getCell(8).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"\",\"√\")");
                //=IF(COUNTIF(T95:T109,"不合格")>0,"×","")
                row.getCell(10).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"×\",\"\")");
                //row.getCell(10).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"×\", \"\")");//=IF(COUNTIF(T7:T21, "不合格") = COUNTA(T7:T21), "", "×")
                row.getCell(16).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"\",\"√\")");
                row.getCell(19).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")>0,\"×\",\"\")");
                //row.getCell(16).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"√\", \"\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")
                //row.getCell(19).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"不合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"×\", \"\")");//=IF(COUNTIF(T7:T21, "不合格") = COUNTA(T7:T21), "", "×")
            }
        }

    }

    /**
     *
     * @param wb
     * @param sheetname
     */
    private void copySheet(XSSFWorkbook wb,String sheetname) {
        String sourceSheetName = "模板";
        String targetSheetName = sheetname;
        XSSFSheet sourceSheet = wb.getSheet(sourceSheetName);

        // 创建新的工作表，并将源工作表中的内容复制到新工作表
        XSSFSheet targetSheet = wb.createSheet(targetSheetName);
        copySheetInfo(wb,sourceSheet, targetSheet);


    }


    /**
     *
     * @param wb
     * @param sourceSheet
     * @param targetSheet
     */
    private static void copySheetInfo(XSSFWorkbook wb, Sheet sourceSheet, Sheet targetSheet) {
        int maxColumnNum = 0;
        for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row newRow = targetSheet.createRow(i);
            if (sourceRow != null) {
                copyRow(sourceRow, newRow, wb);
                if (sourceRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = sourceRow.getLastCellNum();
                }
            }
        }

        // 复制列宽
        for (int j = 0; j < maxColumnNum; j++) {
            targetSheet.setColumnWidth(j, sourceSheet.getColumnWidth(j));
        }

        // 复制合并单元格
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            targetSheet.addMergedRegion(mergedRegion);
        }

        // 复制打印区域
        PrintSetup sourcePrintSetup = sourceSheet.getPrintSetup();
        PrintSetup targetPrintSetup = targetSheet.getPrintSetup();
        targetPrintSetup.setLandscape(sourcePrintSetup.getLandscape());
        targetPrintSetup.setPaperSize(sourcePrintSetup.getPaperSize());
        targetPrintSetup.setFitWidth(sourcePrintSetup.getFitWidth());
        targetPrintSetup.setFitHeight(sourcePrintSetup.getFitHeight());
        targetPrintSetup.setScale(sourcePrintSetup.getScale());

        // 复制页眉
        Header sourceHeader = sourceSheet.getHeader();
        Header targetHeader = targetSheet.getHeader();
        targetHeader.setCenter(sourceHeader.getCenter());
        targetHeader.setLeft(sourceHeader.getLeft());
        targetHeader.setRight(sourceHeader.getRight());

        // 复制页脚
        Footer sourceFooter = sourceSheet.getFooter();
        Footer targetFooter = targetSheet.getFooter();
        targetFooter.setCenter(sourceFooter.getCenter());
        targetFooter.setLeft(sourceFooter.getLeft());
        targetFooter.setRight(sourceFooter.getRight());


    }

    /**
     *
     * @param sourceRow
     * @param newRow
     * @param wb
     */
    private static void copyRow(Row sourceRow, Row newRow, Workbook wb) {
        newRow.setHeight(sourceRow.getHeight());
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);
            if (sourceCell != null) {
                CellStyle sourceCellStyle = sourceCell.getCellStyle();
                CellStyle newCellStyle = wb.createCellStyle();
                newCellStyle.cloneStyleFrom(sourceCellStyle);
                newCell.setCellStyle(newCellStyle);

                CellType cellType = sourceCell.getCellTypeEnum();

                switch (cellType) {
                    case BOOLEAN:
                        newCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case STRING:
                        newCell.setCellValue(sourceCell.getRichStringCellValue());
                        break;
                    case NUMERIC:
                        newCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case FORMULA:
                        newCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    case BLANK:
                        // do nothing
                        break;
                    default:
                        // do nothing
                        break;
                }
            }
        }
    }


    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(tableNum * 22 + index + 6).getCell(1).setCellValue(1+index);
        sheet.getRow(tableNum * 22 + index + 6).getCell(2).setCellValue(datum.get("ccname").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(6).setCellValue(String.valueOf(datum.get("yxps")));
        sheet.getRow(tableNum * 22 + index + 6).getCell(7).setCellValue(datum.get("filename").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(17).setCellValue(datum.get("合格率").toString());
        if (datum.get("ccname").toString().contains("*")){
            Double value = Double.valueOf(datum.get("合格率").toString());
            if (value == 100){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }else if (datum.get("ccname").toString().contains("△")){
            Double value = Double.valueOf(datum.get("合格率").toString());
            if (value >= 95){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }else {
            Double value = Double.valueOf(datum.get("合格率").toString());
            if (value >= 80){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }

    }

    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param htdlist
     * @param fbgc
     */
    private void fillTitleData(XSSFSheet sheet, int tableNum, String proname, String htd,JjgHtd htdlist,String fbgc){
        sheet.getRow(tableNum * 22 +1).getCell(2).setCellValue(htd);
        sheet.getRow(tableNum * 22 +1).getCell(8).setCellValue(fbgc);
        sheet.getRow(tableNum * 22 +2).getCell(2).setCellValue(getgcbw(htdlist.getZhq(),htdlist.getZhz()));
        sheet.getRow(tableNum * 22 +2).getCell(8).setCellValue(htdlist.getSgdw());
        sheet.getRow(tableNum * 22 +2).getCell(15).setCellValue(htdlist.getJldw());
        sheet.getRow(tableNum * 22 +1).getCell(15).setCellValue(proname);
    }

    /**
     *分部
     * @param zhq
     * @param zhz
     * @return
     */
    private String getgcbw(String zhq, String zhz) {
        DecimalFormat decimalFormat = new DecimalFormat("#.###");
        double a = Double.valueOf(zhq);
        double b = Double.valueOf(zhz);

        int aa = (int) a / 1000;
        int bb = (int) b / 1000;
        double cc = a % 1000;
        double dd = b % 1000;
        String result = "K"+aa+"+"+decimalFormat.format(cc)+"--"+"K"+bb+"+"+decimalFormat.format(dd);

       return result;
    }

    /**
     * 分部
     * @param wb
     * @param gettableNum
     * @param sheetname
     */
    private void createTable(XSSFWorkbook wb,int gettableNum,String sheetname) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 21, i*22);
        }
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFRow row = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("质量结论（合格/不合格）".equals(row.getCell(19).toString())) {
                for (int j = 1;j<16;j++){
                    sheet.getRow(i+j).getCell(19).setCellType(CellType.BLANK);
                }
            }

        }
        if(gettableNum > 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 20, 0, gettableNum*22-1);
        }
    }

    /**
     * 分部
     * @param data
     * @return
     */
    private int gettableNum(List<Map<String, Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("fbgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%15==0){
                num += value/15;
            }else {
                num += value/15+1;
            }
        }
        return num;
    }

    /**
     * 查找文件
     * @param folderPath
     * @return
     */
    private static List<String> filterFiles(String folderPath) {
        List<String> matchingFiles = new ArrayList<>();

        File folder = new File(folderPath);
        File[] files = folder.listFiles();

        if (files != null) {
            for (File file : files) {
                //if (file.isFile() && !file.getName().contains("桥梁") && !file.getName().contains("隧道")) {
                    matchingFiles.add(file.getName());
                //}
            }
        }
        return matchingFiles;
    }


    @Autowired
    private JjgLqsQlService jjgLqsQlService;

    @Override
    public void generateBGZBG(String proname, String userid) throws IOException {

        List<JjgHtd> htdlist = jjgHtdService.gethtd(proname);

        List<Map<String,Object>> ljlist = new ArrayList<>();
        List<Map<String,Object>> qllist = new ArrayList<>();
        List<Map<String,Object>> fhllist = new ArrayList<>();

        List<Map<String, Object>> tsflist = new ArrayList<>();
        List<Map<String, Object>> pslist = new ArrayList<>();
        List<Map<String, Object>> xqlist = new ArrayList<>();
        List<Map<String, Object>> hdlist = new ArrayList<>();
        List<Map<String, Object>> zdlist = new ArrayList<>();

        List<Map<String, Object>> lmysdlist = new ArrayList<>();
        List<Map<String, Object>> lmwclist = new ArrayList<>();
        List<Map<String, Object>> lmwclallist = new ArrayList<>();
        List<Map<String, Object>> lmwclcflist = new ArrayList<>();
        List<Map<String, Object>> lmssxslist = new ArrayList<>();
        List<Map<String, Object>> lmhntqdlist = new ArrayList<>();
        List<Map<String, Object>> lmhntxlbgclist = new ArrayList<>();
        List<Map<String, Object>> lmpzdlist = new ArrayList<>();
        List<Map<String, Object>> lmczlist = new ArrayList<>();
        List<Map<String, Object>> lmhplist = new ArrayList<>();
        List<Map<String, Object>> ldhdlist = new ArrayList<>();
        List<Map<String, Object>> khlist = new ArrayList<>();
        List<Map<String, Object>> zxfhdlist = new ArrayList<>();
        List<Map<String, Object>> lmhdlist = new ArrayList<>();
        List<Map<String, Object>> qmpzdlist = new ArrayList<>();
        List<Map<String, Object>> qmhphplist = new ArrayList<>();
        List<Map<String, Object>> qmkhlist = new ArrayList<>();
        List<Map<String, Object>> sdysdlist = new ArrayList<>();
        List<Map<String, Object>> sdczlist = new ArrayList<>();
        List<Map<String, Object>> sdssxslist = new ArrayList<>();
        List<Map<String, Object>> sdhntqdlist = new ArrayList<>();
        List<Map<String, Object>> sdhntxlbgclist = new ArrayList<>();
        List<Map<String, Object>> sdpzdlist = new ArrayList<>();
        List<Map<String, Object>> sdkhlist = new ArrayList<>();
        List<Map<String, Object>> sdzxfhdlist = new ArrayList<>();
        List<Map<String, Object>> sdldhdlist = new ArrayList<>();
        List<Map<String, Object>> sdhdlist = new ArrayList<>();
        List<Map<String, Object>> sdhplist = new ArrayList<>();
        List<Map<String, Object>> qlxblist = new ArrayList<>();
        List<Map<String, Object>> qlsblist = new ArrayList<>();
        List<Map<String, Object>> qmxlist = new ArrayList<>();
        List<Map<String, Object>> cqlist = new ArrayList<>();
        List<Map<String, Object>> ztlist = new ArrayList<>();
        List<Map<String, Object>> sdlmlist = new ArrayList<>();
        List<Map<String, Object>> sdgcsdzxfhdlist = new ArrayList<>();
        List<Map<String, Object>> bzlist = new ArrayList<>();
        List<Map<String, Object>> bxlist = new ArrayList<>();
        List<Map<String, Object>> jafhllist = new ArrayList<>();
        List<Map<String, Object>> qdcclist = new ArrayList<>();
        List<Map<String, Object>> sdcjqdlist = new ArrayList<>();
        List<Map<String, Object>> qmxhzlist = new ArrayList<>();
        List<Map<String, Object>> sdlmhzblist = new ArrayList<>();



        List<Map<String, Object>> ljhzlist = new ArrayList<>();
        List<Map<String, Object>> lmhzlist = new ArrayList<>();
        List<Map<String, Object>> xmcclist = new ArrayList<>();
        List<Map<String, Object>> qlgclist = new ArrayList<>();
        List<Map<String, Object>> sdgclist = new ArrayList<>();
        List<Map<String, Object>> jabzlist = new ArrayList<>();
        List<Map<String, Object>> ysdfxblist = new ArrayList<>();
        List<Map<String, Object>> qlszdfxblist = new ArrayList<>();
        List<Map<String, Object>> sdcqfxblist = new ArrayList<>();
        List<Map<String, Object>> ljhzblist = new ArrayList<>();
        List<Map<String, Object>> qlpdlist = new ArrayList<>();
        List<Map<String, Object>> sdpdlist = new ArrayList<>();
        List<Map<String, Object>> htdpdlist = new ArrayList<>();
        List<Map<String, Object>> xsxmpdlist = new ArrayList<>();
        List<Map<String, Object>> lmhdldflist = new ArrayList<>();

        for (JjgHtd jjgHtd : htdlist) {
            String htd = jjgHtd.getName();
            String lx = jjgHtd.getLx();
            CommonInfoVo commonInfoVo = new CommonInfoVo();
            commonInfoVo.setProname(proname);
            commonInfoVo.setHtd(htd);
            commonInfoVo.setUserid(userid);
            if (lx.contains("路基工程")){
                //表3.4.1-1 抽查统计表
                List<Map<String,Object>> ljhtdlist = getLjcjtjbData(proname,htd);
                ljlist.addAll(ljhtdlist);
                //表4.1.1-1
                List<Map<String, Object>> tsf = gettsfData(commonInfoVo);
                tsflist.addAll(tsf);
                //表4.1.1-2
                List<Map<String, Object>> ps = getpsData(commonInfoVo);
                pslist.addAll(ps);
                //表4.1.1-3
                List<Map<String, Object>> xq = getxqData(commonInfoVo);
                xqlist.addAll(xq);
                //表4.1.1-4
                List<Map<String, Object>> hd = gethdData(commonInfoVo);
                hdlist.addAll(hd);
                //表4.1.1-4
                List<Map<String, Object>> zd = getzdData(commonInfoVo);
                zdlist.addAll(zd);

                //表5.1.1-1  路基单位工程质量评定汇总表
                List<Map<String, Object>> ljhzb = getljhzbData(commonInfoVo);
                ljhzblist.addAll(ljhzb);
                //表4.1.1-6
                List<Map<String,Object>> ljjchzlist = getljhzData(commonInfoVo);
                ljhzlist.addAll(ljjchzlist);


            }
            if (lx.contains("路面工程")){
                //表4.1.2-1
                List<Map<String, Object>> lmysd = getlmysdData(commonInfoVo);
                lmysdlist.addAll(lmysd);

                List<Map<String, Object>> lmwcall = getlmwcallData(commonInfoVo);
                lmwclallist.addAll(lmwcall);
                //表4.1.2-2
                List<Map<String, Object>> lmwc = getlmwcData(commonInfoVo);
                lmwclist.addAll(lmwc);

                List<Map<String, Object>> lmwclcf = getlmwclcfData(commonInfoVo);
                lmwclcflist.addAll(lmwclcf);
                //车辙
                List<Map<String, Object>> cz = getczData(commonInfoVo);
                lmczlist.addAll(cz);

                //表4.1.2-3
                List<Map<String, Object>> lmssxs = getlmssxsData(commonInfoVo);
                lmssxslist.addAll(lmssxs);

                //表4.1.2-4
                List<Map<String, Object>> hntqd = gethntqdData(commonInfoVo);
                lmhntqdlist.addAll(hntqd);

                //表4.1.2-5
                List<Map<String, Object>> hntxlbgc = gethntxlbgcData(commonInfoVo);
                lmhntxlbgclist.addAll(hntxlbgc);

                //表4.1.2-6 平整度
                List<Map<String, Object>> pzd = getpzdData(commonInfoVo);
                lmpzdlist.addAll(pzd);

                //表4.1.2-7 抗滑
                List<Map<String, Object>> kh = getkhData(commonInfoVo);
                khlist.addAll(kh);

                //表4.1.2-8(1) 钻芯法厚度
                List<Map<String, Object>> zxfhd = getzxfhdData(commonInfoVo);
                zxfhdlist.addAll(zxfhd);

                //表4.1.2-8(2) 路面雷达厚度
                List<Map<String, Object>> ldhd = getldhdData(commonInfoVo);
                ldhdlist.addAll(ldhd);

                //表4.1.2-8(3) 厚度
                List<Map<String, Object>> lmhd = getlmhdData(commonInfoVo);
                lmhdlist.addAll(lmhd);


                //表4.1.2-9 横坡
                List<Map<String, Object>> hp = gethpData(commonInfoVo);
                lmhplist.addAll(hp);

                //表4.1.2-10 汇总
                List<Map<String,Object>> lmjchzlist = getlmhzData(commonInfoVo);
                lmhzlist.addAll(lmjchzlist);

                //桥面平整度
                List<Map<String, Object>> qmpzd = getqmpzdData(commonInfoVo);
                qmpzdlist.addAll(qmpzd);

                //桥面横坡
                List<Map<String, Object>> qmhp = getqmhpData(commonInfoVo);
                qmhphplist.addAll(qmhp);

                //桥面抗滑
                List<Map<String, Object>> qmkh = getqmkhData(commonInfoVo);
                qmkhlist.addAll(qmkh);

                //桥面汇总
                List<Map<String,Object>> lmqlhzlist = getlmqlhzData(commonInfoVo);
                qmxhzlist.addAll(lmqlhzlist);

                //隧道路面压实度
                List<Map<String, Object>> sdysd = getsdysdData(commonInfoVo);
                sdysdlist.addAll(sdysd);

                //隧道路面车辙
                List<Map<String, Object>> sdcz = getsdczData(commonInfoVo);
                sdczlist.addAll(sdcz);

                //隧道路面渗水系数
                List<Map<String, Object>> sdssxs = getsdssxsData(commonInfoVo);
                sdssxslist.addAll(sdssxs);

                //隧道混凝土路面强度
                List<Map<String, Object>> sdhntqd = getsdhntqdData(commonInfoVo);
                if (CollectionUtils.isNotEmpty(sdhntqd)){
                    sdhntqdlist.addAll(sdhntqd);
                }

                //隧道混凝土路面相邻板高差检
                List<Map<String, Object>> sdhntxlbgc = getsdhntxlbgcData(commonInfoVo);
                if (CollectionUtils.isNotEmpty(sdhntxlbgc)){
                    sdhntxlbgclist.addAll(sdhntxlbgc);
                }

                //隧道路面平整度
                List<Map<String, Object>> sdpzd = getsdpzdData(commonInfoVo);
                sdpzdlist.addAll(sdpzd);

                //隧道路面抗滑
                List<Map<String, Object>> sdkh = getsdkhData(commonInfoVo);
                sdkhlist.addAll(sdkh);

                //隧道路面钻芯法厚度  表4.1.2-21（1）  隧道路面钻芯法厚度检测结果汇总表
                List<Map<String, Object>> sdzxfhd = getsdzxfhdData(commonInfoVo);
                sdzxfhdlist.addAll(sdzxfhd);

                //隧道路面雷达厚度表4.1.2-21（2）
                List<Map<String, Object>> sdldhd = getsdldhdData(commonInfoVo);
                sdldhdlist.addAll(sdldhd);

                //隧道路面厚度 表4.1.2-21（3）
                List<Map<String, Object>> sdhd = getsdhdData(commonInfoVo);
                sdhdlist.addAll(sdhd);

                //隧道路面横坡
                List<Map<String, Object>> sdhp = getsdhpData(commonInfoVo);
                sdhplist.addAll(sdhp);

                // 隧道路面汇总
                List<Map<String,Object>> xmcc = getxmccData(commonInfoVo);
                xmcclist.addAll(xmcc);


                List<Map<String,Object>> sdlmhzlist = getsdlmhzData(commonInfoVo);
                if (sdlmhzlist !=null && sdlmhzlist.size()>0){
                    sdlmhzblist.addAll(sdlmhzlist);
                }

            }
            if (lx.contains("桥梁工程")){

                List<Map<String,Object>> ljhtdlist = getLjqlcjData(proname, htd);
                qllist.addAll(ljhtdlist);

                //表4.1.3-1 桥梁下部检测结果汇总表
                List<Map<String,Object>> qlxb = getqlxbData(commonInfoVo);
                qlxblist.addAll(qlxb);

                //表4.1.3-2 桥梁上部检测结果汇总表
                List<Map<String,Object>> qlsb = getqlsbData(commonInfoVo);
                qlsblist.addAll(qlsb);

                List<Map<String,Object>> qmx = getqmxData(commonInfoVo);
                qmxlist.addAll(qmx);

                List<Map<String, Object>> qlpd = getqlpdData(commonInfoVo);
                qlpdlist.addAll(qlpd);

                List<Map<String,Object>> qlgc = getqlgcData(commonInfoVo);
                qlgclist.addAll(qlgc);



            }
            if (lx.contains("交安工程")){
                //表3.4.3-1 抽检
                List<Map<String,Object>> ljhtdlist = getLjjaData(proname, htd);
                fhllist.addAll(ljhtdlist);

                //表4.1.5-1  标志检测结果汇总表
                List<Map<String,Object>> bz = getbzData(commonInfoVo);
                bzlist.addAll(bz);

                //表4.1.5-2  标线检测结果汇总表
                List<Map<String,Object>> bx = getbxData(commonInfoVo);
                bxlist.addAll(bx);

                //表4.1.5-3  防护栏（波形梁）检测结果汇总表
                List<Map<String,Object>> fhl = getfhlData(commonInfoVo);
                jafhllist.addAll(fhl);

                //表4.1.5-4  防护栏（砼防护栏）检测结果汇总表
                List<Map<String,Object>> qdcc = getqdccData(commonInfoVo);
                qdcclist.addAll(qdcc);

                List<Map<String,Object>> jabz = getjabzData(commonInfoVo);
                jabzlist.addAll(jabz);
            }
            if (lx.contains("隧道工程")){

                List<Map<String,Object>> sdcjlist = getsdcjData(commonInfoVo);
                sdcjqdlist.addAll(sdcjlist);

                //表4.1.4-1 衬砌检测结果汇总表
                List<Map<String,Object>> cq = getcqData(commonInfoVo);
                cqlist.addAll(cq);

                //4.1.4-2总体检测结果汇总表
                List<Map<String,Object>> zt = getztData(commonInfoVo);
                ztlist.addAll(zt);

                //表4.1.4-3隧道路面检测结果汇总表
                List<Map<String,Object>> sdlm = getsdlmData(commonInfoVo);
                sdlmlist.addAll(sdlm);

                //表4.1.4-2-1 隧道钻芯法厚度检测结果汇总表
                List<Map<String,Object>> sdgcsdzxfhd = getsdgcsdzxfhdData(commonInfoVo);
                sdgcsdzxfhdlist.addAll(sdgcsdzxfhd);

                List<Map<String, Object>> sdpd = getsdpdData(commonInfoVo);
                sdpdlist.addAll(sdpd);

                List<Map<String,Object>> sdgc = getsdgcData(commonInfoVo);
                sdgclist.addAll(sdgc);

            }




            // 表4.1.6-1  路基工程压实度检测数据分析表
            List<Map<String,Object>> ysdfxb = getysdfxbData(commonInfoVo);
            ysdfxblist.addAll(ysdfxb);

            List<Map<String,Object>> qlszdfxb = getqlszdfxbData(commonInfoVo);
            qlszdfxblist.addAll(qlszdfxb);

            List<Map<String,Object>> sdcqfxb = getsdcqfxbData(commonInfoVo);
            sdcqfxblist.addAll(sdcqfxb);

            List<Map<String,Object>> htdpd = gethtdpdData(commonInfoVo);
            htdpdlist.addAll(htdpd);

        }

        List<Map<String,Object>> xsxmpd = getxsxmpdData(proname);
        xsxmpdlist.addAll(xsxmpd);

        XSSFWorkbook wb = null;
        File f = new File(filepath + File.separator + proname + File.separator + "报告中表格.xlsx");
        File fdir = new File(filepath + File.separator + proname);
        if (!fdir.exists()) {
            fdir.mkdirs();
        }
        //try {
            /*File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "报告中表格.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));*/
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/报告中表格.xlsx");
            Files.copy(inputStream, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            if (CollectionUtils.isNotEmpty(ljlist)){
                DBExcelData(wb,ljlist,proname);
            }
            if (CollectionUtils.isNotEmpty(qllist)){
                DBExcelQlData(wb,qllist,proname);
            }
            if (CollectionUtils.isNotEmpty(fhllist)){
                DBExcelJaData(wb,fhllist,proname);
            }
            if (CollectionUtils.isNotEmpty(sdcjqdlist)){
                DBExcelSdcjData(wb,sdcjqdlist);
            }
            if (CollectionUtils.isNotEmpty(tsflist)){
                DBExceltsfData(wb,tsflist);
            }
            if (CollectionUtils.isNotEmpty(pslist)){
                DBExcelpsData(wb,pslist);
            }

            if (CollectionUtils.isNotEmpty(xqlist)){
                DBExcelxqData(wb,xqlist);
            }

            if (CollectionUtils.isNotEmpty(hdlist)){
                DBExcelhdData(wb,hdlist);
            }

            if (CollectionUtils.isNotEmpty(zdlist)){
                DBExcelzdData(wb,zdlist);
            }

            if (CollectionUtils.isNotEmpty(ljhzlist)){
                DBExcelljhzData(wb,ljhzlist);
            }
            if (CollectionUtils.isNotEmpty(lmysdlist)){
                DBExcellmysdData(wb,lmysdlist);
            }

            if (CollectionUtils.isNotEmpty(lmwclist)){
                DBExcellmwcdData(wb,lmwclist);
            }
            if (CollectionUtils.isNotEmpty(lmwclallist)){
                DBExcellmwcalldData(wb,lmwclallist);
            }
            if (CollectionUtils.isNotEmpty(lmwclcflist)){
                DBExcellmwclcfdData(wb,lmwclcflist);
            }
            if (CollectionUtils.isNotEmpty(lmczlist)){
                DBExcellmczData(wb,lmczlist);
            }
            if (CollectionUtils.isNotEmpty(lmssxslist)){
                DBExcellmssxsData(wb,lmssxslist);
            }
            if (CollectionUtils.isNotEmpty(lmhntqdlist)){
                DBExcellmhntqdData(wb,lmhntqdlist);
            }
            if (CollectionUtils.isNotEmpty(lmhntxlbgclist)){
                DBExcellmhntxlbgcData(wb,lmhntxlbgclist);
            }
            if (CollectionUtils.isNotEmpty(lmpzdlist)){
                DBExcellmpzdData(wb,lmpzdlist);
            }
            if (CollectionUtils.isNotEmpty(khlist)){
                DBExcelkhData(wb,khlist);
            }
            if (CollectionUtils.isNotEmpty(zxfhdlist)){
                DBExcelzxfhdData(wb,zxfhdlist);
            }
            if (CollectionUtils.isNotEmpty(ldhdlist)){
                DBExcelldhdData(wb,ldhdlist);
            }
            if (CollectionUtils.isNotEmpty(lmhdlist)){
                DBExcellmhdData(wb,lmhdlist);
            }
            if (CollectionUtils.isNotEmpty(lmhplist)){
                DBExcellmhpData(wb,lmhplist);
            }
            if (CollectionUtils.isNotEmpty(lmhzlist)){
                DBExcellmhzData(wb,lmhzlist);
            }
            if (CollectionUtils.isNotEmpty(qmpzdlist)){
                DBExcelqmpzdData(wb,qmpzdlist);
            }
            if (CollectionUtils.isNotEmpty(qmhphplist)){
                DBExcelqmhpdData(wb,qmhphplist);
            }
            if (CollectionUtils.isNotEmpty(qmkhlist)){
                DBExcelqmkhData(wb,qmkhlist);
            }
            if (CollectionUtils.isNotEmpty(qmxhzlist)){
                DBExcelqmxhzData(wb,qmxhzlist);
            }
            if (CollectionUtils.isNotEmpty(sdysdlist)){
                DBExcelsdysdData(wb,sdysdlist);
            }
            if (CollectionUtils.isNotEmpty(sdczlist)){
                DBExcelsdczData(wb,sdczlist);
            }
            if (CollectionUtils.isNotEmpty(sdssxslist)){
                DBExcelsdssxsData(wb,sdssxslist);
            }
            if (CollectionUtils.isNotEmpty(sdhntqdlist)){
                DBExcelsdhntqdData(wb,sdhntqdlist);
            }
            if (CollectionUtils.isNotEmpty(sdhntxlbgclist)){
                DBExcelsdhntxlbgcData(wb,sdhntxlbgclist);
            }
            if (CollectionUtils.isNotEmpty(sdpzdlist)){
                DBExcelsdpzdData(wb,sdpzdlist);
            }
            if (CollectionUtils.isNotEmpty(sdkhlist)){
                DBExcelsdkhData(wb,sdkhlist);
            }
            if (CollectionUtils.isNotEmpty(sdzxfhdlist)){
                DBExcelsdzxfhdData(wb,sdzxfhdlist);
            }
            if (CollectionUtils.isNotEmpty(sdldhdlist)){
                DBExcelsdldhdData(wb,sdldhdlist);
            }
            if (CollectionUtils.isNotEmpty(sdhdlist)){
                DBExcelsdhdData(wb,sdhdlist);
            }
            if (CollectionUtils.isNotEmpty(sdhplist)){
                DBExcelsdhpData(wb,sdhplist);
            }
            if (CollectionUtils.isNotEmpty(sdlmhzblist)){
                DBExcelsdlmhzData(wb,sdlmhzblist);
            }
            if (CollectionUtils.isNotEmpty(xmcclist)){
                DBExcelxmccData(wb,xmcclist);
            }
            if (CollectionUtils.isNotEmpty(qlxblist)){
                DBExcelqlxbData(wb,qlxblist);
            }
            if (CollectionUtils.isNotEmpty(qlsblist)){
                DBExcelqlsbData(wb,qlsblist);
            }
            if (CollectionUtils.isNotEmpty(qmxlist)){
                DBExcelqmxData(wb,qmxlist);
            }
            if (CollectionUtils.isNotEmpty(qlgclist)){
                DBExcelqlpdData(wb,qlgclist);
            }
            if (CollectionUtils.isNotEmpty(cqlist)){
                DBExcelsdcqData(wb,cqlist);
            }
            if (CollectionUtils.isNotEmpty(ztlist)){
                DBExcelsdztData(wb,ztlist);
            }
            if (CollectionUtils.isNotEmpty(sdgcsdzxfhdlist)){
                DBExcelsdgcsdzxfhdData(wb,sdgcsdzxfhdlist);
            }
            if (CollectionUtils.isNotEmpty(sdlmlist)){
                DBExcelsdlmData(wb,sdlmlist);
            }
            if (CollectionUtils.isNotEmpty(sdgclist)){
                DBExcelsdgcData(wb,sdgclist);
            }
            if (CollectionUtils.isNotEmpty(bzlist)){
                DBExceljabzData(wb,bzlist);
            }
            if (CollectionUtils.isNotEmpty(bxlist)){
                DBExceljabxData(wb,bxlist);
            }
            if (CollectionUtils.isNotEmpty(jafhllist)){
                DBExceljafhlData(wb,jafhllist);
            }
            if (CollectionUtils.isNotEmpty(qdcclist)){
                DBExcelqdccData(wb,qdcclist);
            }
            if (CollectionUtils.isNotEmpty(jabzlist)){
                DBExceljabzhzData(wb,jabzlist);
            }
            if (CollectionUtils.isNotEmpty(ysdfxblist)){
                DBExcelysdfxData(wb,ysdfxblist);
            }
            if (CollectionUtils.isNotEmpty(qlszdfxblist)){
                DBExcelqlszdfxbData(wb,qlszdfxblist);
            }
            if (CollectionUtils.isNotEmpty(sdcqfxblist)){
                DBExcelsdcqfxData(wb,sdcqfxblist);
            }

            //删除空的工作簿
            JjgFbgcCommonUtils.hiddenSheets(wb);

            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();
        /*}catch (Exception e) {
            if(f.exists()){
                f.delete();
            }
            throw new JjgysException(20001, "生成报告中表格错误，请检查数据的正确性");
        }*/


    }

    /**
     * 表4.1.4-2-1 隧道钻芯法厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdgcsdzxfhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list1 = jjgFbgcSdgcGssdlqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            Double sMax = 0.0,zMax = 0.0;
            Double sMin = Double.MAX_VALUE,zMin = Double.MAX_VALUE;
            String zdbz = "",ydbz = "",zzdbz = "",yzdbz = "",smcsjz = "",zhdsjz = "";
            double smcjcds = 0,smchgs = 0,zhdjcds = 0,zhdhgs = 0;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                a = true;
                double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                sMax = (max > sMax) ? max : sMax;

                double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                sMin = (min < sMin) ? min : sMin;

                double max1 = Double.valueOf(map.get("总厚度平均值最大值").toString());
                zMax = (max1 > zMax) ? max1 : zMax;

                double min1 = Double.valueOf(map.get("总厚度平均值最小值").toString());
                zMin = (min1 < zMin) ? min1 : zMin;

                if (lmlx.equals("隧道左幅")){
                    zdbz = map.get("上面层代表值").toString();
                    zzdbz = map.get("总厚度代表值").toString();
                }
                if (lmlx.equals("隧道右幅")){
                    ydbz = map.get("上面层代表值").toString();
                    yzdbz = map.get("总厚度代表值").toString();
                }
                smcsjz = map.get("上面层设计值").toString();
                zhdsjz = map.get("总厚度设计值").toString();
                smcjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                smchgs += Double.valueOf(map.get("上面层厚度合格点数").toString());
                zhdjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                zhdhgs += Double.valueOf(map.get("总厚度合格点数").toString());
            }
            if (a){
                Map map = new HashMap();
                map.put("htd",commonInfoVo.getHtd());
                map.put("sdzxfhdlmlx","沥青路面");
                map.put("sdzxfhdlb","上面层厚度");
                map.put("sdzxfhdpjzfw",sMin+"~"+sMax);
                map.put("sdzxfhdzdbz",zdbz);
                map.put("sdzxfhdydbz",ydbz);
                map.put("sdzxfhdsjz",smcsjz);
                map.put("sdzxfhdjcds",decf.format(smcjcds));
                map.put("sdzxfhdhgs",decf.format(smchgs));
                map.put("sdzxfhdhgl",smcjcds!=0 ? df.format(smchgs/smcjcds*100) : 0);
                resultList.add(map);
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdzxfhdlmlx","沥青路面");
                map1.put("sdzxfhdlb","总厚度");
                map1.put("sdzxfhdpjzfw",zMin+"~"+zMax);
                map1.put("sdzxfhdzdbz",zzdbz);
                map1.put("sdzxfhdydbz",yzdbz);
                map1.put("sdzxfhdsjz",zhdsjz);
                map1.put("sdzxfhdjcds",decf.format(zhdjcds));
                map1.put("sdzxfhdhgs",decf.format(zhdhgs));
                map1.put("sdzxfhdhgl",zhdjcds!=0 ? df.format(zhdhgs/zhdjcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list2 = jjgFbgcSdgcSdhntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list2 != null && list2.size()>0){
            double jcds = 0;
            double hgds = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list2) {
                a = true;
                jcds += Double.valueOf(map.get("检测点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());

                double max = Double.valueOf(map.get("最大值").toString());
                lmMax = (max > lmMax) ? max : lmMax;
                double min = Double.valueOf(map.get("最小值").toString());
                lmMin = (min < lmMin) ? min : lmMin;

            }
            if (a){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdzxfhdlmlx","混凝土路面");
                map2.put("sdzxfhdlb","总厚度");
                map2.put("sdzxfhdpjzfw",lmMin+"~"+lmMax);
                map2.put("sdzxfhdzdbz",list2.get(0).get("代表值"));
                map2.put("sdzxfhdydbz",list2.get(0).get("代表值"));
                map2.put("sdzxfhdsjz",list2.get(0).get("设计值"));
                map2.put("sdzxfhdjcds",decf.format(jcds));
                map2.put("sdzxfhdhgs",decf.format(hgds));
                map2.put("sdzxfhdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map2);
            }
        }
        return resultList;

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdcqfxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.6-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("检测项目")) {
                sheet.getRow(index).getCell(1).setCellValue(datum.get("检测项目").toString());
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }
            if (datum.containsKey("检测总点数")) {
                if (!datum.get("检测总点数").equals("") && datum.get("检测总点数").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("检测总点数").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("合格点数")) {
                if (!datum.get("合格点数").equals("") && datum.get("合格点数").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("合格点数").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }
            if (datum.containsKey("合格率")) {
                if (!datum.get("合格率").equals("") && datum.get("合格率").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("合格率").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("ds")) {
                if (!datum.get("ds").equals("") && datum.get("ds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

            } else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlszdfxbData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.6-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            /*if (datum.containsKey("gdz")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("gdz").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }*/
            if (datum.containsKey("max")) {
                if (!datum.get("max").equals("") && datum.get("max").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("max").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }

            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("min")) {
                if (!datum.get("min").equals("") && datum.get("min").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("min").toString()));

                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("jcds")) {
                if (!datum.get("jcds").equals("") && datum.get("jcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("jcds").toString()));

                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("hgds")) {
                if (!datum.get("hgds").equals("") && datum.get("hgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hgds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("hgl")) {
                if (!datum.get("hgl").equals("") && datum.get("hgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hgl").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelysdfxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.6-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("bzz")) {
                if (!datum.get("bzz").equals("") && datum.get("bzz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("bzz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("jz")) {
                if (!datum.get("jz").equals("") && datum.get("jz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("jz").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("max")) {
                if (!datum.get("max").equals("") && datum.get("max").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("max").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }
            
            if (datum.containsKey("min")) {
                if (!datum.get("min").equals("") && datum.get("min").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("min").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            if (datum.containsKey("dbz")) {
                if (!datum.get("dbz").equals("") && datum.get("dbz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("dbz").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("jcds")) {
                if (!datum.get("jcds").equals("") && datum.get("jcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jcds").toString()));

                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            if (datum.containsKey("hgds")) {
                if (!datum.get("hgds").equals("") && datum.get("hgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("hgds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            if (datum.containsKey("hgl")) {
                if (!datum.get("hgl").equals("") && datum.get("hgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("hgl").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(8).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljabzhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-5");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("szdzds")){
                if (!datum.get("szdzds").equals("") && datum.get("szdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("szdzds").toString()));
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(3).createCell(index).setCellValue("");
            }

            if (datum.containsKey("szdhgds")){
                if (!datum.get("szdhgds").equals("") && datum.get("szdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(4).createCell(index).setCellValue("");
            }

            if (datum.containsKey("szdhgl")){
                if (!datum.get("szdhgl").equals("") && datum.get("szdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(5).createCell(index).setCellValue("");
            }

            if (datum.containsKey("jkzds")){
                if (!datum.get("jkzds").equals("") && datum.get("jkzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("jkzds").toString()));
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(6).createCell(index).setCellValue("");
            }

            if (datum.containsKey("jkhgds")){
                if (!datum.get("jkhgds").equals("") && datum.get("jkhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("jkhgds").toString()));
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(7).createCell(index).setCellValue("");
            }

            if (datum.containsKey("jkhgl")){
                if (!datum.get("jkhgl").equals("") && datum.get("jkhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("jkhgl").toString()));
                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(8).createCell(index).setCellValue("");
            }


            if (datum.containsKey("bzbhdzds")){
                if (!datum.get("bzbhdzds").equals("") && datum.get("bzbhdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("bzbhdzds").toString()));
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(9).createCell(index).setCellValue("");
            }

            if (datum.containsKey("bzbhdhgds")){
                if (!datum.get("bzbhdhgds").equals("") && datum.get("bzbhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("bzbhdhgds").toString()));
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(10).createCell(index).setCellValue("");
            }

            if (datum.containsKey("bzbhdhgl")){
                if (!datum.get("bzbhdhgl").equals("") && datum.get("bzbhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("bzbhdhgl").toString()));
                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(11).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xszds")){
                if (!datum.get("xszds").equals("") && datum.get("xszds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("xszds").toString()));
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(12).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xshgds")){
                if (!datum.get("xshgds").equals("") && datum.get("xshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(13).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xshgl")){
                if (!datum.get("xshgl").equals("") && datum.get("xshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(14).createCell(index).setCellValue("");
            }


            if (datum.containsKey("xzds")){
                if (!datum.get("xzds").equals("") && datum.get("xzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("xzds").toString()));
                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(15).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xhgds")){
                if (!datum.get("xhgds").equals("") && datum.get("xhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("xhgds").toString()));
                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(16).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xhgl")){
                if (!datum.get("xhgl").equals("") && datum.get("xhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("xhgl").toString()));
                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(17).createCell(index).setCellValue("");
            }

            if (datum.containsKey("bxhdzds")){
                if (!datum.get("bxhdzds").equals("") && datum.get("bxhdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("bxhdzds").toString()));
                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(18).createCell(index).setCellValue("");
            }

            if (datum.containsKey("bxhdhgds")){
                if (!datum.get("bxhdhgds").equals("") && datum.get("bxhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("bxhdhgds").toString()));
                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(19).createCell(index).setCellValue("");
            }

            if (datum.containsKey("bxhdhgl")){
                if (!datum.get("bxhdhgl").equals("") && datum.get("bxhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("bxhdhgl").toString()));
                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(20).createCell(index).setCellValue("");
            }

            if (datum.containsKey("jzds")){
                if (!datum.get("jzds").equals("") && datum.get("jzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("jzds").toString()));
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(21).createCell(index).setCellValue("");
            }

            if (datum.containsKey("jhgds")){
                if (!datum.get("jhgds").equals("") && datum.get("jhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("jhgds").toString()));
                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(22).createCell(index).setCellValue("");
            }

            if (datum.containsKey("jhgl")){
                if (!datum.get("jhgl").equals("") && datum.get("jhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("jhgl").toString()));
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(23).createCell(index).setCellValue("");
            }

            if (datum.containsKey("lzds")){
                if (!datum.get("lzds").equals("") && datum.get("lzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("lzds").toString()));
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(24).createCell(index).setCellValue("");
            }

            if (datum.containsKey("lhgds")){
                if (!datum.get("lhgds").equals("") && datum.get("lhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("lhgds").toString()));
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(25).createCell(index).setCellValue("");
            }

            if (datum.containsKey("lhgl")){
                if (!datum.get("lhgl").equals("") && datum.get("lhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("lhgl").toString()));
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(26).createCell(index).setCellValue("");
            }

            if (datum.containsKey("szds")){
                if (!datum.get("szds").equals("") && datum.get("szds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("szds").toString()));
                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(27).createCell(index).setCellValue("");
            }

            if (datum.containsKey("shgds")){
                if (!datum.get("shgds").equals("") && datum.get("shgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("shgds").toString()));
                }else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(28).createCell(index).setCellValue("");
            }

            if (datum.containsKey("shgl")){
                if (!datum.get("shgl").equals("") && datum.get("shgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("shgl").toString()));
                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(29).createCell(index).setCellValue("");
            }

            if (datum.containsKey("gzds")){
                if (!datum.get("gzds").equals("") && datum.get("gzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("gzds").toString()));
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(30).createCell(index).setCellValue("");
            }

            if (datum.containsKey("ghgds")){
                if (!datum.get("ghgds").equals("") && datum.get("ghgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("ghgds").toString()));
                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(31).createCell(index).setCellValue("");
            }

            if (datum.containsKey("ghgl")){
                if (!datum.get("ghgl").equals("") && datum.get("ghgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("ghgl").toString()));
                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(32).createCell(index).setCellValue("");
            }

            if (datum.containsKey("qdzds")){
                if (!datum.get("qdzds").equals("") && datum.get("qdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(33).createCell(index).setCellValue(Double.valueOf(datum.get("qdzds").toString()));
                }else {
                    sheet.getRow(33).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(33).createCell(index).setCellValue("");
            }
            if (datum.containsKey("qdhgs")){
                if (!datum.get("qdhgs").equals("") && datum.get("qdhgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(34).createCell(index).setCellValue(Double.valueOf(datum.get("qdhgs").toString()));
                }else {
                    sheet.getRow(34).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(34).createCell(index).setCellValue("");
            }

            if (datum.containsKey("qdhgl")){
                if (!datum.get("qdhgl").equals("") && datum.get("qdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(35).createCell(index).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
                }else {
                    sheet.getRow(35).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(35).createCell(index).setCellValue("");
            }

            if (datum.containsKey("cczds")){
                if (!datum.get("cczds").equals("") && datum.get("cczds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(36).createCell(index).setCellValue(Double.valueOf(datum.get("cczds").toString()));
                }else {
                    sheet.getRow(36).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(36).createCell(index).setCellValue("");
            }

            if (datum.containsKey("cchgs")){
                if (!datum.get("cchgs").equals("") && datum.get("cchgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(37).createCell(index).setCellValue(Double.valueOf(datum.get("cchgs").toString()));
                }else {
                    sheet.getRow(37).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(37).createCell(index).setCellValue("");
            }

            if (datum.containsKey("cchgl")){
                if (!datum.get("cchgl").equals("") && datum.get("cchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(38).createCell(index).setCellValue(Double.valueOf(datum.get("cchgl").toString()));
                }else {
                    sheet.getRow(38).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(38).createCell(index).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqdccData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("qdzds")) {
                if (!datum.get("qdzds").equals("") && datum.get("qdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("qdzds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("qdhgs")) {
                if (!datum.get("qdhgs").equals("") && datum.get("qdhgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qdhgs").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("qdhgl")) {
                if (!datum.get("qdhgl").equals("") && datum.get("qdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("cczds")) {
                if (!datum.get("cczds").equals("") && datum.get("cczds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cczds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("cchgs")) {
                if (!datum.get("cchgs").equals("") && datum.get("cchgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cchgs").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("cchgl")) {
                if (!datum.get("cchgl").equals("") && datum.get("cchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cchgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            index++;
        }


    }


    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljafhlData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("jzds")) {
                if (!datum.get("jzds").equals("") && datum.get("jzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("jzds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("jhgds")) {
                if (!datum.get("jhgds").equals("") && datum.get("jhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("jhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("jhgl")) {
                if (!datum.get("jhgl").equals("") && datum.get("jhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("jhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("lzds")) {
                if (!datum.get("lzds").equals("") && datum.get("lzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lzds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("lhgds")) {
                if (!datum.get("lhgds").equals("") && datum.get("lhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("lhgl")) {
                if (!datum.get("lhgl").equals("") && datum.get("lhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }


            if (datum.containsKey("szds")) {
                if (!datum.get("szds").equals("") && datum.get("szds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("szds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            if (datum.containsKey("shgds")) {
                if (!datum.get("shgds").equals("") && datum.get("shgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("shgds").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(8).setCellValue("");
            }

            if (datum.containsKey("shgl")) {
                if (!datum.get("shgl").equals("") && datum.get("shgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("shgl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(9).setCellValue("");
            }

            if (datum.containsKey("gzds")) {
                if (!datum.get("gzds").equals("") && datum.get("gzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("gzds").toString()));
                }else {
                    sheet.getRow(index).getCell(10).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(10).setCellValue("");
            }

            if (datum.containsKey("ghgds")) {
                if (!datum.get("ghgds").equals("") && datum.get("ghgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("ghgds").toString()));
                }else {
                    sheet.getRow(index).getCell(11).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(11).setCellValue("");
            }

            if (datum.containsKey("ghgl")) {
                if (!datum.get("ghgl").equals("") && datum.get("ghgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("ghgl").toString()));
                }else {
                    sheet.getRow(index).getCell(12).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(12).setCellValue("");
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljabxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("xzds")) {
                if (!datum.get("xzds").equals("") && datum.get("xzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("xzds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("xhgds")) {
                if (!datum.get("xhgds").equals("") && datum.get("xhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("xhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("xhgl")) {
                if (!datum.get("xhgl").equals("") && datum.get("xhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("xhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("bxhdzds")) {
                if (!datum.get("bxhdzds").equals("") && datum.get("bxhdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("bxhdzds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("bxhdhgds")) {
                if (!datum.get("bxhdhgds").equals("") && datum.get("bxhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("bxhdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("bxhdhgl")) {
                if (!datum.get("bxhdhgl").equals("") && datum.get("bxhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("bxhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljabzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        /**
         * [{htd=LJ-1, jkzds=26.0, jkhgds=24.0, jkhgl=92.31, szdzds=326.0, szdhgds=279.0,
         * szdhgl=85.58, hdzds=346.0, hdhgds=346.0, hdhgl=100.00, xszds=873, xshgds=873, xshgl=100.00}]
         */
        XSSFSheet sheet = wb.getSheet("表4.1.5-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("szdzds")) {
                if (!datum.get("szdzds").equals("") && datum.get("szdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("szdzds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("szdhgds")) {
                if (!datum.get("szdhgds").equals("") && datum.get("szdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("szdhgl")) {
                if (!datum.get("szdhgl").equals("") && datum.get("szdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("jkzds")){
                if (!datum.get("jkzds").equals("") && datum.get("jkzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("jkzds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("jkhgds")){
                if (!datum.get("jkhgds").equals("") && datum.get("jkhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("jkhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("jkhgl")){
                if (!datum.get("jkhgl").equals("") && datum.get("jkhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jkhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }


            if (datum.containsKey("bzbhdzds")) {
                if (!datum.get("bzbhdzds").equals("") && datum.get("bzbhdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("bzbhdzds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            if (datum.containsKey("bzbhdhgds")) {
                if (!datum.get("bzbhdhgds").equals("") && datum.get("bzbhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("bzbhdhgds").toString()));
                } else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(8).setCellValue("");
            }

            if (datum.containsKey("bzbhdhgl")) {
                if (!datum.get("bzbhdhgl").equals("") && datum.get("bzbhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("bzbhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(9).setCellValue("");
            }

            if (datum.containsKey("xszds")) {
                if (!datum.get("xszds").equals("") && datum.get("xszds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("xszds").toString()));
                }else {
                    sheet.getRow(index).getCell(10).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(10).setCellValue("");
            }

            if (datum.containsKey("xshgds")) {
                if (!datum.get("xshgds").equals("") && datum.get("xshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
                }else {
                    sheet.getRow(index).getCell(11).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(11).setCellValue("");
            }

            if (datum.containsKey("xshgl")) {
                if (!datum.get("xshgl").equals("") && datum.get("xshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
                }else {
                    sheet.getRow(index).getCell(12).setCellValue("");
                }
            } else {
                sheet.getRow(index).getCell(12).setCellValue("");
            }

            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());

                if (datum.containsKey("cqqdjcds")){
                    if (!datum.get("cqqdjcds").equals("") && datum.get("cqqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("cqqdjcds").toString()));
                    }else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("cqqdhgds")){
                    if (!datum.get("cqqdhgds").equals("") && datum.get("cqqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("cqqdhgds").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("cqqdhgl")){
                    if (!datum.get("cqqdhgl").equals("") && datum.get("cqqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("cqqdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }


                if (datum.containsKey("cqhdjcds")){
                    if (!datum.get("cqhdjcds").equals("") && datum.get("cqhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }

                if (datum.containsKey("cqhdhgds")){
                    if (!datum.get("cqhdhgds").equals("") && datum.get("cqhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
                    }else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

                if (datum.containsKey("cqhdhgl")){
                    if (!datum.get("cqhdhgl").equals("") && datum.get("cqhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("dmpzdjcds")){
                    if (!datum.get("dmpzdjcds").equals("") && datum.get("dmpzdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("dmpzdjcds").toString()));
                    }else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("dmpzdhgds")){
                    if (!datum.get("dmpzdhgds").equals("") && datum.get("dmpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("dmpzdhgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("dmpzdhgl")){
                    if (!datum.get("dmpzdhgl").equals("") && datum.get("dmpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("dmpzdhgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("kdjcds")){
                    if (!datum.get("kdjcds").equals("") && datum.get("kdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("kdjcds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }

                if (datum.containsKey("kdhgds")){
                    if (!datum.get("kdhgds").equals("") && datum.get("kdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("kdhgds").toString()));
                    }else {
                        sheet.getRow(13).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }

                if (datum.containsKey("kdhgl")){
                    if (!datum.get("kdhgl").equals("") && datum.get("kdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("kdhgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }

                if (datum.containsKey("jkjcds")){
                    if (!datum.get("jkjcds").equals("") && datum.get("jkjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("jkjcds").toString()));
                    }else {
                        sheet.getRow(15).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

                if (datum.containsKey("jkhgds")){
                    if (!datum.get("jkhgds").equals("") && datum.get("jkhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("jkhgds").toString()));
                    }else {
                        sheet.getRow(16).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }

                if (datum.containsKey("jkhgl")){
                    if (!datum.get("jkhgl").equals("") && datum.get("jkhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("jkhgl").toString()));
                    }
                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdjcds")){
                    if (!datum.get("ysdjcds").equals("") && datum.get("ysdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
                    }else {
                        sheet.getRow(18).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgds")){
                    if (!datum.get("ysdhgds").equals("") && datum.get("ysdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
                    }else {
                        sheet.getRow(19).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgl")){
                    if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                    }else {
                        sheet.getRow(20).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }

                //车辙
                if (datum.containsKey("czjcds")){
                    if (!datum.get("czjcds").equals("") && datum.get("czjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("czjcds").toString()));
                    }else {
                        sheet.getRow(21).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czhgds")){
                    if (!datum.get("czhgds").equals("") && datum.get("czhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
                    }else {
                        sheet.getRow(22).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czhgl")){
                    if (!datum.get("czhgl").equals("") && datum.get("czhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
                    }else {
                        sheet.getRow(23).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xsjcds")){
                    if (!datum.get("xsjcds").equals("") && datum.get("xsjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("xsjcds").toString()));
                    }else {
                        sheet.getRow(24).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xshgds")){
                    if (!datum.get("xshgds").equals("") && datum.get("xshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
                    }else {
                        sheet.getRow(25).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xshgl")){
                    if (!datum.get("xshgl").equals("") && datum.get("xshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
                    }else {
                        sheet.getRow(26).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qdjcds")){
                    if (!datum.get("qdjcds").equals("") && datum.get("qdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("qdjcds").toString()));
                    }else {
                        sheet.getRow(27).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }
                if (datum.containsKey("qdhgds")){
                    if (!datum.get("qdhgds").equals("") && datum.get("qdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("qdhgds").toString()));
                    }else {
                        sheet.getRow(28).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }
                if (datum.containsKey("qdhgl")){
                    if (!datum.get("qdhgl").equals("") && datum.get("qdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
                    }else {
                        sheet.getRow(29).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gcjcds")){
                    if (!datum.get("gcjcds").equals("") && datum.get("gcjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("gcjcds").toString()));
                    }else {
                        sheet.getRow(30).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gchgds")){
                    if (!datum.get("gchgds").equals("") && datum.get("gchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("gchgds").toString()));
                    }else {
                        sheet.getRow(31).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gchgl")){
                    if (!datum.get("gchgl").equals("") && datum.get("gchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("gchgl").toString()));
                    }else {
                        sheet.getRow(32).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }
                //sss
                if (datum.containsKey("pzdjcds")){
                    if (!datum.get("pzdjcds").equals("") && datum.get("pzdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(33).createCell(index).setCellValue(Double.valueOf(datum.get("pzdjcds").toString()));
                    }else {
                        sheet.getRow(33).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(33).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdhgds")){
                    if (!datum.get("pzdhgds").equals("") && datum.get("pzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(34).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhgds").toString()));
                    }else {
                        sheet.getRow(34).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(34).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdhgl")){
                    if (!datum.get("pzdhgl").equals("") && datum.get("pzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(35).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhgl").toString()));
                    }else {
                        sheet.getRow(35).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(35).createCell(index).setCellValue("");
                }

                //kang
                if (datum.containsKey("khjcds")){
                    if (!datum.get("khjcds").equals("") && datum.get("khjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(36).createCell(index).setCellValue(Double.valueOf(datum.get("khjcds").toString()));
                    }else {
                        sheet.getRow(36).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(36).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhgds")){
                    if (!datum.get("khhgds").equals("") && datum.get("khhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(37).createCell(index).setCellValue(Double.valueOf(datum.get("khhgds").toString()));
                    }else {
                        sheet.getRow(37).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(37).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhgl")){
                    if (!datum.get("khhgl").equals("") && datum.get("khhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(38).createCell(index).setCellValue(Double.valueOf(datum.get("khhgl").toString()));
                    }else {
                        sheet.getRow(38).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(38).createCell(index).setCellValue("");
                }


                if (datum.containsKey("hdjcds")){
                    if (!datum.get("hdjcds").equals("") && datum.get("hdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(39).createCell(index).setCellValue(Double.valueOf(datum.get("hdjcds").toString()));
                    }else {
                        sheet.getRow(39).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(39).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdhgds")){
                    if (!datum.get("hdhgds").equals("") && datum.get("hdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(40).createCell(index).setCellValue(Double.valueOf(datum.get("hdhgds").toString()));
                    }else {
                        sheet.getRow(40).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(40).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdhgl")){
                    if (!datum.get("hdhgl").equals("") && datum.get("hdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(41).createCell(index).setCellValue(Double.valueOf(datum.get("hdhgl").toString()));
                    }else {
                        sheet.getRow(41).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(41).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthdjcds")){
                    if (!datum.get("hnthdjcds").equals("") && datum.get("hnthdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(42).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdjcds").toString()));
                    }else {
                        sheet.getRow(42).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(42).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthdhgds")){
                    if (!datum.get("hnthdhgds").equals("") && datum.get("hnthdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(43).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdhgds").toString()));
                    }else {
                        sheet.getRow(43).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(43).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthdhgl")){
                    if (!datum.get("hnthdhgl").equals("") && datum.get("hnthdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(44).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdhgl").toString()));
                    }else {
                        sheet.getRow(44).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(44).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hpjcds")){
                    if (!datum.get("hpjcds").equals("") && datum.get("hpjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(45).createCell(index).setCellValue(Double.valueOf(datum.get("hpjcds").toString()));
                    }else {
                        sheet.getRow(45).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(45).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hphgds")){
                    if (!datum.get("hphgds").equals("") && datum.get("hphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(46).createCell(index).setCellValue(Double.valueOf(datum.get("hphgds").toString()));
                    }else {
                        sheet.getRow(46).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(46).createCell(index).setCellValue("");
                }
                if (datum.containsKey("hphgl")){
                    if (!datum.get("hphgl").equals("") && datum.get("hphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(47).createCell(index).setCellValue(Double.valueOf(datum.get("hphgl").toString()));
                    }else {
                        sheet.getRow(47).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(47).createCell(index).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdlmData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-3");
        int index = 2;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("ysdjcds")){
                if (!datum.get("ysdjcds").equals("") && datum.get("ysdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(3).createCell(index).setCellValue("");
            }

            if (datum.containsKey("ysdhgds")){
                if (!datum.get("ysdhgds").equals("") && datum.get("ysdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(4).createCell(index).setCellValue("");
            }

            if (datum.containsKey("ysdhgl")){
                if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }

            }else {
                sheet.getRow(5).createCell(index).setCellValue("");
            }

            //车辙
            if (datum.containsKey("czjcds")){
                if (!datum.get("czjcds").equals("") && datum.get("czjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("czjcds").toString()));
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(6).createCell(index).setCellValue("");
            }

            if (datum.containsKey("czhgds")){
                if (!datum.get("czhgds").equals("") && datum.get("czhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(7).createCell(index).setCellValue("");
            }

            if (datum.containsKey("czhgl")){
                if (!datum.get("czhgl").equals("") && datum.get("czhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(8).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xsjcds")){
                if (!datum.get("xsjcds").equals("") && datum.get("xsjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("xsjcds").toString()));
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(9).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xshgds")){
                if (!datum.get("xshgds").equals("") && datum.get("xshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(10).createCell(index).setCellValue("");
            }

            if (datum.containsKey("xshgl")){
                if (!datum.get("xshgl").equals("") && datum.get("xshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(11).createCell(index).setCellValue("");
            }

            if (datum.containsKey("qdjcds")){
                if (!datum.get("qdjcds").equals("") && datum.get("qdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("qdjcds").toString()));
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(12).createCell(index).setCellValue("");
            }

            if (datum.containsKey("qdhgds")){
                if (!datum.get("qdhgds").equals("") && datum.get("qdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("qdhgds").toString()));
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(13).createCell(index).setCellValue("");
            }

            if (datum.containsKey("qdhgl")){
                if (!datum.get("qdhgl").equals("") && datum.get("qdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(14).createCell(index).setCellValue("");
            }

            if (datum.containsKey("gcjcds")){
                if (!datum.get("gcjcds").equals("") && datum.get("gcjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("gcjcds").toString()));
                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(15).createCell(index).setCellValue("");
            }

            if (datum.containsKey("gchgds")){
                if (!datum.get("gchgds").equals("") && datum.get("gchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("gchgds").toString()));
                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(16).createCell(index).setCellValue("");
            }

            if (datum.containsKey("gchgl")){
                if (!datum.get("gchgl").equals("") && datum.get("gchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("gchgl").toString()));
                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(17).createCell(index).setCellValue("");
            }

            //平整度
            if (datum.containsKey("pzdjcds")){
                if (!datum.get("pzdjcds").equals("") && datum.get("pzdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("pzdjcds").toString()));
                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(18).createCell(index).setCellValue("");
            }

            if (datum.containsKey("pzdhgds")){
                if (!datum.get("pzdhgds").equals("") && datum.get("pzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhgds").toString()));
                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(19).createCell(index).setCellValue("");
            }

            if (datum.containsKey("pzdhgl")){
                if (!datum.get("pzdhgl").equals("") && datum.get("pzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhgl").toString()));
                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(20).createCell(index).setCellValue("");
            }

            //抗滑
            if (datum.containsKey("khjcds")){
                if (!datum.get("khjcds").equals("") && datum.get("khjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("khjcds").toString()));
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(21).createCell(index).setCellValue("");
            }

            if (datum.containsKey("khhgds")){
                if (!datum.get("khhgds").equals("") && datum.get("khhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("khhgds").toString()));
                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(22).createCell(index).setCellValue("");
            }

            if (datum.containsKey("khhgl")){
                if (!datum.get("khhgl").equals("") && datum.get("khhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("khhgl").toString()));
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(23).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hdjcds")){
                if (!datum.get("hdjcds").equals("") && datum.get("hdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("hdjcds").toString()));
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(24).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hdhgds")){
                if (!datum.get("hdhgds").equals("") && datum.get("hdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("hdhgds").toString()));
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(25).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hdhgl")){
                if (!datum.get("hdhgl").equals("") && datum.get("hdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("hdhgl").toString()));
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(26).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hnthdjcds")){
                if (!datum.get("hnthdjcds").equals("") && datum.get("hnthdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdjcds").toString()));
                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(27).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hnthdhgds")){
                if (!datum.get("hnthdhgds").equals("") && datum.get("hnthdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdhgds").toString()));
                }
                else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(28).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hnthdhgl")){
                if (!datum.get("hnthdhgl").equals("") && datum.get("hnthdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdhgl").toString()));
                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(29).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hpjcds")){
                if (!datum.get("hpjcds").equals("") && datum.get("hpjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("hpjcds").toString()));
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(30).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hphgds")){
                if (!datum.get("hphgds").equals("") && datum.get("hphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("hphgds").toString()));
                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(31).createCell(index).setCellValue("");
            }

            if (datum.containsKey("hphgl")){
                if (!datum.get("hphgl").equals("") && datum.get("hphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("hphgl").toString()));
                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }
            }else {
                sheet.getRow(32).createCell(index).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdztData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("kdjcds")){
                if (!datum.get("kdjcds").equals("") && datum.get("kdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("kdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("kdhgds")){
                if (!datum.get("kdhgds").equals("") && datum.get("kdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("kdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("kdhgl")){
                if (!datum.get("kdhgl").equals("") && datum.get("kdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("kdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("jkjcds")){
                if (!datum.get("jkjcds").equals("") && datum.get("jkjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("jkjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("jkhgds")){
                if (!datum.get("jkhgds").equals("") && datum.get("jkhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("jkhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("jkhgl")){
                if (!datum.get("jkhgl").equals("") && datum.get("jkhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jkhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdcqData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("cqqdjcds")){
                    if (!datum.get("cqqdjcds").equals("") && datum.get("cqqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("cqqdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
                if (datum.containsKey("cqqdhgds")){
                    if (!datum.get("cqqdhgds").equals("") && datum.get("cqqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("cqqdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

                if (datum.containsKey("cqqdhgl")){
                    if (!datum.get("cqqdhgl").equals("") && datum.get("cqqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("cqqdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("cqhdjcds")){
                    if (!datum.get("cqhdjcds").equals("") && datum.get("cqhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

                if (datum.containsKey("cqhdhgds")){
                    if (!datum.get("cqhdhgds").equals("") && datum.get("cqhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

                if (datum.containsKey("cqhdhgl")){
                    if (!datum.get("cqhdhgl").equals("") && datum.get("cqhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

                if (datum.containsKey("dmpzdjcds")){
                    if (!datum.get("dmpzdjcds").equals("") && datum.get("dmpzdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("dmpzdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }

                if (datum.containsKey("dmpzdhgds")){
                    if (!datum.get("dmpzdhgds").equals("") && datum.get("dmpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("dmpzdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }

                if (datum.containsKey("dmpzdhgl")){
                    if (!datum.get("dmpzdhgl").equals("") && datum.get("dmpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("dmpzdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(9).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlpdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.3-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());

                if (datum.containsKey("tqdjcds")) {
                    if (!datum.get("tqdjcds").equals("") && datum.get("tqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("tqdjcds").toString()));
                    } else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("tqdhgds")) {
                    if (!datum.get("tqdhgds").equals("") && datum.get("tqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("tqdhgds").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("tqdhgl")) {
                    if (!datum.get("tqdhgl").equals("") && datum.get("tqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("tqdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }

                if (datum.containsKey("jgccjcds")) {
                    if (!datum.get("jgccjcds").equals("") && datum.get("jgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("jgccjcds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }

                if (datum.containsKey("jgcchgds")) {
                    if (!datum.get("jgcchgds").equals("") && datum.get("jgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("jgcchgds").toString()));
                    } else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

                if (datum.containsKey("jgcchgl")) {
                    if (!datum.get("jgcchgl").equals("") && datum.get("jgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("jgcchgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("bhchdjcds")) {
                    if (!datum.get("bhchdjcds").equals("") && datum.get("bhchdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("bhchdjcds").toString()));
                    }else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("bhchdhgds")) {
                    if (!datum.get("bhchdhgds").equals("") && datum.get("bhchdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("bhchdhgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("bhchdhgl")) {
                    if (!datum.get("bhchdhgl").equals("") && datum.get("bhchdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("bhchdhgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("szdjcds")) {
                    if (!datum.get("szdjcds").equals("") && datum.get("szdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("szdjcds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }

                if (datum.containsKey("szdhgds")) {
                    if (!datum.get("szdhgds").equals("") && datum.get("szdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
                    }else {
                        sheet.getRow(13).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }

                if (datum.containsKey("szdhgl")) {
                    if (!datum.get("szdhgl").equals("") && datum.get("szdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbqdjcds")) {
                    if (!datum.get("qlsbqdjcds").equals("") && datum.get("qlsbqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbqdjcds").toString()));
                    }else {
                        sheet.getRow(15).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbqdhgds")) {
                    if (!datum.get("qlsbqdhgds").equals("") && datum.get("qlsbqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbqdhgds").toString()));
                    }else {
                        sheet.getRow(16).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbqdhgl")) {
                    if (!datum.get("qlsbqdhgl").equals("") && datum.get("qlsbqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbqdhgl").toString()));
                    }else {
                        sheet.getRow(17).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbjgccjcds")) {
                    if (!datum.get("qlsbjgccjcds").equals("") && datum.get("qlsbjgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbjgccjcds").toString()));
                    }else {
                        sheet.getRow(18).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbjgcchgds")) {
                    if (!datum.get("qlsbjgcchgds").equals("") && datum.get("qlsbjgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbjgcchgds").toString()));
                    }else {
                        sheet.getRow(19).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbjgcchgl")) {
                    if (!datum.get("qlsbjgcchgl").equals("") && datum.get("qlsbjgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbjgcchgl").toString()));
                    }else {
                        sheet.getRow(20).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbbhchdjcds")){
                    if (!datum.get("qlsbbhchdjcds").equals("") && datum.get("qlsbbhchdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbbhchdjcds").toString()));
                    }else {
                        sheet.getRow(21).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbbhchdhgds")){
                    if (!datum.get("qlsbbhchdhgds").equals("") && datum.get("qlsbbhchdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgds").toString()));
                    }else {
                        sheet.getRow(22).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qlsbbhchdhgl")){
                    if (!datum.get("qlsbbhchdhgl").equals("") && datum.get("qlsbbhchdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgl").toString()));
                    }else {
                        sheet.getRow(23).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmpzdzds")){
                    if (!datum.get("qmpzdzds").equals("") && datum.get("qmpzdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("qmpzdzds").toString()));
                    }else {
                        sheet.getRow(24).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmpzdhgds")){
                    if (!datum.get("qmpzdhgds").equals("") && datum.get("qmpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("qmpzdhgds").toString()));
                    }else {
                        sheet.getRow(25).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmpzdhgl")){
                    if (!datum.get("qmpzdhgl").equals("") && datum.get("qmpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("qmpzdhgl").toString()));
                    }else {
                        sheet.getRow(26).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }


                if (datum.containsKey("qmhpzds")){
                    if (!datum.get("qmhpzds").equals("") && datum.get("qmhpzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
                    }else {
                        sheet.getRow(27).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmhphgds")){
                    if (!datum.get("qmhphgds").equals("") && datum.get("qmhphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
                    }else {
                        sheet.getRow(28).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmhphgl")){
                    if (!datum.get("qmhphgl").equals("") && datum.get("qmhphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
                    }else {
                        sheet.getRow(29).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmkhgzsdzds")){
                    if (!datum.get("qmkhgzsdzds").equals("") && datum.get("qmkhgzsdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("qmkhgzsdzds").toString()));
                    }else {
                        sheet.getRow(30).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmkhgzsdhgds")){
                    if (!datum.get("qmkhgzsdhgds").equals("") && datum.get("qmkhgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgds").toString()));
                    }else {
                        sheet.getRow(31).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmkhgzsdhgl")){
                    if (!datum.get("qmkhgzsdhgl").equals("") && datum.get("qmkhgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgl").toString()));
                    }else {
                        sheet.getRow(32).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.3-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

                if (datum.containsKey("qmpzdzds")){
                    if (!datum.get("qmpzdzds").equals("") && datum.get("qmpzdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("qmpzdzds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }

                if (datum.containsKey("qmpzdhgds")){
                    if (!datum.get("qmpzdhgds").equals("") && datum.get("qmpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qmpzdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

                if (datum.containsKey("qmpzdhgl")){
                    if (!datum.get("qmpzdhgl").equals("") && datum.get("qmpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmpzdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("qmhpzds")){
                    if (!datum.get("qmhpzds").equals("") && datum.get("qmhpzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

                if (datum.containsKey("qmhphgds")){
                    if (!datum.get("qmhphgds").equals("") && datum.get("qmhphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

                if (datum.containsKey("qmhphgl")){
                    if (!datum.get("qmhphgl").equals("") && datum.get("qmhphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

                if (datum.containsKey("qmkhgzsdzds")){
                    if (!datum.get("qmkhgzsdzds").equals("") && datum.get("qmkhgzsdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qmkhgzsdzds").toString()));
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }

                if (datum.containsKey("qmkhgzsdhgds")){
                    if (!datum.get("qmkhgzsdhgds").equals("") && datum.get("qmkhgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }

                if (datum.containsKey("qmkhgzsdhgl")){
                    if (!datum.get("qmkhgzsdhgl").equals("") && datum.get("qmkhgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(9).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlsbData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.3-2");
        XSSFSheet sheet = wb.getSheet("表4.1.3-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("qlsbqdjcds")){
                    if (!datum.get("qlsbqdjcds").equals("") && datum.get("qlsbqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("qlsbqdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }

                if (datum.containsKey("qlsbqdhgds")){
                    if (!datum.get("qlsbqdhgds").equals("") && datum.get("qlsbqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qlsbqdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

                if (datum.containsKey("qlsbqdhgl")){
                    if (!datum.get("qlsbqdhgl").equals("") && datum.get("qlsbqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qlsbqdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("qlsbjgccjcds")){
                    if (!datum.get("qlsbjgccjcds").equals("") && datum.get("qlsbjgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qlsbjgccjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

                if (datum.containsKey("qlsbjgcchgds")){
                    if (!datum.get("qlsbjgcchgds").equals("") && datum.get("qlsbjgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qlsbjgcchgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

                if (datum.containsKey("qlsbjgcchgl")){
                    if (!datum.get("qlsbjgcchgl").equals("") && datum.get("qlsbjgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qlsbjgcchgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

                if (datum.containsKey("qlsbbhchdjcds")){
                    if (!datum.get("qlsbbhchdjcds").equals("") && datum.get("qlsbbhchdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qlsbbhchdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }

                if (datum.containsKey("qlsbbhchdhgds")){
                    if (!datum.get("qlsbbhchdhgds").equals("") && datum.get("qlsbbhchdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }

                if (datum.containsKey("qlsbbhchdhgl")){
                    if (!datum.get("qlsbbhchdhgl").equals("") && datum.get("qlsbbhchdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(9).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlxbData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.3-1");
        /**
         * {htd=TJ-1, tqdjcds=74, tqdhgds=74, tqdhgl=100.00, jgccjcds=60, jgcchgds=60, jgcchgl=100.00, bhchdjcds=444, bhchdhgds=345, bhchdhgl=77.70, szdjcds=40, szdhgds=35, szdhgl=87.50}
         */
        XSSFSheet sheet = wb.getSheet("表4.1.3-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (!datum.isEmpty()){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("tqdjcds")){
                    if (!datum.get("tqdjcds").equals("") && datum.get("tqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("tqdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }

                if (datum.containsKey("tqdhgds")){
                    if (!datum.get("tqdhgds").equals("") && datum.get("tqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("tqdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

                if (datum.containsKey("tqdhgl")){
                    if (!datum.get("tqdhgl").equals("") && datum.get("tqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("tqdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("jgccjcds")){
                    if (!datum.get("jgccjcds").equals("") && datum.get("jgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("jgccjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

                if (datum.containsKey("jgcchgds")){
                    if (!datum.get("jgcchgds").equals("") && datum.get("jgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("jgcchgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

                if (datum.containsKey("jgcchgl")){
                    if (!datum.get("jgcchgl").equals("") && datum.get("jgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jgcchgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

                if (datum.containsKey("bhchdjcds")){
                    if (!datum.get("bhchdjcds").equals("") && datum.get("bhchdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("bhchdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }

                if (datum.containsKey("bhchdhgds")){
                    if (!datum.get("bhchdhgds").equals("") && datum.get("bhchdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("bhchdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }

                if (datum.containsKey("bhchdhgl")){
                    if (!datum.get("bhchdhgl").equals("") && datum.get("bhchdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("bhchdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(9).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }

                if (datum.containsKey("szdjcds")){
                    if (!datum.get("szdjcds").equals("") && datum.get("szdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("szdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(10).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(10).setCellValue("");
                }

                if (datum.containsKey("szdhgds")){
                    if (!datum.get("szdhgds").equals("") && datum.get("szdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
                    } else {
                        sheet.getRow(index).getCell(11).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(11).setCellValue("");
                }

                if (datum.containsKey("szdhgl")){
                    if (!datum.get("szdhgl").equals("") && datum.get("szdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(12).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(12).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelxmccData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-24");
        int index = 4;
        for (Map<String, Object> datum : data) {
            if (datum.size()>1){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());

                if (datum.containsKey("ysdzds")) {
                    if (!datum.get("ysdzds").equals("") && datum.get("ysdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("ysdzds").toString()));
                    }else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgds")) {
                    if (!datum.get("ysdhgds").equals("") && datum.get("ysdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgl")) {
                    if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }

                if (datum.containsKey("wczds")) {
                    if (!datum.get("wczds").equals("") && datum.get("wczds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("wczds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }
                if (datum.containsKey("wchgds")) {
                    if (!datum.get("wchgds").equals("") && datum.get("wchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("wchgds").toString()));
                    }else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

                if (datum.containsKey("wchgl")) {
                    if (!datum.get("wchgl").equals("") && datum.get("wchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("wchgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czzds")) {
                    if (!datum.get("czzds").equals("") && datum.get("czzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("czzds").toString()));
                    } else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czhgds")) {
                    if (!datum.get("czhgds").equals("") && datum.get("czhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czhgl")) {
                    if (!datum.get("czhgl").equals("") && datum.get("czhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }
                } else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ssxszds")){
                    if (!datum.get("ssxszds").equals("") && datum.get("ssxszds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("ssxszds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ssxshgds")){
                    if (!datum.get("ssxshgds").equals("") && datum.get("ssxshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("ssxshgds").toString()));
                    }else {
                        sheet.getRow(13).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ssxshgl")){
                    if (!datum.get("ssxshgl").equals("") && datum.get("ssxshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("ssxshgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqpzdzds")){
                    if (!datum.get("lqpzdzds").equals("") && datum.get("lqpzdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("lqpzdzds").toString()));
                    } else {
                        sheet.getRow(15).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqpzdhgds")){
                    if (!datum.get("lqpzdhgds").equals("") && datum.get("lqpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("lqpzdhgds").toString()));
                    }else {
                        sheet.getRow(16).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqpzdhgl")){
                    if (!datum.get("lqpzdhgl").equals("") && datum.get("lqpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("lqpzdhgl").toString()));
                    }else {
                        sheet.getRow(17).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }

                if (datum.containsKey("mcxszds")){
                    if (!datum.get("mcxszds").equals("") && datum.get("mcxszds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("mcxszds").toString()));
                    }else {
                        sheet.getRow(18).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }

                if (datum.containsKey("mcxshgds")){
                    if (!datum.get("mcxshgds").equals("") && datum.get("mcxshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("mcxshgds").toString()));
                    }else {
                        sheet.getRow(19).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }

                if (datum.containsKey("mcxshgl")){
                    if (!datum.get("mcxshgl").equals("") && datum.get("mcxshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("mcxshgl").toString()));
                    }else {
                        sheet.getRow(20).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gzsdzds")){
                    if (!datum.get("gzsdzds").equals("") && datum.get("gzsdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("gzsdzds").toString()));
                    } else {
                        sheet.getRow(21).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gzsdhgds")){
                    if (!datum.get("gzsdhgds").equals("") && datum.get("gzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgds").toString()));
                    }else {
                        sheet.getRow(22).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gzsdhgl")){
                    if (!datum.get("gzsdhgl").equals("") && datum.get("gzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgl").toString()));
                    }else {
                        sheet.getRow(23).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqhdzds")){
                    if (!datum.get("lqhdzds").equals("") && datum.get("lqhdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("lqhdzds").toString()));
                    }else {
                        sheet.getRow(24).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqhdhgds")){
                    if (!datum.get("lqhdhgds").equals("") && datum.get("lqhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("lqhdhgds").toString()));
                    }else {
                        sheet.getRow(25).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqhdhgl")){
                    if (!datum.get("lqhdhgl").equals("") && datum.get("lqhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("lqhdhgl").toString()));
                    }else {
                        sheet.getRow(26).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqhpzds")){
                    if (!datum.get("lqhpzds").equals("") && datum.get("lqhpzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("lqhpzds").toString()));
                    }else {
                        sheet.getRow(27).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqhphgds")){
                    if (!datum.get("lqhphgds").equals("") && datum.get("lqhphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("lqhphgds").toString()));
                    }else {
                        sheet.getRow(28).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lqhphgl")){
                    if (!datum.get("lqhphgl").equals("") && datum.get("lqhphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("lqhphgl").toString()));
                    }else {
                        sheet.getRow(29).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntqdzds")){
                    if (!datum.get("hntqdzds").equals("") && datum.get("hntqdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("hntqdzds").toString()));
                    }else {
                        sheet.getRow(30).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntqdhgds")){
                    if (!datum.get("hntqdhgds").equals("") && datum.get("hntqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgds").toString()));
                    }else {
                        sheet.getRow(31).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntqdhgl")){
                    if (!datum.get("hntqdhgl").equals("") && datum.get("hntqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgl").toString()));
                    }else {
                        sheet.getRow(32).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }

                if (datum.containsKey("tlmxlbgczds")){
                    if (!datum.get("tlmxlbgczds").equals("") && datum.get("tlmxlbgczds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(33).createCell(index).setCellValue(Double.valueOf(datum.get("tlmxlbgczds").toString()));
                    }else {
                        sheet.getRow(33).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(33).createCell(index).setCellValue("");
                }

                if (datum.containsKey("tlmxlbgchgds")){
                    if (!datum.get("tlmxlbgchgds").equals("") && datum.get("tlmxlbgchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(34).createCell(index).setCellValue(Double.valueOf(datum.get("tlmxlbgchgds").toString()));
                    }else {
                        sheet.getRow(34).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(34).createCell(index).setCellValue("");
                }
                if (datum.containsKey("tlmxlbgchgl")){
                    if (!datum.get("tlmxlbgchgl").equals("") && datum.get("tlmxlbgchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(35).createCell(index).setCellValue(Double.valueOf(datum.get("tlmxlbgchgl").toString()));
                    }else {
                        sheet.getRow(35).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(35).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntpzdzds")){
                    if (!datum.get("hntpzdzds").equals("") && datum.get("hntpzdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(36).createCell(index).setCellValue(Double.valueOf(datum.get("hntpzdzds").toString()));
                    }else {
                        sheet.getRow(36).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(36).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntpzdhgds")){
                    if (!datum.get("hntpzdhgds").equals("") && datum.get("hntpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(37).createCell(index).setCellValue(Double.valueOf(datum.get("hntpzdhgds").toString()));
                    }else {
                        sheet.getRow(37).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(37).createCell(index).setCellValue("");
                }
                if (datum.containsKey("hntpzdhgl")){
                    if (!datum.get("hntpzdhgl").equals("") && datum.get("hntpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(38).createCell(index).setCellValue(Double.valueOf(datum.get("hntpzdhgl").toString()));
                    }else {
                        sheet.getRow(38).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(38).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khzds")){
                    if (!datum.get("khzds").equals("") && datum.get("khzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(39).createCell(index).setCellValue(Double.valueOf(datum.get("khzds").toString()));
                    }else {
                        sheet.getRow(39).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(39).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhgds")){
                    if (!datum.get("khhgds").equals("") && datum.get("khhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(40).createCell(index).setCellValue(Double.valueOf(datum.get("khhgds").toString()));
                    }else {
                        sheet.getRow(40).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(40).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhgl")){
                    if (!datum.get("khhgl").equals("") && datum.get("khhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(41).createCell(index).setCellValue(Double.valueOf(datum.get("khhgl").toString()));
                    }else {
                        sheet.getRow(41).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(41).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthdzds") && datum.get("hnthdzds") !=null){
                    if (!datum.get("hnthdzds").equals("") && datum.get("hnthdzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(42).createCell(index).setCellValue(Double.valueOf(datum.get("hnthdzds").toString()));
                    }else {
                        sheet.getRow(42).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(42).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthgds")&& datum.get("hnthgds") !=null){
                    if (!datum.get("hnthgds").equals("") && datum.get("hnthgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(43).createCell(index).setCellValue(Double.valueOf(datum.get("hnthgds").toString()));
                    }else {
                        sheet.getRow(43).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(43).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthgl")&& datum.get("hnthgl") !=null){
                    if (!datum.get("hnthgl").equals("") && datum.get("hnthgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(44).createCell(index).setCellValue(Double.valueOf(datum.get("hnthgl").toString()));
                    }else {
                        sheet.getRow(44).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(44).createCell(index).setCellValue("");
                }


                if (datum.containsKey("hnthpzds")){
                    if (!datum.get("hnthpzds").equals("") && datum.get("hnthpzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(45).createCell(index).setCellValue(Double.valueOf(datum.get("hnthpzds").toString()));
                    }else {
                        sheet.getRow(45).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(45).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthphgds")){
                    if (!datum.get("hnthphgds").equals("") && datum.get("hnthphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(46).createCell(index).setCellValue(Double.valueOf(datum.get("hnthphgds").toString()));
                    }else {
                        sheet.getRow(46).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(46).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hnthphdl")){
                    if (!datum.get("hnthphdl").equals("") && datum.get("hnthphdl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(47).createCell(index).setCellValue(Double.valueOf(datum.get("hnthphdl").toString()));
                    }else {
                        sheet.getRow(47).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(47).createCell(index).setCellValue("");
                }
                index++;
            }
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdlmhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        System.out.println(data);
        XSSFSheet sheet = wb.getSheet("表4.1.2-23");
        int index = 4;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("ysdzs")){
                    if (!datum.get("ysdzs").equals("") && datum.get("ysdzs").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("ysdzs").toString()));
                    }else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgs")){
                    if (!datum.get("ysdhgs").equals("") && datum.get("ysdhgs").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgs").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgl")){
                    if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdczjcds")){
                    if (!datum.get("sdczjcds").equals("") && datum.get("sdczjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("sdczjcds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdczhgds")){
                    if (!datum.get("sdczhgds").equals("") && datum.get("sdczhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("sdczhgds").toString()));
                    }else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdczhgl")){
                    if (!datum.get("sdczhgl").equals("") && datum.get("sdczhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("sdczhgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdssxsjcds")){
                    if (!datum.get("sdssxsjcds").equals("") && datum.get("sdssxsjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("sdssxsjcds").toString()));
                    }else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdssxshgds")){
                    if (!datum.get("sdssxshgds").equals("") && datum.get("sdssxshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("sdssxshgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdssxshgl")){
                    if (!datum.get("sdssxshgl").equals("") && datum.get("sdssxshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("sdssxshgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdpzdlqjcds")){
                    if (!datum.get("sdpzdlqjcds").equals("") && datum.get("sdpzdlqjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("sdpzdlqjcds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdpzdlqhgds")){
                    if (!datum.get("sdpzdlqhgds").equals("") && datum.get("sdpzdlqhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("sdpzdlqhgds").toString()));
                    }else {
                        sheet.getRow(13).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdpzdlqhgl")){
                    if (!datum.get("sdpzdlqhgl").equals("") && datum.get("sdpzdlqhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("sdpzdlqhgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqmcxsjcds")){
                    if (!datum.get("khlqmcxsjcds").equals("") && datum.get("khlqmcxsjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxsjcds").toString()));
                    }else {
                        sheet.getRow(15).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqmcxshgds")){
                    if (!datum.get("khlqmcxshgds").equals("") && datum.get("khlqmcxshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgds").toString()));
                    }else {
                        sheet.getRow(16).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqmcxshgl")){
                    if (!datum.get("khlqmcxshgl").equals("") && datum.get("khlqmcxshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgl").toString()));
                    }else {
                        sheet.getRow(17).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqgzsdjcds")){
                    if (!datum.get("khlqgzsdjcds").equals("") && datum.get("khlqgzsdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdjcds").toString()));
                    }else {
                        sheet.getRow(18).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqgzsdhgds")){
                    if (!datum.get("khlqgzsdhgds").equals("") && datum.get("khlqgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgds").toString()));
                    }else {
                        sheet.getRow(19).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqgzsdhgl")){
                    if (!datum.get("khlqgzsdhgl").equals("") && datum.get("khlqgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgl").toString()));
                    }else {
                        sheet.getRow(20).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdlqzds")){
                    if (!datum.get("hdlqzds").equals("") && datum.get("hdlqzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("hdlqzds").toString()));
                    }else {
                        sheet.getRow(21).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdlqhgds")){
                    if (!datum.get("hdlqhgds").equals("") && datum.get("hdlqhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("hdlqhgds").toString()));
                    }else {
                        sheet.getRow(22).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdlqhgl")){
                    if (!datum.get("hdlqhgl").equals("") && datum.get("hdlqhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("hdlqhgl").toString()));
                    }else {
                        sheet.getRow(23).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hplqzds")){
                    if (!datum.get("hplqzds").equals("") && datum.get("hplqzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("hplqzds").toString()));
                    }else {
                        sheet.getRow(24).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hplqhgds")){
                    if (!datum.get("hplqhgds").equals("") && datum.get("hplqhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("hplqhgds").toString()));
                    }else {
                        sheet.getRow(25).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hplqhgl")){
                    if (!datum.get("hplqhgds").equals("") && datum.get("hplqhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("hplqhgl").toString()));
                    }else {
                        sheet.getRow(26).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdpzdhntjcds")){
                    if (!datum.get("sdpzdhntjcds").equals("") && datum.get("sdpzdhntjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(33).createCell(index).setCellValue(Double.valueOf(datum.get("sdpzdhntjcds").toString()));
                    }else {
                        sheet.getRow(33).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(33).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdpzdhnthgds")){
                    if (!datum.get("sdpzdhnthgds").equals("") && datum.get("sdpzdhnthgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(34).createCell(index).setCellValue(Double.valueOf(datum.get("sdpzdhnthgds").toString()));
                    }else {
                        sheet.getRow(34).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(34).createCell(index).setCellValue("");
                }

                if (datum.containsKey("sdpzdhnthgl")){
                    if (!datum.get("sdpzdhnthgl").equals("") && datum.get("sdpzdhnthgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(35).createCell(index).setCellValue(Double.valueOf(datum.get("sdpzdhnthgl").toString()));
                    }else {
                        sheet.getRow(35).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(35).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntmcxsjcds")){
                    if (!datum.get("khhntmcxsjcds").equals("") && datum.get("khhntmcxsjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(36).createCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxsjcds").toString()));
                    }else {
                        sheet.getRow(36).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(36).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntmcxshgds")){
                    if (!datum.get("khhntmcxshgds").equals("") && datum.get("khhntmcxshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(37).createCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgds").toString()));
                    }else {
                        sheet.getRow(37).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(37).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntmcxshgl")){
                    if (!datum.get("khhntmcxshgl").equals("") && datum.get("khhntmcxshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(38).createCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgl").toString()));
                    }else {
                        sheet.getRow(38).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(38).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntgzsdjcds")){
                    if (!datum.get("khhntgzsdjcds").equals("") && datum.get("khhntgzsdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(39).createCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdjcds").toString()));
                    }else {
                        sheet.getRow(39).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(39).createCell(index).setCellValue("");
                }
                if (datum.containsKey("khhntgzsdhgds")){
                    if (!datum.get("khhntgzsdhgds").equals("") && datum.get("khhntgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(40).createCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgds").toString()));
                    }else {
                        sheet.getRow(40).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(40).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntgzsdhgl")){
                    if (!datum.get("khhntgzsdhgl").equals("") && datum.get("khhntgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(41).createCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgl").toString()));
                    }else {
                        sheet.getRow(41).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(41).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdhntzds")){
                    if (!datum.get("hdhntzds").equals("") && datum.get("hdhntzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(42).createCell(index).setCellValue(Double.valueOf(datum.get("hdhntzds").toString()));
                    }else {
                        sheet.getRow(42).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(42).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdhnthgds")){
                    if (!datum.get("hdhnthgds").equals("") && datum.get("hdhnthgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(43).createCell(index).setCellValue(Double.valueOf(datum.get("hdhnthgds").toString()));
                    }else {
                        sheet.getRow(43).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(43).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdhnthgl")){
                    if (!datum.get("hdhnthgl").equals("") && datum.get("hdhnthgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(44).createCell(index).setCellValue(Double.valueOf(datum.get("hdhnthgl").toString()));
                    }else {
                        sheet.getRow(44).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(44).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hphntzds")){
                    if (!datum.get("hphntzds").equals("") && datum.get("hphntzds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(45).createCell(index).setCellValue(Double.valueOf(datum.get("hphntzds").toString()));
                    }else {
                        sheet.getRow(45).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(45).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hphnthgds")){
                    if (!datum.get("hphnthgds").equals("") && datum.get("hphnthgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(46).createCell(index).setCellValue(Double.valueOf(datum.get("hphnthgds").toString()));
                    }else {
                        sheet.getRow(46).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(46).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hphnthgl")){
                    if (!datum.get("hphnthgl").equals("") && datum.get("hphnthgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(47).createCell(index).setCellValue(Double.valueOf(datum.get("hphnthgl").toString()));
                    }else {
                        sheet.getRow(47).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(47).createCell(index).setCellValue("");
                }
                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhpData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-22");
        XSSFSheet sheet = wb.getSheet("表4.1.2-22");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdhplmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdhplmlx").toString());
            }
            if (datum.containsKey("sdhpzds")){
                if (!datum.get("sdhpzds").equals("") && datum.get("sdhpzds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("sdhpzds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("sdhphgds")){
                if (!datum.get("sdhphgds").equals("") && datum.get("sdhphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdhphgds").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("sdhphgl")){
                if (!datum.get("sdhphgl").equals("") && datum.get("sdhphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdhphgl").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-21(3)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-21(3)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdlmlx").toString());
            }

            if (datum.containsKey("sdjcds")){
                if (!datum.get("sdjcds").equals("") && datum.get("sdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("sdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }
            if (datum.containsKey("sdhgs")){
                if (!datum.get("sdhgs").equals("") && datum.get("sdhgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdhgs").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("sdhgl")){
                if (!datum.get("sdhgl").equals("") && datum.get("sdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdldhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-21(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-21(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdldhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdldhdlmlx").toString());
            }
            if (datum.containsKey("sdldhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdldhdlb").toString());
            }
            if (datum.containsKey("sdldhdbhfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdldhdbhfw").toString());
            }
            if (datum.containsKey("sdldhdsjz")){
                if (!datum.get("sdldhdsjz").equals("") && datum.get("sdldhdsjz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdldhdsjz").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("sdldhdjcds")){
                if (!datum.get("sdldhdjcds").equals("") && datum.get("sdldhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdldhdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("sdldhdhgds")){
                if (!datum.get("sdldhdhgds").equals("") && datum.get("sdldhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdldhdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("sdldhdhgl")){
                if (!datum.get("sdldhdhgl").equals("") && datum.get("sdldhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdldhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            index++;
        }
    }

    /**
     *表4.1.4-2-1 隧道钻芯法厚度检测结果汇总表
     * @param wb
     * @param data
     */
    private void DBExcelsdgcsdzxfhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-2-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdzxfhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdzxfhdlmlx").toString());
            }
            if (datum.containsKey("sdzxfhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdzxfhdlb").toString());

            }
            if (datum.containsKey("sdzxfhdpjzfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdzxfhdpjzfw").toString());
            }

            if (datum.containsKey("sdzxfhdzdbz")){
                if (!datum.get("sdzxfhdzdbz").equals("") && datum.get("sdzxfhdzdbz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdzxfhdzdbz").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdydbz")){
                if (!datum.get("sdzxfhdydbz").equals("") && datum.get("sdzxfhdydbz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdzxfhdydbz").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdsjz")){
                if (!datum.get("sdzxfhdsjz").equals("") && datum.get("sdzxfhdsjz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdzxfhdsjz").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdjcds")){
                if (!datum.get("sdzxfhdjcds").equals("") && datum.get("sdzxfhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdzxfhdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdhgs")){
                if (!datum.get("sdzxfhdhgs").equals("") && datum.get("sdzxfhdhgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("sdzxfhdhgs").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(8).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdhgl")){
                if (!datum.get("sdzxfhdhgl").equals("") && datum.get("sdzxfhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("sdzxfhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(9).setCellValue("");
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdzxfhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-21(1)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdzxfhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdzxfhdlmlx").toString());
            }
            if (datum.containsKey("sdzxfhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdzxfhdlb").toString());

            }
            if (datum.containsKey("sdzxfhdpjzfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdzxfhdpjzfw").toString());
            }
            if (datum.containsKey("sdzxfhdzdbz")){
                if (!datum.get("sdzxfhdzdbz").equals("") && datum.get("sdzxfhdzdbz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdzxfhdzdbz").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }


            if (datum.containsKey("sdzxfhdydbz")){
                if (!datum.get("sdzxfhdydbz").equals("") && datum.get("sdzxfhdydbz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdzxfhdydbz").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdsjz")){
                if (!datum.get("sdzxfhdsjz").equals("") && datum.get("sdzxfhdsjz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdzxfhdsjz").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdjcds")){
                if (!datum.get("sdzxfhdjcds").equals("") && datum.get("sdzxfhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdzxfhdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdhgs")){
                if (!datum.get("sdzxfhdhgs").equals("") && datum.get("sdzxfhdhgs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("sdzxfhdhgs").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(8).setCellValue("");
            }

            if (datum.containsKey("sdzxfhdhgl")){
                if (!datum.get("sdzxfhdhgl").equals("") && datum.get("sdzxfhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("sdzxfhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(9).setCellValue("");
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdkhData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-20");
        XSSFSheet sheet = wb.getSheet("表4.1.2-20");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdkhlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdkhlmlx").toString());
            }
            if (datum.containsKey("sdkhzb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdkhzb").toString());
            }
            if (datum.containsKey("sdkhsjz")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdkhsjz").toString());
            }
            if (datum.containsKey("sdkhbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("sdkhbhfw").toString());
            }

            if (datum.containsKey("sdkhjcds")){
                if (!datum.get("sdkhjcds").equals("") && datum.get("sdkhjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdkhjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("sdkhhgds")){
                if (!datum.get("sdkhhgds").equals("") && datum.get("sdkhhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdkhhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("sdkhhgl")){
                if (!datum.get("sdkhhgl").equals("") && datum.get("sdkhhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdkhhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdpzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-19");
        XSSFSheet sheet = wb.getSheet("表4.1.2-19");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdpzdzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdpzdzb").toString());
            }
            if (datum.containsKey("sdpzdlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdpzdlmlx").toString());
            }
            if (datum.containsKey("sdpzdgdz")){
                if (!datum.get("sdpzdgdz").equals("") && datum.get("sdpzdgdz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdpzdgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("sdpzdbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("sdpzdbhfw").toString());
            }
            if (datum.containsKey("sdpzdjcds")){
                if (!datum.get("sdpzdjcds").equals("") && datum.get("sdpzdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdpzdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("sdpzdhgds")){
                if (!datum.get("sdpzdhgds").equals("") && datum.get("sdpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdpzdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("sdpzdhgl")){
                if (!datum.get("sdpzdhgl").equals("") && datum.get("sdpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdpzdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhntxlbgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhntqdData(XSSFWorkbook wb, List<Map<String, Object>> data) {

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdssxsData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-16");
        XSSFSheet sheet = wb.getSheet("表4.1.2-16");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdssxsgdz")){
                if (!datum.get("sdssxsgdz").equals("") && datum.get("sdssxsgdz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("sdssxsgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("sdssxsmin")){
                if (!datum.get("sdssxsmin").equals("") && datum.get("sdssxsmin").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("sdssxsmin").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("sdssxsmax")){
                if (!datum.get("sdssxsmax").equals("") && datum.get("sdssxsmax").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdssxsmax").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("sdssxsjcds")){
                if (!datum.get("sdssxsjcds").equals("") && datum.get("sdssxsjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdssxsjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("sdssxshgds")){
                if (!datum.get("sdssxshgds").equals("") && datum.get("sdssxshgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdssxshgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("sdssxshgl")){
                if (!datum.get("sdssxshgl").equals("") && datum.get("sdssxshgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdssxshgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdczData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-15(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-15(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdczzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdczzb").toString());
            }
            if (datum.containsKey("sdczlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdczlmlx").toString());
            }
            if (datum.containsKey("sdczgdz")){
                if (!datum.get("sdczgdz").equals("") && datum.get("sdczgdz").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdczgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("sdczbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("sdczbhfw").toString());
            }

            if (datum.containsKey("sdczjcds")){
                if (!datum.get("sdczjcds").equals("") && datum.get("sdczjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdczjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("sdczhgds")){
                if (!datum.get("sdczhgds").equals("") && datum.get("sdczhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdczhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("sdczhgl")){
                if (!datum.get("sdczhgl").equals("") && datum.get("sdczhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdczhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdysdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-15");
        Map<String, List<Map<String,Object>>> groupedData = data.stream()
                .filter(m -> m.get("bz") != null && !m.get("bz").toString().isEmpty())
                .collect(Collectors.groupingBy(m -> m.get("bz").toString()));
        groupedData.forEach((group, grouphtdData) -> {
            if (group.equals("试验室标准密度")){
                XSSFSheet sheet = wb.getSheet("表4.1.2-15");
                int index = 3;
                for (Map<String, Object> datum : grouphtdData) {
                    sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                    if (datum.containsKey("ysdlx")){
                        sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
                    }
                    if (datum.containsKey("ysdsczbh")){
                        sheet.getRow(index).getCell(2).setCellValue(datum.get("ysdsczbh").toString());
                    }

                    if (datum.containsKey("ysdzdbz")){
                        if (!datum.get("ysdzdbz").equals("") && datum.get("ysdzdbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdzdbz").toString()));
                        }else {
                            sheet.getRow(index).getCell(3).setCellValue("");
                        }

                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }

                    if (datum.containsKey("ysdydbz")){
                        if (!datum.get("ysdydbz").equals("") && datum.get("ysdydbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("ysdydbz").toString()));
                        }else {
                            sheet.getRow(index).getCell(4).setCellValue("");
                        }

                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                    if (datum.containsKey("ysdgdz")){
                        if (!datum.get("ysdgdz").equals("") && datum.get("ysdgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ysdgdz").toString()));
                        }else {
                            sheet.getRow(index).getCell(5).setCellValue("");
                        }

                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }

                    if (datum.containsKey("ysdzs")){
                        if (!datum.get("ysdzs").equals("") && datum.get("ysdzs").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("ysdzs").toString()));
                        }else {
                            sheet.getRow(index).getCell(6).setCellValue("");
                        }

                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }

                    if (datum.containsKey("ysdhgs")){
                        if (!datum.get("ysdhgs").equals("") && datum.get("ysdhgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("ysdhgs").toString()));
                        }else {
                            sheet.getRow(index).getCell(7).setCellValue("");
                        }

                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }

                    if (datum.containsKey("ysdhgl")){
                        if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                        }else {
                            sheet.getRow(index).getCell(8).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }
                    index ++;
                }

            }else if (group.equals("最大理论密度")){
                XSSFSheet sheet = wb.getSheet("表4.1.2-15-2");
                int index = 3;
                for (Map<String, Object> datum : grouphtdData) {
                    sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                    if (datum.containsKey("ysdlx")){
                        sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
                    }
                    if (datum.containsKey("ysdsczbh")){
                        sheet.getRow(index).getCell(2).setCellValue(datum.get("ysdsczbh").toString());
                    }

                    if (datum.containsKey("ysdzdbz")){
                        if (!datum.get("ysdzdbz").equals("") && datum.get("ysdzdbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdzdbz").toString()));
                        }else {
                            sheet.getRow(index).getCell(3).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }

                    if (datum.containsKey("ysdydbz")){
                        if (!datum.get("ysdydbz").equals("") && datum.get("ysdydbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("ysdydbz").toString()));
                        }else {
                            sheet.getRow(index).getCell(4).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                    if (datum.containsKey("ysdgdz")){
                        if (!datum.get("ysdgdz").equals("") && datum.get("ysdgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ysdgdz").toString()));
                        }else {
                            sheet.getRow(index).getCell(5).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }

                    if (datum.containsKey("ysdzs")){
                        if (!datum.get("ysdzs").equals("") && datum.get("ysdzs").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("ysdzs").toString()));
                        }else {
                            sheet.getRow(index).getCell(6).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }

                    if (datum.containsKey("ysdhgs")){
                        if (!datum.get("ysdhgs").equals("") && datum.get("ysdhgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("ysdhgs").toString()));
                        }else {
                            sheet.getRow(index).getCell(7).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }

                    if (datum.containsKey("ysdhgl")){
                        if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                            sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                        }else {
                            sheet.getRow(index).getCell(8).setCellValue("");
                        }
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }
                    index ++;
                }
            }

        });

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmxhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        System.out.println(data);
        XSSFSheet sheet = wb.getSheet("表4.1.2-14");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("qmpzdjcds")){
                    if (!datum.get("qmpzdjcds").equals("") && datum.get("qmpzdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("qmpzdjcds").toString()));
                    }else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmpzdhgds")){
                    if (!datum.get("qmpzdhgds").equals("") && datum.get("qmpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("qmpzdhgds").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmpzdhgl")){
                    if (!datum.get("qmpzdhgl").equals("") && datum.get("qmpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("qmpzdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }
                if (datum.containsKey("qmhpzds")){
                    if (!datum.get("qmhpzds").equals("") && datum.get("qmhpzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmhphgds")){
                    if (!datum.get("qmhphgds").equals("") && datum.get("qmhphgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
                    }else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

                if (datum.containsKey("qmhphgl")){
                    if (!datum.get("qmhphgl").equals("") && datum.get("qmhphgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("mcxszds")){
                    if (!datum.get("mcxszds").equals("") && datum.get("mcxszds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("mcxszds").toString()));
                    }else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("mcxshgds")){
                    if (!datum.get("mcxshgds").equals("") && datum.get("mcxshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("mcxshgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("mcxshgl")){
                    if (!datum.get("mcxshgl").equals("") && datum.get("mcxshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("mcxshgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("gzsdzds")){
                    if (!datum.get("gzsdzds").equals("") && datum.get("gzsdzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("gzsdzds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }
                if (datum.containsKey("gzsdhgds")){
                    if (!datum.get("gzsdhgds").equals("") && datum.get("gzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgds").toString()));
                    }else {
                        sheet.getRow(13).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }
                if (datum.containsKey("gzsdhgl")){
                    if (!datum.get("gzsdhgl").equals("") && datum.get("gzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }
                index++;
            }

        }
    }


    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmkhData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-13");
        int index = 3;

        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmkhlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmkhlmlx").toString());
            }
            if (datum.containsKey("qmkhzb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("qmkhzb").toString());
            }
            if (datum.containsKey("qmkhsjz")){
                if (!datum.get("qmkhsjz").equals("") && datum.get("qmkhsjz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmkhsjz").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("qmkhbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("qmkhbhfw").toString());
            }

            if (datum.containsKey("qmkhjcds")){
                if (!datum.get("qmkhjcds").equals("") && datum.get("qmkhjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qmkhjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("qmkhhgds")){
                if (!datum.get("qmkhhgds").equals("") && datum.get("qmkhhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qmkhhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("qmkhhgl")){
                if (!datum.get("qmkhhgl").equals("") && datum.get("qmkhhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qmkhhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmhpdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-12");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmhplmlx")) {
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmhplmlx").toString());
            }
            if (datum.containsKey("qmhpzds")){
                if (!datum.get("qmhpzds").equals("") && datum.get("qmhpzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("qmhphgds")){
                if (!datum.get("qmhphgds").equals("") && datum.get("qmhphgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("qmhphgl")){
                if (!datum.get("qmhphgl").equals("") && datum.get("qmhphgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmpzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-11");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmpzdzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmpzdzb").toString());
            }
            if (datum.containsKey("qmpzdlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("qmpzdlmlx").toString());
            }
            if (datum.containsKey("qmpzdgdz")){
                if (!datum.get("qmpzdgdz").equals("") && datum.get("qmpzdgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmpzdgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("qmpzdbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("qmpzdbhfw").toString());
            }

            if (datum.containsKey("qmpzdjcds")){
                if (!datum.get("qmpzdjcds").equals("") && datum.get("qmpzdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qmpzdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("qmpzdhgds")){
                if (!datum.get("qmpzdhgds").equals("") && datum.get("qmpzdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qmpzdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("qmpzdhgl")){
                if (!datum.get("qmpzdhgl").equals("") && datum.get("qmpzdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qmpzdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-10");
        int index = 4;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());
                //sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("ysdzds")){
                    if (!datum.get("ysdzds").equals("") && datum.get("ysdzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("ysdzds").toString()));
                    }else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgds")){
                    if (!datum.get("ysdhgds").equals("") && datum.get("ysdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgl")){
                    if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmwczds")){
                    if (!datum.get("lmwczds").equals("") && datum.get("lmwczds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("lmwczds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmwchgds")){
                    if (!datum.get("lmwchgds").equals("") && datum.get("lmwchgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("lmwchgds").toString()));
                    }else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmwchgl")){
                    if (!datum.get("lmwchgl").equals("") && datum.get("lmwchgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("lmwchgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czzds")){
                    if (!datum.get("czzds").equals("") && datum.get("czzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("czzds").toString()));
                    }else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czhgds")){
                    if (!datum.get("czhgds").equals("") && datum.get("czhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("czhgl")){
                    if (!datum.get("czhgl").equals("") && datum.get("czhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmssxsssjcds")){
                    if (!datum.get("lmssxsssjcds").equals("") && datum.get("lmssxsssjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("lmssxsssjcds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmssxssshgds")){
                    if (!datum.get("lmssxssshgds").equals("") && datum.get("lmssxssshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("lmssxssshgds").toString()));
                    }else {
                        sheet.getRow(13).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmssxssshgl")){
                    if (!datum.get("lmssxssshgl").equals("") && datum.get("lmssxssshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("lmssxssshgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdlqjcds")){
                    if (!datum.get("pzdlqjcds").equals("") && datum.get("pzdlqjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("pzdlqjcds").toString()));
                    } else {
                        sheet.getRow(15).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdlqhgds")){
                    if (!datum.get("pzdlqhgds").equals("") && datum.get("pzdlqhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("pzdlqhgds").toString()));
                    }else {
                        sheet.getRow(16).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdlqhgl")){
                    if (!datum.get("pzdlqhgl").equals("") && datum.get("pzdlqhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("pzdlqhgl").toString()));
                    }else {
                        sheet.getRow(17).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqmcxsjcds")){
                    if (!datum.get("khlqmcxsjcds").equals("") && datum.get("khlqmcxsjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxsjcds").toString()));
                    }else {
                        sheet.getRow(18).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqmcxshgds")){
                    if (!datum.get("khlqmcxshgds").equals("") && datum.get("khlqmcxshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgds").toString()));
                    }else {
                        sheet.getRow(19).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqmcxshgl")){
                    if (!datum.get("khlqmcxshgl").equals("") && datum.get("khlqmcxshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgl").toString()));
                    }else {
                        sheet.getRow(20).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqgzsdjcds")){
                    if (!datum.get("khlqgzsdjcds").equals("") && datum.get("khlqgzsdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdjcds").toString()));
                    }else {
                        sheet.getRow(21).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqgzsdhgds")){
                    if (!datum.get("khlqgzsdhgds").equals("") && datum.get("khlqgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgds").toString()));
                    }else {
                        sheet.getRow(22).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khlqgzsdhgl")){
                    if (!datum.get("khlqgzsdhgl").equals("") && datum.get("khlqgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgl").toString()));
                    }else {
                        sheet.getRow(23).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmhdlqjcds")){
                    if (!datum.get("lmhdlqjcds").equals("") && datum.get("lmhdlqjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("lmhdlqjcds").toString()));
                    }else {
                        sheet.getRow(24).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmhdlqhgs")){
                    if (!datum.get("lmhdlqhgs").equals("") && datum.get("lmhdlqhgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("lmhdlqhgs").toString()));
                    }else {
                        sheet.getRow(25).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmhdlqhgl")){
                    if (!datum.get("lmhdlqhgl").equals("") && datum.get("lmhdlqhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("lmhdlqhgl").toString()));
                    }else {
                        sheet.getRow(26).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hplqzds")){
                    if (!datum.get("hplqzds").equals("") && datum.get("hplqzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("hplqzds").toString()));
                    }else {
                        sheet.getRow(27).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hplqhgds")){
                    if (!datum.get("hplqhgds").equals("") && datum.get("hplqhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("hplqhgds").toString()));
                    }else {
                        sheet.getRow(28).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }
                if (datum.containsKey("hplqhgl")){
                    if (!datum.get("hplqhgl").equals("") && datum.get("hplqhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("hplqhgl").toString()));
                    } else {
                        sheet.getRow(29).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntqdzds")){
                    if (!datum.get("hntqdzds").equals("") && datum.get("hntqdzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("hntqdzds").toString()));
                    }else {
                        sheet.getRow(30).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntqdhgds")){
                    if (!datum.get("hntqdhgds").equals("") && datum.get("hntqdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgds").toString()));
                    }else {
                        sheet.getRow(31).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntqdhgl")){
                    if (!datum.get("hntqdhgl").equals("") && datum.get("hntqdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgl").toString()));
                    }else {
                        sheet.getRow(32).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntxlbgczds")){
                    if (!datum.get("hntxlbgczds").equals("") && datum.get("hntxlbgczds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(33).createCell(index).setCellValue(Double.valueOf(datum.get("hntxlbgczds").toString()));
                    }else {
                        sheet.getRow(33).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(33).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntxlbgchgds")){
                    if (!datum.get("hntxlbgchgds").equals("") && datum.get("hntxlbgchgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(34).createCell(index).setCellValue(Double.valueOf(datum.get("hntxlbgchgds").toString()));
                    }else {
                        sheet.getRow(34).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(34).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hntxlbgchgl")){
                    if (!datum.get("hntxlbgchgl").equals("") && datum.get("hntxlbgchgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(35).createCell(index).setCellValue(Double.valueOf(datum.get("hntxlbgchgl").toString()));
                    }else {
                        sheet.getRow(35).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(35).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdhntjcds")){
                    if (!datum.get("pzdhntjcds").equals("") && datum.get("pzdhntjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(36).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhntjcds").toString()));
                    }else {
                        sheet.getRow(36).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(36).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdhnthgds")){
                    if (!datum.get("pzdhnthgds").equals("") && datum.get("pzdhnthgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(37).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhnthgds").toString()));
                    }else {
                        sheet.getRow(37).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(37).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pzdhnthgl")){
                    if (!datum.get("pzdhnthgl").equals("") && datum.get("pzdhnthgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(38).createCell(index).setCellValue(Double.valueOf(datum.get("pzdhnthgl").toString()));
                    }else {
                        sheet.getRow(38).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(38).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntmcxsjcds")){
                    if (!datum.get("khhntmcxsjcds").equals("") && datum.get("khhntmcxsjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(39).createCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxsjcds").toString()));
                    }else {
                        sheet.getRow(39).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(39).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntmcxshgds")){
                    if (!datum.get("khhntmcxshgds").equals("") && datum.get("khhntmcxshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(40).createCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgds").toString()));
                    }else {
                        sheet.getRow(40).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(40).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntmcxshgl")){
                    if (!datum.get("khhntmcxshgl").equals("") && datum.get("khhntmcxshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(41).createCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgl").toString()));
                    }else {
                        sheet.getRow(41).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(41).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntgzsdjcds")){
                    if (!datum.get("khhntgzsdjcds").equals("") && datum.get("khhntgzsdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(42).createCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdjcds").toString()));
                    }else {
                        sheet.getRow(42).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(42).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntgzsdhgds")){
                    if (!datum.get("khhntgzsdhgds").equals("") && datum.get("khhntgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(43).createCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgds").toString()));
                    }else {
                        sheet.getRow(43).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(43).createCell(index).setCellValue("");
                }

                if (datum.containsKey("khhntgzsdhgl")){
                    if (!datum.get("khhntgzsdhgl").equals("") && datum.get("khhntgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(44).createCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgl").toString()));
                    }else {
                        sheet.getRow(44).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(44).createCell(index).setCellValue("");
                }
                if (datum.containsKey("lmhdhntjcds")){
                    if (!datum.get("lmhdhntjcds").equals("") && datum.get("lmhdhntjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(45).createCell(index).setCellValue(Double.valueOf(datum.get("lmhdhntjcds").toString()));
                    }else {
                        sheet.getRow(45).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(45).createCell(index).setCellValue("     ");
                }

                if (datum.containsKey("lmhdhnthgs")){
                    if (!datum.get("lmhdhnthgs").equals("") && datum.get("lmhdhnthgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(46).createCell(index).setCellValue(Double.valueOf(datum.get("lmhdhnthgs").toString()));
                    }else {
                        sheet.getRow(46).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(46).createCell(index).setCellValue("");
                }

                if (datum.containsKey("lmhdhnthgl")){
                    if (!datum.get("lmhdhnthgl").equals("") && datum.get("lmhdhnthgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(47).createCell(index).setCellValue(Double.valueOf(datum.get("lmhdhnthgl").toString()));
                    }else {
                        sheet.getRow(47).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(47).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hphntzds")){
                    if (!datum.get("hphntzds").equals("") && datum.get("hphntzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(48).createCell(index).setCellValue(Double.valueOf(datum.get("hphntzds").toString()));
                    }else {
                        sheet.getRow(48).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(48).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hphnthgds")){
                    if (!datum.get("hphnthgds").equals("") && datum.get("hphnthgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(49).createCell(index).setCellValue(Double.valueOf(datum.get("hphnthgds").toString()));
                    }else {
                        sheet.getRow(49).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(49).createCell(index).setCellValue("");
                }
                if (datum.containsKey("hphnthgl")){
                    if (!datum.get("hphnthgl").equals("") && datum.get("hphnthgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(50).createCell(index).setCellValue(Double.valueOf(datum.get("hphnthgl").toString()));
                    }else {
                        sheet.getRow(50).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(50).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxysdzds")){
                    if (!datum.get("ljxysdzds").equals("") && datum.get("ljxysdzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(51).createCell(index).setCellValue(Double.valueOf(datum.get("ljxysdzds").toString()));
                    }else {
                        sheet.getRow(51).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(51).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxysdhgds")){
                    if (!datum.get("ljxysdhgds").equals("") && datum.get("ljxysdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(52).createCell(index).setCellValue(Double.valueOf(datum.get("ljxysdhgds").toString()));
                    }else {
                        sheet.getRow(52).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(52).createCell(index).setCellValue("");
                }
                if (datum.containsKey("ljxysdhgl")){
                    if (!datum.get("ljxysdhgl").equals("") && datum.get("ljxysdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(53).createCell(index).setCellValue(Double.valueOf(datum.get("ljxysdhgl").toString()));
                    }else {
                        sheet.getRow(53).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(53).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxczzds")){
                    if (!datum.get("ljxczzds").equals("") && datum.get("ljxczzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(54).createCell(index).setCellValue(Double.valueOf(datum.get("ljxczzds").toString()));
                    }else {
                        sheet.getRow(54).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(54).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxczhgds")){
                    if (!datum.get("ljxczhgds").equals("") && datum.get("ljxczhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(55).createCell(index).setCellValue(Double.valueOf(datum.get("ljxczhgds").toString()));
                    }else {
                        sheet.getRow(55).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(55).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxczhgl")){
                    if (!datum.get("ljxczhgl").equals("") && datum.get("ljxczhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(56).createCell(index).setCellValue(Double.valueOf(datum.get("ljxczhgl").toString()));
                    }else {
                        sheet.getRow(56).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(56).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxlmssxsssjcds")){
                    if (!datum.get("ljxlmssxsssjcds").equals("") && datum.get("ljxlmssxsssjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(57).createCell(index).setCellValue(Double.valueOf(datum.get("ljxlmssxsssjcds").toString()));
                    } else {
                        sheet.getRow(57).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(57).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxlmssxssshgds")){
                    if (!datum.get("ljxlmssxssshgds").equals("") && datum.get("ljxlmssxssshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(58).createCell(index).setCellValue(Double.valueOf(datum.get("ljxlmssxssshgds").toString()));
                    }else {
                        sheet.getRow(58).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(58).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxlmssxssshgl")){
                    if (!datum.get("ljxlmssxssshgl").equals("") && datum.get("ljxlmssxssshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(59).createCell(index).setCellValue(Double.valueOf(datum.get("ljxlmssxssshgl").toString()));
                    }else {
                        sheet.getRow(59).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(59).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxpzdlqjcds")){
                    if (!datum.get("ljxpzdlqjcds").equals("") && datum.get("ljxpzdlqjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(60).createCell(index).setCellValue(Double.valueOf(datum.get("ljxpzdlqjcds").toString()));
                    }else {
                        sheet.getRow(60).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(60).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxpzdlqhgds")){
                    if (!datum.get("ljxpzdlqhgds").equals("") && datum.get("ljxpzdlqhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(61).createCell(index).setCellValue(Double.valueOf(datum.get("ljxpzdlqhgds").toString()));
                    }else {
                        sheet.getRow(61).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(61).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxpzdlqhgl")){
                    if (!datum.get("ljxpzdlqhgl").equals("") && datum.get("ljxpzdlqhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(62).createCell(index).setCellValue(Double.valueOf(datum.get("ljxpzdlqhgl").toString()));
                    }else {
                        sheet.getRow(62).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(62).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxkhlqmcxsjcds")){
                    if (!datum.get("ljxkhlqmcxsjcds").equals("") && datum.get("ljxkhlqmcxsjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(63).createCell(index).setCellValue(Double.valueOf(datum.get("ljxkhlqmcxsjcds").toString()));
                    }else {
                        sheet.getRow(63).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(63).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxkhlqmcxshgds")){
                    if (!datum.get("ljxkhlqmcxshgds").equals("") && datum.get("ljxkhlqmcxshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(64).createCell(index).setCellValue(Double.valueOf(datum.get("ljxkhlqmcxshgds").toString()));
                    }else {
                        sheet.getRow(64).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(64).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxkhlqmcxshgl")){
                    if (!datum.get("ljxkhlqmcxshgl").equals("") && datum.get("ljxkhlqmcxshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(65).createCell(index).setCellValue(Double.valueOf(datum.get("ljxkhlqmcxshgl").toString()));
                    }else {
                        sheet.getRow(65).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(65).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxkhlqgzsdjcds")){
                    if (!datum.get("ljxkhlqgzsdjcds").equals("") && datum.get("ljxkhlqgzsdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(66).createCell(index).setCellValue(Double.valueOf(datum.get("ljxkhlqgzsdjcds").toString()));
                    }else {
                        sheet.getRow(66).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(66).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxkhlqgzsdhgds")){
                    if (!datum.get("ljxkhlqgzsdhgds").equals("") && datum.get("ljxkhlqgzsdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(67).createCell(index).setCellValue(Double.valueOf(datum.get("ljxkhlqgzsdhgds").toString()));
                    }else {
                        sheet.getRow(67).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(67).createCell(index).setCellValue("");
                }
                if (datum.containsKey("ljxkhlqgzsdhgl")){
                    if (!datum.get("ljxkhlqgzsdhgl").equals("") && datum.get("ljxkhlqgzsdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(68).createCell(index).setCellValue(Double.valueOf(datum.get("ljxkhlqgzsdhgl").toString()));
                    }else {
                        sheet.getRow(68).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(68).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxlmhdlqjcds")){
                    if (!datum.get("ljxlmhdlqjcds").equals("") && datum.get("ljxlmhdlqjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(69).createCell(index).setCellValue(Double.valueOf(datum.get("ljxlmhdlqjcds").toString()));
                    }else {
                        sheet.getRow(69).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(69).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxlmhdlqhgs")){
                    if (!datum.get("ljxlmhdlqhgs").equals("") && datum.get("ljxlmhdlqhgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(70).createCell(index).setCellValue(Double.valueOf(datum.get("ljxlmhdlqhgs").toString()));
                    }else {
                        sheet.getRow(70).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(70).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxlmhdlqhgl")){
                    if (!datum.get("ljxlmhdlqhgl").equals("") && datum.get("ljxlmhdlqhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(71).createCell(index).setCellValue(Double.valueOf(datum.get("ljxlmhdlqhgl").toString()));
                    }else {
                        sheet.getRow(71).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(71).createCell(index).setCellValue("");
                }


                if (datum.containsKey("ljxhphntzds")){
                    if (!datum.get("ljxhphntzds").equals("") && datum.get("ljxhphntzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(72).createCell(index).setCellValue(Double.valueOf(datum.get("ljxhphntzds").toString()));
                    }else {
                        sheet.getRow(72).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(72).createCell(index).setCellValue("");
                }
                if (datum.containsKey("ljxhphnthgds")){
                    if (!datum.get("ljxhphnthgds").equals("") && datum.get("ljxhphnthgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(73).createCell(index).setCellValue(Double.valueOf(datum.get("ljxhphnthgds").toString()));
                    }else {
                        sheet.getRow(73).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(73).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ljxhphnthgl")){
                    if (!datum.get("ljxhphnthgl").equals("") && datum.get("ljxhphnthgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                        sheet.getRow(74).createCell(index).setCellValue(Double.valueOf(datum.get("ljxhphnthgl").toString()));
                    }else {
                        sheet.getRow(74).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(74).createCell(index).setCellValue("");
                }


                index++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhpData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-9");
        XSSFSheet sheet = wb.getSheet("表4.1.2-9");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hplmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("hplmlx").toString());
            }
            if (datum.containsKey("hpzds")){
                if (!datum.get("hpzds").equals("") && datum.get("hpzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hpzds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }
            if (datum.containsKey("hphgds")){
                if (!datum.get("hphgds").equals("") && datum.get("hphgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hphgds").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("hphgl")){
                if (!datum.get("hphgl").equals("") && datum.get("hphgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hphgl").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-8(3)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-8(3)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("lmhdlmlx").toString());
            }

            if (datum.containsKey("lmhdjcds")){
                if (!datum.get("lmhdjcds").equals("") && datum.get("lmhdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmhdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }
            if (datum.containsKey("lmhdhgs")){
                if (!datum.get("lmhdhgs").equals("") && datum.get("lmhdhgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmhdhgs").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("lmhdhgl")){
                if (!datum.get("lmhdhgl").equals("") && datum.get("lmhdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelldhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-8(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-8(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ldhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("ldhdlmlx").toString());
            }
            if (datum.containsKey("ldhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("ldhdlb").toString());
            }
            if (datum.containsKey("ldhdbhfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("ldhdbhfw").toString());
            }
            if (datum.containsKey("ldhdsjz")){
                if (!datum.get("ldhdsjz").equals("") && datum.get("ldhdsjz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("ldhdsjz").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("ldhdjcds")){
                if (!datum.get("ldhdjcds").equals("") && datum.get("ldhdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ldhdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("ldhdhgds")){
                if (!datum.get("ldhdhgds").equals("") && datum.get("ldhdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("ldhdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("ldhdhgl")){
                if (!datum.get("ldhdhgl").equals("") && datum.get("ldhdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("ldhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelzxfhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-8(1)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-8(1)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("zxfhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("zxfhdlmlx").toString());
            }
            if (datum.containsKey("zxfhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("zxfhdlb").toString());

            }
            if (datum.containsKey("zxfhdpjzfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("zxfhdpjzfw").toString());
            }
            if (datum.containsKey("zxfhdzdbz")){
                if (!datum.get("zxfhdzdbz").equals("") && datum.get("zxfhdzdbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zxfhdzdbz").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("zxfhdydbz")){
                if (!datum.get("zxfhdydbz").equals("") && datum.get("zxfhdydbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("zxfhdydbz").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("zxfhdsjz")){
                if (!datum.get("zxfhdsjz").equals("") && datum.get("zxfhdsjz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("zxfhdsjz").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            if (datum.containsKey("zxfhdjcds")){
                if (!datum.get("zxfhdjcds").equals("") && datum.get("zxfhdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zxfhdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }

            if (datum.containsKey("zxfhdhgs")){
                if (!datum.get("zxfhdhgs").equals("") && datum.get("zxfhdhgs").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zxfhdhgs").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(8).setCellValue("");
            }

            if (datum.containsKey("zxfhdhgl")){
                if (!datum.get("zxfhdhgl").equals("") && datum.get("zxfhdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("zxfhdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(9).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelkhData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-7");
        XSSFSheet sheet = wb.getSheet("表4.1.2-7");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("khlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("khlmlx").toString());
            }
            if (datum.containsKey("khkpzb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("khkpzb").toString());
            }
            if (datum.containsKey("khsjz")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("khsjz").toString());
            }
            if (datum.containsKey("khbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("khbhfw").toString());
            }

            if (datum.containsKey("khjcds")){
                if (!datum.get("khjcds").equals("") && datum.get("khjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("khjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");

            }
            if (datum.containsKey("khhgds")){
                if (!datum.get("khhgds").equals("") && datum.get("khhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("khhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            if (datum.containsKey("khhgl")){
                if (!datum.get("khhgl").equals("") && datum.get("khhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("khhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmpzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-6");
        XSSFSheet sheet = wb.getSheet("表4.1.2-6");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("pzdzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("pzdzb").toString());
            }
            if (datum.containsKey("pzdlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("pzdlmlx").toString());
            }
            if (datum.containsKey("pzdgdz")){
                if (!datum.get("pzdgdz").equals("") && datum.get("pzdgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("pzdgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }
            if (datum.containsKey("pzdbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("pzdbhfw").toString());
            }
            if (datum.containsKey("pzdjcds")){
                if (!datum.get("pzdjcds").equals("") && datum.get("pzdjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("pzdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("pzdhgds")){
                if (!datum.get("pzdhgds").equals("") && datum.get("pzdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("pzdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            if (datum.containsKey("pzdhgl")){
                if (!datum.get("pzdhgl").equals("") && datum.get("pzdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("pzdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhntxlbgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-5");
        XSSFSheet sheet = wb.getSheet("表4.1.2-5");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hntxlbgcgdz")){
                if (!datum.get("hntxlbgcgdz").equals("") && datum.get("hntxlbgcgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hntxlbgcgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }
            if (datum.containsKey("hntxlbgcmin")){
                if (!datum.get("hntxlbgcmin").equals("") && datum.get("hntxlbgcmin").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hntxlbgcmin").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }
            if (datum.containsKey("hntxlbgcmax")){
                if (!datum.get("hntxlbgcmax").equals("") && datum.get("hntxlbgcmax").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hntxlbgcmax").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }
            if (datum.containsKey("hntxlbgczds")){
                if (!datum.get("hntxlbgczds").equals("") && datum.get("hntxlbgczds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hntxlbgczds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            if (datum.containsKey("hntxlbgchgds")){
                if (!datum.get("hntxlbgchgds").equals("") && datum.get("hntxlbgchgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hntxlbgchgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("hntxlbgchgl")){
                if (!datum.get("hntxlbgchgl").equals("") && datum.get("hntxlbgchgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hntxlbgchgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhntqdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //(wb,data.size(),"表4.1.2-4");
        XSSFSheet sheet = wb.getSheet("表4.1.2-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hntqdgdz")){
                if (!datum.get("hntqdgdz").equals("") && datum.get("hntqdgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hntqdgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }
            if (datum.containsKey("hntqdmin")){
                if (!datum.get("hntqdmin").equals("") && datum.get("hntqdmin").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hntqdmin").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }
            if (datum.containsKey("hntqdmax")){
                if (!datum.get("hntqdmax").equals("") && datum.get("hntqdmax").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hntqdmax").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }
            if (datum.containsKey("hntqdpjz")){
                if (!datum.get("hntqdpjz").equals("") && datum.get("hntqdpjz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hntqdpjz").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            if (datum.containsKey("hntqdzds")){
                if (!datum.get("hntqdzds").equals("") && datum.get("hntqdzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hntqdzds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("hntqdhgds")){
                if (!datum.get("hntqdhgds").equals("") && datum.get("hntqdhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hntqdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            if (datum.containsKey("hntqdhgl")){
                if (!datum.get("hntqdhgl").equals("") && datum.get("hntqdhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("hntqdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmssxsData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-3");
        XSSFSheet sheet = wb.getSheet("表4.1.2-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            sheet.getRow(index).getCell(1).setCellValue(datum.get("lxmc").toString());
            if (datum.containsKey("lmssxsgdz")){
                if (!datum.get("lmssxsgdz").equals("") && datum.get("lmssxsgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmssxsgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("lmssxsmin")){
                if (!datum.get("lmssxsmin").equals("") && datum.get("lmssxsmin").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmssxsmin").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("lmssxsmax")){
                if (!datum.get("lmssxsmax").equals("") && datum.get("lmssxsmax").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmssxsmax").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("lmssxsssjcds")){
                if (!datum.get("lmssxsssjcds").equals("") && datum.get("lmssxsssjcds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmssxsssjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("lmssxssshgds")){
                if (!datum.get("lmssxssshgds").equals("") && datum.get("lmssxssshgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmssxssshgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            if (datum.containsKey("lmssxssshgl")){
                if (!datum.get("lmssxssshgl").equals("") && datum.get("lmssxssshgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("lmssxssshgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmczData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-2(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-2(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("czzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("czzb").toString());
            }
            if (datum.containsKey("czlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("czlmlx").toString());
            }
            if (datum.containsKey("czgdz")){
                if (!datum.get("czgdz").equals("") && datum.get("czgdz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("czgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }
            if (datum.containsKey("czbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("czbhfw").toString());
            }
            if (datum.containsKey("czzds")){
                if (!datum.get("czzds").equals("") && datum.get("czzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("czzds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("czhgds")){
                if (!datum.get("czhgds").equals("") && datum.get("czhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            if (datum.containsKey("czhgl")){
                if (!datum.get("czhgl").equals("") && datum.get("czhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue("");
            }
            index++;
        }


    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmwclcfdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-21");
        XSSFSheet sheet = wb.getSheet("表4.1.2-21");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmwclcfdbz")){
                if (!datum.get("lmwclcfdbz").equals("") && datum.get("lmwclcfdbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("lmwclcfdbz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }
            if (datum.containsKey("lmwclcfmax")){
                if (!datum.get("lmwclcfmax").equals("") && datum.get("lmwclcfmax").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmwclcfmax").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("lmwclcfmin")){
                if (!datum.get("lmwclcfmin").equals("") && datum.get("lmwclcfmin").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmwclcfmin").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("lmwclcfzds")){
                if (!datum.get("lmwclcfzds").equals("") && datum.get("lmwclcfzds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmwclcfzds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }
            if (datum.containsKey("lmwclcfhgds")){
                if (!datum.get("lmwclcfhgds").equals("") && datum.get("lmwclcfhgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmwclcfhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("lmwclcfhgl")){
                if (!datum.get("lmwclcfhgl").equals("") && datum.get("lmwclcfhgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmwclcfhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmwcalldData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmwcdbz")){
                if (!datum.get("lmwcdbz").equals("") && datum.get("lmwcdbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("lmwcdbz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("lmwcmax")){
                if (!datum.get("lmwcmax").equals("") && datum.get("lmwcmax").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmwcmax").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("lmwcmin")){
                if (!datum.get("lmwcmin").equals("") && datum.get("lmwcmin").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmwcmin").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("lmwczds")){
                if (!datum.get("lmwczds").equals("") && datum.get("lmwczds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmwczds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("lmwchgds")){
                if (!datum.get("lmwchgds").equals("") && datum.get("lmwchgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmwchgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("lmwchgl")){
                if (!datum.get("lmwchgl").equals("") && datum.get("lmwchgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmwchgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmwcdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-211");
        XSSFSheet sheet = wb.getSheet("表4.1.2-211");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmwcdbz")){
                if (!datum.get("lmwcdbz").equals("") && datum.get("lmwcdbz").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("lmwcdbz").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("lmwcmax")){
                if (!datum.get("lmwcmax").equals("") && datum.get("lmwcmax").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmwcmax").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("lmwcmin")){
                if (!datum.get("lmwcmin").equals("") && datum.get("lmwcmin").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmwcmin").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("lmwczds")){
                if (!datum.get("lmwczds").equals("") && datum.get("lmwczds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmwczds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("lmwchgds")){
                if (!datum.get("lmwchgds").equals("") && datum.get("lmwchgds").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmwchgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("lmwchgl")){
                if (!datum.get("lmwchgl").equals("") && datum.get("lmwchgl").toString().matches("-?\\d+(\\.\\d+)?")) {
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmwchgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmysdData(XSSFWorkbook wb, List<Map<String, Object>> data) {

        Map<String, List<Map<String,Object>>> groupedData = data.stream()
                .collect(Collectors.groupingBy(m -> m.get("bz").toString()));
        groupedData.forEach((group, grouphtdData) -> {
            if (group.equals("试验室标准密度")){
                XSSFSheet sheet = wb.getSheet("表4.1.2-1");
                int index = 3;
                for (Map<String, Object> datum : grouphtdData) {
                        sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                        sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
                        sheet.getRow(index).getCell(2).setCellValue(datum.get("bz").toString());

                        if (datum.containsKey("bzsczbhfw")){
                            sheet.getRow(index).getCell(3).setCellValue(datum.get("bzsczbhfw").toString());
                        }else {
                            sheet.getRow(index).getCell(3).setCellValue("");
                        }

                        if (datum.containsKey("zmddbz")){
                            if (!datum.get("zmddbz").equals("") && datum.get("zmddbz").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zmddbz").toString()));
                            }else {
                                sheet.getRow(index).getCell(4).setCellValue(datum.get("zmddbz").toString());
                            }

                        }else {
                            sheet.getRow(index).getCell(4).setCellValue("");
                        }
                        if (datum.containsKey("ymddbz")){
                            if (!datum.get("ymddbz").equals("") && datum.get("ymddbz").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ymddbz").toString()));
                            }else {
                                sheet.getRow(index).getCell(5).setCellValue(datum.get("ymddbz").toString());
                            }

                        }else {
                            sheet.getRow(index).getCell(5).setCellValue("");
                        }

                        if (datum.containsKey("bzmdgdz")){
                            if (!datum.get("bzmdgdz").equals("") && datum.get("bzmdgdz").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("bzmdgdz").toString()));
                            }else {
                                sheet.getRow(index).getCell(6).setCellValue("");
                            }
                        }else {
                            sheet.getRow(index).getCell(6).setCellValue("");
                        }

                        if (datum.containsKey("zxbzjcds")){
                            if (!datum.get("zxbzjcds").equals("") && datum.get("zxbzjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zxbzjcds").toString()));
                            }else {
                                sheet.getRow(index).getCell(7).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(7).setCellValue("");
                        }

                        if (datum.containsKey("zxbzhgds")){
                            if (!datum.get("zxbzhgds").equals("") && datum.get("zxbzhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zxbzhgds").toString()));
                            }else {
                                sheet.getRow(index).getCell(8).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(8).setCellValue("");
                        }

                        if (datum.containsKey("hgl")){
                            if (!datum.get("hgl").equals("") && datum.get("hgl").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("hgl").toString()));
                            }else {
                                sheet.getRow(index).getCell(9).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(9).setCellValue("");
                        }
                        index ++;

                }


            }else if (group.equals("最大理论密度")){
                XSSFSheet sheet = wb.getSheet("表4.1.2-1-1");
                int index = 3;
                for (Map<String, Object> datum : grouphtdData) {
                        sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                        sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
                        sheet.getRow(index).getCell(2).setCellValue(datum.get("bz").toString());

                        if (datum.containsKey("bzsczbhfw")){
                            sheet.getRow(index).getCell(3).setCellValue(datum.get("bzsczbhfw").toString());
                        }else {
                            sheet.getRow(index).getCell(3).setCellValue("");
                        }

                        if (datum.containsKey("zmddbz")){
                            if (!datum.get("zmddbz").equals("") && datum.get("zmddbz").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zmddbz").toString()));
                            }else {
                                sheet.getRow(index).getCell(4).setCellValue(datum.get("zmddbz").toString());
                            }

                        }else {
                            sheet.getRow(index).getCell(4).setCellValue("");
                        }
                        if (datum.containsKey("ymddbz")){
                            if (!datum.get("ymddbz").equals("") && datum.get("ymddbz").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ymddbz").toString()));
                            }else {
                                sheet.getRow(index).getCell(5).setCellValue(datum.get("ymddbz").toString());
                            }

                        }else {
                            sheet.getRow(index).getCell(5).setCellValue("");
                        }

                        if (datum.containsKey("bzmdgdz")){
                            if (!datum.get("bzmdgdz").equals("") && datum.get("bzmdgdz").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("bzmdgdz").toString()));
                            }else {
                                sheet.getRow(index).getCell(6).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(6).setCellValue("");
                        }

                        if (datum.containsKey("zxbzjcds")){
                            if (!datum.get("zxbzjcds").equals("") && datum.get("zxbzjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zxbzjcds").toString()));
                            }else {
                                sheet.getRow(index).getCell(7).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(7).setCellValue("");
                        }

                        if (datum.containsKey("zxbzhgds")){
                            if (!datum.get("zxbzhgds").equals("") && datum.get("zxbzhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zxbzhgds").toString()));
                            }else {
                                sheet.getRow(index).getCell(8).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(8).setCellValue("");
                        }

                        if (datum.containsKey("hgl")){
                            if (!datum.get("hgl").equals("") && datum.get("hgl").toString().matches("-?\\d+(\\.\\d+)?")){
                                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("hgl").toString()));
                            }else {
                                sheet.getRow(index).getCell(9).setCellValue("");
                            }

                        }else {
                            sheet.getRow(index).getCell(9).setCellValue("");
                        }
                    index ++;
                }
            }

        });

        /*XSSFSheet sheet = wb.getSheet("表4.1.2-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
                sheet.getRow(index).getCell(2).setCellValue(datum.get("bz").toString());

                if (datum.containsKey("bzsczbhfw")){
                    sheet.getRow(index).getCell(3).setCellValue(datum.get("bzsczbhfw").toString());
                }else {
                    sheet.getRow(index).getCell(3).setCellValue(0);
                }

                if (datum.containsKey("zmddbz")){
                    if (!datum.get("zmddbz").equals("")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zmddbz").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue(datum.get("zmddbz").toString());
                    }

                }else {
                    sheet.getRow(index).getCell(4).setCellValue(0);
                }
                if (datum.containsKey("ymddbz")){
                    if (!datum.get("ymddbz").equals("")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ymddbz").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue(datum.get("ymddbz").toString());
                    }

                }else {
                    sheet.getRow(index).getCell(5).setCellValue(0);
                }

                if (datum.containsKey("bzmdgdz")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("bzmdgdz").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue(0);
                }

                if (datum.containsKey("zxbzjcds")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zxbzjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue(0);
                }

                if (datum.containsKey("zxbzhgds")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zxbzhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue(0);
                }

                if (datum.containsKey("hgl")){
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("hgl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue(0);
                }
                index ++;
            }
        }*/
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelljhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {

        XSSFSheet sheet = wb.getSheet("表4.1.1-6");
        //System.out.println(sheet.getRow(3).getCell(2).getStringCellValue());
        int index = 3;
        for (Map<String, Object> datum : data) {
            //System.out.println(sheet.getRow(3).getCell(2).getStringCellValue());
            if (datum.size()>0){
                sheet.getRow(2).createCell(index).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("ysdjcds")){
                    if (!datum.get("ysdjcds").equals("") && datum.get("ysdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(3).createCell(index).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
                    }else {
                        sheet.getRow(3).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(3).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgds")){
                    if (!datum.get("ysdhgds").equals("") && datum.get("ysdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(4).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
                    }else {
                        sheet.getRow(4).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(4).createCell(index).setCellValue("");
                }

                if (datum.containsKey("ysdhgl")){
                    if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(5).createCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                    }else {
                        sheet.getRow(5).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(5).createCell(index).setCellValue("");
                }
                if (datum.containsKey("cjjcds")){
                    if (!datum.get("cjjcds").equals("") && datum.get("cjjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(6).createCell(index).setCellValue(Double.valueOf(datum.get("cjjcds").toString()));
                    }else {
                        sheet.getRow(6).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(6).createCell(index).setCellValue("");
                }
                if (datum.containsKey("cjhgds")){
                    if (!datum.get("cjhgds").equals("") && datum.get("cjhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(7).createCell(index).setCellValue(Double.valueOf(datum.get("cjhgds").toString()));
                    }else {
                        sheet.getRow(7).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(7).createCell(index).setCellValue("");
                }
                if (datum.containsKey("cjhgl")){
                    if (!datum.get("cjhgl").equals("") && datum.get("cjhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(8).createCell(index).setCellValue(Double.valueOf(datum.get("cjhgl").toString()));
                    }else {
                        sheet.getRow(8).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(8).createCell(index).setCellValue("");
                }

                if (datum.containsKey("wcjcds")){
                    if (!datum.get("wcjcds").equals("") && datum.get("wcjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(9).createCell(index).setCellValue(Double.valueOf(datum.get("wcjcds").toString()));
                    }else {
                        sheet.getRow(9).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(9).createCell(index).setCellValue("");
                }

                if (datum.containsKey("wchgds")){
                    if (!datum.get("wchgds").equals("") && datum.get("wchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(10).createCell(index).setCellValue(Double.valueOf(datum.get("wchgds").toString()));
                    }else {
                        sheet.getRow(10).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(10).createCell(index).setCellValue("");
                }

                if (datum.containsKey("wchgl")){
                    if (!datum.get("wchgl").equals("") && datum.get("wchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(11).createCell(index).setCellValue(Double.valueOf(datum.get("wchgl").toString()));
                    }else {
                        sheet.getRow(11).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(11).createCell(index).setCellValue("");
                }

                if (datum.containsKey("bpjcds")){
                    if (!datum.get("bpjcds").equals("") && datum.get("bpjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(12).createCell(index).setCellValue(Double.valueOf(datum.get("bpjcds").toString()));
                    }else {
                        sheet.getRow(12).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(12).createCell(index).setCellValue("");
                }
                if (datum.containsKey("bphgds")){
                    if (!datum.get("bphgds").equals("") && datum.get("bphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(13).createCell(index).setCellValue(Double.valueOf(datum.get("bphgds").toString()));
                    }
                }else {
                    sheet.getRow(13).createCell(index).setCellValue("");
                }

                if (datum.containsKey("bphgl")){
                    if (!datum.get("bphgl").equals("") && datum.get("bphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(14).createCell(index).setCellValue(Double.valueOf(datum.get("bphgl").toString()));
                    }else {
                        sheet.getRow(14).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(14).createCell(index).setCellValue("");
                }

                if (datum.containsKey("psdmccjcds")){
                    if (!datum.get("psdmccjcds").equals("") && datum.get("psdmccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(15).createCell(index).setCellValue(Double.valueOf(datum.get("psdmccjcds").toString()));
                    }else {
                        sheet.getRow(15).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(15).createCell(index).setCellValue("");
                }

                if (datum.containsKey("psdmcchgds")){
                    if (!datum.get("psdmcchgds").equals("") && datum.get("psdmcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(16).createCell(index).setCellValue(Double.valueOf(datum.get("psdmcchgds").toString()));
                    }else {
                        sheet.getRow(16).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(16).createCell(index).setCellValue("");
                }
                if (datum.containsKey("psdmcchgl")){
                    if (!datum.get("psdmcchgl").equals("") && datum.get("psdmcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(17).createCell(index).setCellValue(Double.valueOf(datum.get("psdmcchgl").toString()));
                    }else {
                        sheet.getRow(17).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(17).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pspqhdjcds")){
                    if (!datum.get("pspqhdjcds").equals("") && datum.get("pspqhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(18).createCell(index).setCellValue(Double.valueOf(datum.get("pspqhdjcds").toString()));
                    }else {
                        sheet.getRow(18).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(18).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pspqhdhgds")){
                    if (!datum.get("pspqhdhgds").equals("") && datum.get("pspqhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(19).createCell(index).setCellValue(Double.valueOf(datum.get("pspqhdhgds").toString()));
                    }else {
                        sheet.getRow(19).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(19).createCell(index).setCellValue("");
                }

                if (datum.containsKey("pspqhdhgl")){
                    if (!datum.get("pspqhdhgl").equals("") && datum.get("pspqhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(20).createCell(index).setCellValue(Double.valueOf(datum.get("pspqhdhgl").toString()));
                    }else {
                        sheet.getRow(20).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(20).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xqtqdjcds")){
                    if (!datum.get("xqtqdjcds").equals("") && datum.get("xqtqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(21).createCell(index).setCellValue(Double.valueOf(datum.get("xqtqdjcds").toString()));
                    }else {
                        sheet.getRow(21).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(21).createCell(index).setCellValue("");
                }
                if (datum.containsKey("xqtqdhgds")){
                    if (!datum.get("xqtqdhgds").equals("") && datum.get("xqtqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(22).createCell(index).setCellValue(Double.valueOf(datum.get("xqtqdhgds").toString()));
                    }else {
                        sheet.getRow(22).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(22).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xqtqdhgl")){
                    if (!datum.get("xqtqdhgl").equals("") && datum.get("xqtqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(23).createCell(index).setCellValue(Double.valueOf(datum.get("xqtqdhgl").toString()));
                    }else {
                        sheet.getRow(23).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(23).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xqjgccjcds")){
                    if (!datum.get("xqjgccjcds").equals("") && datum.get("xqjgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(24).createCell(index).setCellValue(Double.valueOf(datum.get("xqjgccjcds").toString()));
                    }else {
                        sheet.getRow(24).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(24).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xqjgcchgds")){
                    if (!datum.get("xqjgcchgds").equals("") && datum.get("xqjgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(25).createCell(index).setCellValue(Double.valueOf(datum.get("xqjgcchgds").toString()));
                    }else {
                        sheet.getRow(25).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(25).createCell(index).setCellValue("");
                }

                if (datum.containsKey("xqjgcchgl")){
                    if (!datum.get("xqjgcchgl").equals("") && datum.get("xqjgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(26).createCell(index).setCellValue(Double.valueOf(datum.get("xqjgcchgl").toString()));
                    }else {
                        sheet.getRow(26).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(26).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdtqdjcds")){
                    if (!datum.get("hdtqdjcds").equals("") && datum.get("hdtqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(27).createCell(index).setCellValue(Double.valueOf(datum.get("hdtqdjcds").toString()));
                    }else {
                        sheet.getRow(27).createCell(index).setCellValue("");
                    }


                }else {
                    sheet.getRow(27).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdtqdhgds")){
                    if (!datum.get("hdtqdhgds").equals("") && datum.get("hdtqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(28).createCell(index).setCellValue(Double.valueOf(datum.get("hdtqdhgds").toString()));
                    }else {
                        sheet.getRow(28).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(28).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdtqdhgl")){
                    if (!datum.get("hdtqdhgl").equals("") && datum.get("hdtqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(29).createCell(index).setCellValue(Double.valueOf(datum.get("hdtqdhgl").toString()));
                    }else {
                        sheet.getRow(29).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(29).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdjgccjcds")){
                    if (!datum.get("hdjgccjcds").equals("") && datum.get("hdjgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(30).createCell(index).setCellValue(Double.valueOf(datum.get("hdjgccjcds").toString()));
                    }else {
                        sheet.getRow(30).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(30).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdjgcchgds")){
                    if (!datum.get("hdjgcchgds").equals("") && datum.get("hdjgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(31).createCell(index).setCellValue(Double.valueOf(datum.get("hdjgcchgds").toString()));
                    }else {
                        sheet.getRow(31).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(31).createCell(index).setCellValue("");
                }

                if (datum.containsKey("hdjgcchgl")){
                    if (!datum.get("hdjgcchgl").equals("") && datum.get("hdjgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(32).createCell(index).setCellValue(Double.valueOf(datum.get("hdjgcchgl").toString()));
                    }else {
                        sheet.getRow(32).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(32).createCell(index).setCellValue("");
                }

                if (datum.containsKey("zdtqdjcds")){
                    if (!datum.get("zdtqdjcds").equals("") && datum.get("zdtqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(33).createCell(index).setCellValue(Double.valueOf(datum.get("zdtqdjcds").toString()));
                    }else {
                        sheet.getRow(33).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(33).createCell(index).setCellValue("");
                }

                if (datum.containsKey("zdtqdhgds")){
                    if (!datum.get("zdtqdhgds").equals("") && datum.get("zdtqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(34).createCell(index).setCellValue(Double.valueOf(datum.get("zdtqdhgds").toString()));
                    } else {
                        sheet.getRow(34).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(34).createCell(index).setCellValue("");
                }

                if (datum.containsKey("zdtqdhgl")){
                    if (!datum.get("zdtqdhgl").equals("") && datum.get("zdtqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(35).createCell(index).setCellValue(Double.valueOf(datum.get("zdtqdhgl").toString()));
                    }else {
                        sheet.getRow(35).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(35).createCell(index).setCellValue("");
                }

                if (datum.containsKey("zddmccjcds")){
                    if (!datum.get("zddmccjcds").equals("") && datum.get("zddmccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(36).createCell(index).setCellValue(Double.valueOf(datum.get("zddmccjcds").toString()));
                    }else {
                        sheet.getRow(36).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(36).createCell(index).setCellValue("");
                }

                if (datum.containsKey("zddmcchgds")){
                    if (!datum.get("zddmcchgds").equals("") && datum.get("zddmcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(37).createCell(index).setCellValue(Double.valueOf(datum.get("zddmcchgds").toString()));
                    }else {
                        sheet.getRow(37).createCell(index).setCellValue("");
                    }

                }else {
                    sheet.getRow(37).createCell(index).setCellValue("");
                }
                if (datum.containsKey("zddmcchgl")){
                    if (!datum.get("zddmcchgl").equals("") && datum.get("zddmcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(38).createCell(index).setCellValue(Double.valueOf(datum.get("zddmcchgl").toString()));
                    }else {
                        sheet.getRow(38).createCell(index).setCellValue("");
                    }
                }else {
                    sheet.getRow(38).createCell(index).setCellValue("");
                }
                index ++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.1-5");
        XSSFSheet sheet = wb.getSheet("表4.1.1-5");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("zdtqdjcds")){
                if (!datum.get("zdtqdjcds").equals("") && datum.get("zdtqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("zdtqdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("zdtqdhgds")){
                if (!datum.get("zdtqdhgds").equals("") && datum.get("zdtqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("zdtqdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("zdtqdhgl")){
                if (!datum.get("zdtqdhgl").equals("") && datum.get("zdtqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("zdtqdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("zddmccjcds")){
                if (!datum.get("zddmccjcds").equals("") && datum.get("zddmccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zddmccjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("zddmcchgds")){
                if (!datum.get("zddmcchgds").equals("") && datum.get("zddmcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("zddmcchgds").toString()));
                } else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }

            if (datum.containsKey("zddmcchgl")){
                if (!datum.get("zddmcchgl").equals("") && datum.get("zddmcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("zddmcchgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            index ++;
        }


    }

    /**
     *
     * @param wb
     */
    private void DBExcelhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.1-4");
        XSSFSheet sheet = wb.getSheet("表4.1.1-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hdtqdjcds")){
                if (!datum.get("hdtqdjcds").equals("") && datum.get("hdtqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hdtqdjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }

            if (datum.containsKey("hdtqdhgds")){
                if (!datum.get("hdtqdhgds").equals("") && datum.get("hdtqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hdtqdhgds").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(2).setCellValue("");
            }

            if (datum.containsKey("hdtqdhgl")){
                if (!datum.get("hdtqdhgl").equals("") && datum.get("hdtqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hdtqdhgl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(3).setCellValue("");
            }

            if (datum.containsKey("hdjgccjcds")){
                if (!datum.get("hdjgccjcds").equals("") && datum.get("hdjgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hdjgccjcds").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
            }else {
                sheet.getRow(index).getCell(4).setCellValue("");
            }

            if (datum.containsKey("hdjgcchgds")){
                if (!datum.get("hdjgcchgds").equals("") && datum.get("hdjgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hdjgcchgds").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(5).setCellValue("");
            }
            if (datum.containsKey("hdjgcchgl")){
                if (!datum.get("hdjgcchgl").equals("") && datum.get("hdjgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hdjgcchgl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

            }else {
                sheet.getRow(index).getCell(6).setCellValue("");
            }

            index ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelxqData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.1-3");
        XSSFSheet sheet = wb.getSheet("表4.1.1-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("xqtqdjcds")){
                    if (!datum.get("xqtqdjcds").equals("") && datum.get("xqtqdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("xqtqdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
                if (datum.containsKey("xqtqdhgds")){
                    if (!datum.get("xqtqdhgds").equals("") && datum.get("xqtqdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("xqtqdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

                if (datum.containsKey("xqtqdhgl")){
                    if (!datum.get("xqtqdhgl").equals("") && datum.get("xqtqdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("xqtqdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("xqjgccjcds")){
                    if (!datum.get("xqjgccjcds").equals("") && datum.get("xqjgccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("xqjgccjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

                if (datum.containsKey("xqjgcchgds")){
                    if (!datum.get("xqjgcchgds").equals("") && datum.get("xqjgcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("xqjgcchgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }
                if (datum.containsKey("xqjgcchgl")){
                    if (!datum.get("xqjgcchgl").equals("") && datum.get("xqjgcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("xqjgcchgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
                index ++;
            }

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelpsData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.1-2");
        XSSFSheet sheet = wb.getSheet("表4.1.1-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("psdmccjcds")){
                    if (!datum.get("psdmccjcds").equals("") && datum.get("psdmccjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("psdmccjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }
                if (datum.containsKey("psdmcchgds")){
                    if (!datum.get("psdmcchgds").equals("") && datum.get("psdmcchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("psdmcchgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }
                if (datum.containsKey("psdmcchgl")){
                    if (!datum.get("psdmcchgl").equals("") && datum.get("psdmcchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("psdmcchgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("pspqhdjcds")){
                    if (!datum.get("pspqhdjcds").equals("") && datum.get("pspqhdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("pspqhdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }

                if (datum.containsKey("pspqhdhgds")){
                    if (!datum.get("pspqhdhgds").equals("") && datum.get("pspqhdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("pspqhdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

                if (datum.containsKey("pspqhdhgl")){
                    if (!datum.get("pspqhdhgl").equals("") && datum.get("pspqhdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("pspqhdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }
                index ++;
            }
        }
    }

    /**

     * @param wb
     */
    private void DBExceltsfData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.1-1");
        XSSFSheet sheet = wb.getSheet("表4.1.1-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            if (datum.size()>0){
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("ysdjcds")){
                    if (!datum.get("ysdjcds").equals("") && datum.get("ysdjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(1).setCellValue("");
                }

                if (datum.containsKey("ysdhgds")){
                    if (!datum.get("ysdhgds").equals("") && datum.get("ysdhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(2).setCellValue("");
                }

                if (datum.containsKey("ysdhgl")){
                    if (!datum.get("ysdhgl").equals("") && datum.get("ysdhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(3).setCellValue("");
                }

                if (datum.containsKey("cjjcds")){
                    if (!datum.get("cjjcds").equals("") && datum.get("cjjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cjjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(4).setCellValue("");
                }
                if (datum.containsKey("cjhgds")){
                    if (!datum.get("cjhgds").equals("") && datum.get("cjhgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cjhgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(5).setCellValue("");
                }

                if (datum.containsKey("cjhgl")){
                    if (!datum.get("cjhgl").equals("") && datum.get("cjhgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cjhgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue("");
                }

                if (datum.containsKey("wcjcds")){
                    if (!datum.get("wcjcds").equals("") && datum.get("wcjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("wcjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(7).setCellValue("");
                }

                if (datum.containsKey("wchgds")){
                    if (!datum.get("wchgds").equals("") && datum.get("wchgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("wchgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(8).setCellValue("");
                }

                if (datum.containsKey("wchgl")){
                    if (!datum.get("wchgl").equals("") && datum.get("wchgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("wchgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(9).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(9).setCellValue("");
                }

                if (datum.containsKey("bpjcds")){
                    if (!datum.get("bpjcds").equals("") && datum.get("bpjcds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("bpjcds").toString()));
                    }else {
                        sheet.getRow(index).getCell(10).setCellValue("");
                    }
                }else {
                    sheet.getRow(index).getCell(10).setCellValue("");
                }

                if (datum.containsKey("bphgds")){
                    if (!datum.get("bphgds").equals("") && datum.get("bphgds").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("bphgds").toString()));
                    }else {
                        sheet.getRow(index).getCell(11).setCellValue("");

                    }
                }else {
                    sheet.getRow(index).getCell(11).setCellValue("");
                }
                if (datum.containsKey("bphgl")){
                    if (!datum.get("bphgl").equals("") && datum.get("bphgl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("bphgl").toString()));
                    }else {
                        sheet.getRow(index).getCell(12).setCellValue("");
                    }

                }else {
                    sheet.getRow(index).getCell(12).setCellValue("");
                }
                index ++;
            }

        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelSdcjData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表3.4.4-1");
        XSSFSheet sheet = wb.getSheet("表3.4.4-1");
        DecimalFormat df = new DecimalFormat("0.0");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            double tcccs = 0;
            double tccjs = 0;
            if (datum.containsKey("tcccs")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("tcccs").toString()));
                tcccs = Double.valueOf(datum.get("tcccs").toString());

            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("tccjs")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("tccjs").toString()));
                tccjs = Double.valueOf(datum.get("tccjs").toString());

            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            //String a= tccjs != 0 ? df.format(tcccs/tccjs*100) : "0";
            String a= tcccs != 0 ? df.format(tccjs/tcccs*100) : "0";
            sheet.getRow(index).getCell(3).setCellValue(a);

            double cccs = 0;
            double ccjs = 0;
            if (datum.containsKey("cccs")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cccs").toString()));
                cccs = Double.valueOf(datum.get("cccs").toString());
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("ccjs")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ccjs").toString()));
                ccjs = Double.valueOf(datum.get("ccjs").toString());
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            //String b= ccjs != 0 ? df.format(cccs/ccjs*100) : "0";
            String b= cccs != 0 ? df.format(ccjs/cccs*100) : "0";
            sheet.getRow(index).getCell(6).setCellValue(b);

            double zccs = 0;
            double zcjs = 0;
            if (datum.containsKey("zccs")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zccs").toString()));
                zccs = Double.valueOf(datum.get("zccs").toString());
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            if (datum.containsKey("zcjs")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zcjs").toString()));
                zcjs = Double.valueOf(datum.get("zcjs").toString());
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }
            //String c= zcjs != 0 ? df.format(zccs/zcjs*100) : "0";
            String c= zccs != 0 ? df.format(zcjs/zccs*100) : "0";
            sheet.getRow(index).getCell(9).setCellValue(c);

            double dccs = 0;
            double dcjs = 0;
            if (datum.containsKey("dccs")){
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("dccs").toString()));
                dccs = Double.valueOf(datum.get("dccs").toString());
            }else {
                sheet.getRow(index).getCell(10).setCellValue(0);
            }
            if (datum.containsKey("dcjs")){
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("dcjs").toString()));
                dcjs = Double.valueOf(datum.get("dcjs").toString());
            }else {
                sheet.getRow(index).getCell(11).setCellValue(0);
            }
            //String d= dcjs != 0 ? df.format(dccs/dcjs*100) : "0";
            String d= dccs != 0 ? df.format(dcjs/dccs*100) : "0";
            sheet.getRow(index).getCell(12).setCellValue(d);

            index ++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     * @param proname
     */
    private void DBExcelJaData(XSSFWorkbook wb, List<Map<String, Object>> data, String proname) {
        //createRow(wb,data.size(),"表3.4.3-1");
        XSSFSheet sheet = wb.getSheet("表3.4.3-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("bzsys")){
                if (!datum.get("bzsys").equals("") && datum.get("bzsys").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("bzsys").toString()));
                }else {
                    sheet.getRow(index).getCell(1).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("bzccs")){
                if (!datum.get("bzccs").equals("") && datum.get("bzccs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("bzccs").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            if (datum.containsKey("bzcjpl")){
                if (!datum.get("bzcjpl").equals("") && datum.get("bzcjpl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("bzcjpl").toString()));
                }else {
                    sheet.getRow(index).getCell(3).setCellValue(0);
                }

            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("bxsys")){
                if (!datum.get("bxsys").equals("") && datum.get("bxsys").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("bxsys").toString()));
                }else {
                    sheet.getRow(index).getCell(4).setCellValue(0);
                }

            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("bxccs")){
                if (!datum.get("bxccs").equals("") && datum.get("bxccs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("bxccs").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("bxcjpl")){
                if (!datum.get("bxcjpl").equals("") && datum.get("bxcjpl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("bxcjpl").toString()));
                }else {
                    sheet.getRow(index).getCell(6).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            if (datum.containsKey("fhlsys")){
                if (!datum.get("fhlsys").equals("") && datum.get("fhlsys").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("fhlsys").toString()));
                }else {
                    sheet.getRow(index).getCell(7).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            if (datum.containsKey("fhlccs")){
                if (!datum.get("fhlccs").equals("") && datum.get("fhlccs").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("fhlccs").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }
            if (datum.containsKey("fhlcjpl")){
                if (!datum.get("fhlcjpl").equals("") && datum.get("fhlcjpl").toString().matches("-?\\d+(\\.\\d+)?")){
                    sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("fhlcjpl").toString()));
                }else {
                    sheet.getRow(index).getCell(9).setCellValue(0);
                }
            }else {
                sheet.getRow(index).getCell(9).setCellValue(0);
            }

            index ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     * @param proname
     */
    private void DBExcelQlData(XSSFWorkbook wb, List<Map<String, Object>> data, String proname) throws IOException {
        if (data.size()>0){
            //createRow(wb,data.size(),"表3.4.2-1");
            DecimalFormat df = new DecimalFormat("0.0");
            XSSFSheet sheet = wb.getSheet("表3.4.2-1");

            int index = 3;
            for (Map<String, Object> datum : data) {
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

                double tdccs = 0;
                double tdcjs = 0;
                if (datum.containsKey("tdcjs")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("tdcjs").toString()));
                    tdcjs = Integer.valueOf(datum.get("tdcjs").toString());
                }else {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
                }
                if (datum.containsKey("tdccs")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("tdccs").toString()));
                    tdccs = Integer.valueOf(datum.get("tdccs").toString());
                }else {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(0));
                }
                //String a= tdcjs != 0 ? df.format(tdccs/tdcjs*100) : "0";
                String a= tdccs != 0 ? df.format(tdcjs/tdccs*100) : "0";
                sheet.getRow(index).getCell(3).setCellValue(a);


                double dccs = 0;
                double dcjs = 0;
                if (datum.containsKey("dcjs")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("dcjs").toString()));
                    dcjs = Double.valueOf(datum.get("dcjs").toString());
                }else {
                    sheet.getRow(index).getCell(5).setCellValue(0);
                }
                if (datum.containsKey("dccs")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("dccs").toString()));
                    dccs = Double.valueOf(datum.get("dccs").toString());
                }else {
                    sheet.getRow(index).getCell(4).setCellValue(0);
                }
                //String b= dcjs != 0 ? df.format(dccs/dcjs*100) : "0";
                String b= dccs != 0 ? df.format(dcjs/dccs*100) : "0";
                sheet.getRow(index).getCell(6).setCellValue(b);

                double zccs = 0;
                double zcjs = 0;
                if (datum.containsKey("zcjs")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zcjs").toString()));
                    zcjs = Double.valueOf(datum.get("zcjs").toString());
                }else {
                    sheet.getRow(index).getCell(8).setCellValue(0);
                }
                if (datum.containsKey("zccs")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zccs").toString()));
                    zccs = Double.valueOf(datum.get("zccs").toString());
                }else {
                    sheet.getRow(index).getCell(7).setCellValue(0);
                }
                //String c = zcjs != 0 ? df.format(zccs/zcjs*100) : "0";
                String c = zccs != 0 ? df.format(zcjs/zccs*100) : "0";
                sheet.getRow(index).getCell(9).setCellValue(c);

                double xccs = 0;
                double xcjs = 0;
                if (datum.containsKey("xcjs")){
                    sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("xcjs").toString()));
                    xcjs = Double.valueOf(datum.get("xcjs").toString());
                }else {
                    sheet.getRow(index).getCell(11).setCellValue(0);
                }

                if (datum.containsKey("xccs")){
                    sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("xccs").toString()));
                    xccs = Double.valueOf(datum.get("xccs").toString());
                }else {
                    sheet.getRow(index).getCell(10).setCellValue(0);
                }
                //String d = xcjs != 0 ? df.format(xccs/xcjs*100) : "0";
                String d = xccs != 0 ? df.format(xcjs/xccs*100) : "0";
                sheet.getRow(index).getCell(12).setCellValue(d);
                index ++;
            }


        }

    }

    /**
     *
     *
     * @param wb
     * @param data
     * @param proname
     * @throws IOException
     */
    private void DBExcelData(XSSFWorkbook wb, List<Map<String, Object>> data, String proname) throws IOException {
        if (data.size()>0){
            //createRow(wb,data.size(),"表3.4.1-1");
            XSSFSheet sheet = wb.getSheet("表3.4.1-1");
            int index = 3;
            for (Map<String, Object> datum : data) {
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("hdsys")){
                    if (!datum.get("hdsys").equals("") && datum.get("hdsys").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hdsys").toString()));
                    }else {
                        sheet.getRow(index).getCell(1).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(1).setCellValue(0);
                }

                if (datum.containsKey("hdccs")){
                    if (!datum.get("hdccs").equals("") && datum.get("hdccs").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hdccs").toString()));
                    }else {
                        sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
                    }
                }else {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
                }

                if (datum.containsKey("hdcjpl")){
                    if (!datum.get("hdcjpl").equals("") && datum.get("hdcjpl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hdcjpl").toString()));
                    }else {
                        sheet.getRow(index).getCell(3).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(3).setCellValue(0);
                }

                if (datum.containsKey("zdsys")){
                    if (!datum.get("zdsys").equals("") && datum.get("zdsys").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zdsys").toString()));
                    }else {
                        sheet.getRow(index).getCell(4).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(4).setCellValue(0);
                }

                if (datum.containsKey("zdccs")){
                    if (!datum.get("zdccs").equals("") && datum.get("zdccs").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("zdccs").toString()));
                    }else {
                        sheet.getRow(index).getCell(5).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(5).setCellValue(0);
                }

                if (datum.containsKey("zdcjpl")){
                    if (!datum.get("zdcjpl").equals("") && datum.get("zdcjpl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("zdcjpl").toString()));
                    }else {
                        sheet.getRow(index).getCell(6).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(6).setCellValue(0);
                }

                if (datum.containsKey("xqsys")){
                    if (!datum.get("xqsys").equals("") && datum.get("xqsys").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("xqsys").toString()));
                    }else {
                        sheet.getRow(index).getCell(7).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(7).setCellValue(0);
                }
                if (datum.containsKey("xqccs")){
                    if (!datum.get("xqccs").equals("") && datum.get("xqccs").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("xqccs").toString()));
                    }else {
                        sheet.getRow(index).getCell(8).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(8).setCellValue(0);
                }

                if (datum.containsKey("xqcjpl")){
                    if (!datum.get("xqcjpl").equals("") && datum.get("xqcjpl").toString().matches("-?\\d+(\\.\\d+)?")){
                        sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("xqcjpl").toString()));
                    }else {
                        sheet.getRow(index).getCell(9).setCellValue(0);
                    }
                }else {
                    sheet.getRow(index).getCell(9).setCellValue(0);
                }
                index ++;
            }
        }
    }

    /**
     *
     * @param wb
     * @param gettableNum
     * @param sheetname
     */
    private void createRow(XSSFWorkbook wb,int gettableNum,String sheetname) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 3, 4, i*1);
        }
    }


    /**
     * 表5.3-1  建设项目工程质量评定汇总表
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getxsxmpdData(String proname) throws IOException {
        String path = filespath+ File.separator+proname+File.separator;
        List<String> filteredFiles = filterFiles(path);

        List<Map<String,Object>> mapList = new ArrayList<>();

        for (String filteredFile : filteredFiles) {
            String pdnpath = path+filteredFile+ File.separator;
            XSSFWorkbook wb = null;
            File f = new File(pdnpath+"00评定表.xlsx");
            if (!f.exists()){
                break;
            }else {
                FileInputStream in = new FileInputStream(f);
                wb = new XSSFWorkbook(in);
                XSSFSheet sheet = wb.getSheet("合同段");
                Map map = new HashMap();
                for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) {
                    XSSFRow row = sheet.getRow(i);

                    if (row.getCell(0).getStringCellValue().equals("合同段质量等级")){
                        String djValue = row.getCell(1).getStringCellValue();
                        map.put("htdValue",sheet.getRow(1).getCell(1).getStringCellValue());
                        map.put("djValue",djValue);
                        mapList.add(map);
                    }

                }
            }
        }
        return mapList;
    }

    /**
     * 表5.2-1  合同段工程质量评定汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> gethtdpdData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> resultlist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()) {
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            List<Map<String, Object>> list = processhtdSheet(wb);
            resultlist.addAll(list);
        }
        return resultlist;
    }

    /**
     * 表5.1.3-1  隧道单位工程质量评定汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdpdData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()){
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.contains("隧道")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
        }
        return dwgclist;

    }

    /**
     * 表5.1.2-1  桥梁单位工程质量评定汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqlpdData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()){
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.contains("桥")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
        }
        return dwgclist;

    }

    /**
     * 表5.1.1-1  路基单位工程质量评定汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getljhzbData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()){
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);


            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.equals("分部-路基")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
        }
        return dwgclist;
    }

    /**
     * 表4.1.6-3  隧道工程衬砌厚度检测数据分析表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdcqfxbData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> cqhd = jjgFbgcSdgcCqhdService.lookJdbjg(commonInfoVo);//[{合格点数=96, 检测总点数=96, 检测项目=李家湾隧道左线, 合格率=100.00}, {合格点数=96, 检测总点数=96, 检测项目=李家湾隧道右线, 合格率=100.00}]
        if (cqhd!=null && cqhd.size()>0){
            for (Map<String, Object> map : cqhd) {
                map.put("htd",commonInfoVo.getHtd());
                String sdmc = map.get("检测项目").toString();
                int num = jjgFbgcSdgcCqhdService.getds(commonInfoVo,sdmc);
                map.put("ds",num);
                resultList.add(map);
            }
        }
        return resultList;

    }

    /**
     * 表4.1.6-2  桥梁下部墩台竖直度检测数据分析表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getqlszdfxbData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> szd = jjgFbgcQlgcXbSzdService.lookJdbjg(commonInfoVo);
        if (szd!=null && szd.size()>0){
            /*String[] valus = szd.get(0).get("yxpc").toString().replace("[", "").replace("]", "").split(",");
            for (int i =0; i< valus.length;i++){
                Map map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("gdz",valus[i]);
                map.put("max",getmaxvalue(szd.get(0).get("scz").toString()));
                map.put("min",getminvalue(szd.get(0).get("scz").toString()));
                map.put("jcds",szd.get(0).get("总点数"));
                map.put("hgds",szd.get(0).get("合格点数"));
                map.put("hgl",szd.get(0).get("合格率"));
                resultList.add(map);
            }*/
            Map map = new HashMap<>();
            map.put("htd",commonInfoVo.getHtd());
            //map.put("gdz","");
            map.put("max",getmaxvalue(szd.get(0).get("scz").toString()));
            map.put("min",getminvalue(szd.get(0).get("scz").toString()));
            map.put("jcds",szd.get(0).get("总点数"));
            map.put("hgds",szd.get(0).get("合格点数"));
            map.put("hgl",szd.get(0).get("合格率"));
            resultList.add(map);

        }
        return resultList;

    }

    /**
     * 表4.1.6-1  路基工程压实度检测数据分析表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getysdfxbData(CommonInfoVo commonInfoVo) throws IOException {

        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list1 = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
        //[{合格点数=7, 压实度值=[96.5, 98.2, 97.0, 95.5, 98.3, 99.1, 97.0], 检测点数=7, 平均值=97.37, 规定值=96, 结果=合格, 压实度项目=灰土, 合格率=100.00, 标准差=1.230, 代表值=96.70}]
        if (list1!=null && list1.size()>0){

            double jd=0,hgd=0;
            int index = 0; // 用来判断list1的第几个
            for (Map<String, Object> stringObjectMap : list1) {
                Map map = new HashMap<>(); // add会添加引用，因此必须重新生成map
                String js = stringObjectMap.get("检测点数").toString();
                String hg = stringObjectMap.get("合格点数").toString();
                jd = Double.valueOf(js);
                hgd = Double.valueOf(hg);
                map.put("htd",commonInfoVo.getHtd());
                map.put("bzz",list1.get(index).get("规定值"));
                map.put("jz",Double.valueOf(list1.get(index).get("规定值").toString())-5);
                map.put("max",getmaxvalue(list1.get(index).get("压实度值").toString()));
                map.put("min",getminvalue(list1.get(index).get("压实度值").toString()));
                map.put("dbz",list1.get(index).get("代表值"));

                map.put("jcds",decf.format(jd));
                map.put("hgds",decf.format(hgd));
                map.put("hgl",jd!=0 ? df.format(hgd/jd*100) : 0);
                resultList.add(map);
                index++;
            }

        }
        return resultList;
    }


    /**
     * 表4.1.5-5  交通安全设施检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getjabzData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        resultmap.put("htd",commonInfoVo.getHtd());
        List<Map<String, Object>> list1 = getbzData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getbxData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getfhlData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list4 = getqdccData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            for (Map<String, Object> map : list4) {
                resultmap.putAll(map);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     *表4.1.5-4  防护栏（砼防护栏）检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqdccData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list1 = jjgFbgcJtaqssJathlqdService.lookJdbjg(commonInfoVo);
        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        if (list1!=null && list1.size()>0){
            map.put("qdzds",list1.get(0).get("总点数"));
            map.put("qdhgs",list1.get(0).get("合格点数"));
            map.put("qdhgl",list1.get(0).get("合格率"));

        }

        List<Map<String, Object>> list2 = jjgFbgcJtaqssJathldmccService.lookJdbjg(commonInfoVo);
        if (list2!=null && list2.size()>0){
            map.put("cczds",list2.get(0).get("检测总点数"));
            map.put("cchgs",list2.get(0).get("合格点数"));
            map.put("cchgl",list2.get(0).get("合格率"));

        }
        resultList.add(map);
        return resultList;

    }

    /**
     *表4.1.5-3  防护栏（波形梁）检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getfhlData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcJtaqssJabxfhlService.lookJdbjg(commonInfoVo);
        //[{不合格点数=2, 合格点数=2, 总点数=2, 检测项目=波形梁板基底金属厚度, 合格率=100.00, 规定值或允许偏差=≥4},
        // {不合格点数=2, 合格点数=2, 总点数=2, 检测项目=波形梁钢护栏立柱壁厚, 合格率=100.00, 规定值或允许偏差=方柱6.0},
        // {不合格点数=0, 合格点数=0, 总点数=6, 检测项目=波形梁钢护栏横梁中心高度, 合格率=0.00, 规定值或允许偏差=三波板600±20},
        // {不合格点数=2, 合格点数=2, 总点数=2, 检测项目=波形梁钢护栏立柱埋入深度, 合格率=100.00, 规定值或允许偏差=方柱≥1650.0}]
        if (list!=null && list.size()>0){
            Map resultmap = new LinkedHashMap<>();
            resultmap.put("htd",commonInfoVo.getHtd());
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                if (jcxm.equals("波形梁板基底金属厚度")){
                    resultmap.put("jzds",map.get("总点数"));
                    resultmap.put("jhgds",map.get("合格点数"));
                    resultmap.put("jhgl",map.get("合格率"));

                }else if (jcxm.equals("波形梁钢护栏立柱壁厚")){
                    resultmap.put("lzds",map.get("总点数"));
                    resultmap.put("lhgds",map.get("合格点数"));
                    resultmap.put("lhgl",map.get("合格率"));

                }else if (jcxm.equals("波形梁钢护栏立柱埋入深度")){
                    resultmap.put("szds",map.get("总点数"));
                    resultmap.put("shgds",map.get("合格点数"));
                    resultmap.put("shgl",map.get("合格率"));

                }else if (jcxm.equals("波形梁钢护栏横梁中心高度")){
                    resultmap.put("gzds",map.get("总点数"));
                    resultmap.put("ghgds",map.get("合格点数"));
                    resultmap.put("ghgl",map.get("合格率"));
                }
            }
            resultList.add(resultmap);
        }
        return resultList;

    }

    /**
     * 表4.1.5-2  标线检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getbxData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
        //[{不合格点数=33, 合格点数=897, 总点数=930, 检测项目=交安标线厚度, 合格率=96.45, 规定值或允许偏差=白、黄线2+0.5,-0.1白、黄线7+10},
        // {不合格点数=11, 合格点数=919, 总点数=930, 检测项目=交安标线白线逆反射系数, 合格率=98.82, 规定值或允许偏差=震动≥150;双组分≥350;热熔≥150}]
        if (list!=null && list.size()>0){
            Map resultmap = new LinkedHashMap<>();
            resultmap.put("htd",commonInfoVo.getHtd());
            double zds = 0.0;
            double hgds = 0.0;
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                if (jcxm.equals("交安标线厚度")){
                    resultmap.put("bxhdzds",map.get("总点数"));
                    resultmap.put("bxhdhgds",map.get("合格点数"));
                    resultmap.put("bxhdhgl",map.get("合格率"));
                }else {
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("xzds",decf.format(zds));
            resultmap.put("xhgds",decf.format(hgds));
            resultmap.put("xhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultList.add(resultmap);
        }
        return resultList;

    }

    /**
     * 表4.1.5-1  标志检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getbzData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> bz = jjgFbgcJtaqssJabzService.lookJdbjg(commonInfoVo);
        if (bz!=null && bz.size()>0){
            Map resultmap = new LinkedHashMap<>();
            resultmap.put("htd",commonInfoVo.getHtd());
            double zds = 0.0;
            double hgds = 0.0;
            for (Map<String, Object> map : bz) {
                String xm = map.get("项目").toString();
                if (xm.contains("立柱竖直度")){
                    resultmap.put("szdzds",map.get("总点数"));
                    resultmap.put("szdhgds",map.get("合格点数"));
                    resultmap.put("szdhgl",map.get("合格率"));

                }
                if (xm.contains("标志板净空")){
                    resultmap.put("jkzds",map.get("总点数"));
                    resultmap.put("jkhgds",map.get("合格点数"));
                    resultmap.put("jkhgl",map.get("合格率"));

                }
                if (xm.contains("标志板厚度")){
                    resultmap.put("bzbhdzds",map.get("总点数"));
                    resultmap.put("bzbhdhgds",map.get("合格点数"));
                    resultmap.put("bzbhdhgl",map.get("合格率"));

                }
                if (xm.equals("标志面反光膜逆反射系数")){
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("xszds",decf.format(zds));
            resultmap.put("xshgds",decf.format(hgds));
            resultmap.put("xshgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultList.add(resultmap);

        }
        return resultList;

    }


    /**
     *表4.1.4-3  隧道工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdgcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        List<Map<String, Object>> list1 = getcqData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getztData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getsdlmData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     * 厚度待确认，自动化指标未编写
     * 表4.1.4-3隧道路面检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdlmData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        boolean a =false,b =false,c =false,d =false,e =false,f =false,aa= false,bb =false,cc = false,dd = false,ee = false,e2 = false;
        Map map = new LinkedHashMap();

        //沥青路面压实度
        List<Map<String, Object>> list1 = jjgFbgcSdgcSdlqlmysdService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            a = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list1) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("ysdjcds",decf.format(zds));
            map.put("ysdhgds",decf.format(hgds));
            map.put("ysdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //沥青路面车辙
        List<Map<String, Object>> list = jjgFbgcSdgcZdhczService.lookJdbjg(commonInfoVo);
        if (list!=null && list.size()>0){
            aa = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list) {
                zds += Double.valueOf(objectMap.get("总点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("czjcds",decf.format(zds));
            map.put("czhgds",decf.format(hgds));
            map.put("czhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //沥青路面渗水系数
        List<Map<String, Object>> list2 = jjgFbgcSdgcLmssxsService.lookJdbjg(commonInfoVo);
        if (list2!=null && list2.size()>0){
            b = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list2) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("xsjcds",decf.format(zds));
            map.put("xshgds",decf.format(hgds));
            map.put("xshgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //砼路面面强度
        List<Map<String, Object>> list3 = jjgFbgcSdgcHntlmqdService.lookJdbjg(commonInfoVo);
        if (list3!=null && list3.size()>0){
            c = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list3) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("qdjcds",decf.format(zds));
            map.put("qdhgds",decf.format(hgds));
            map.put("qdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //砼路面相邻板高差
        List<Map<String, Object>> list4 = jjgFbgcSdgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        if (list4!=null && list4.size()>0){
            d = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list4) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("gcjcds",decf.format(zds));
            map.put("gchgds",decf.format(hgds));
            map.put("gchgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //平整度
        List<Map<String, Object>> list10 = jjgFbgcSdgcZdhpzdService.lookJdbjg(commonInfoVo);
        if (list10!=null && list10.size()>0){
            bb = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list10) {
                zds += Double.valueOf(objectMap.get("总点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("pzdjcds",decf.format(zds));
            map.put("pzdhgds",decf.format(hgds));
            map.put("pzdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }
        //抗滑
        /*List<Map<String, Object>> list8 = jjgFbgcSdgcZdhgzsdService.lookJdbjg(commonInfoVo);
        double khzds = 0;
        double khhgds = 0;
        if (list8!=null && list8.size()>0) {
            cc = true;
            for (Map<String, Object> objectMap : list8) {
                khzds += Double.valueOf(objectMap.get("总点数").toString());
                khhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
        }
        List<Map<String, Object>> list9 = jjgFbgcSdgcZdhmcxsService.lookJdbjg(commonInfoVo);
        if (list9!=null && list9.size()>0) {
            dd = true;
            for (Map<String, Object> objectMap : list9) {
                khzds += Double.valueOf(objectMap.get("总点数").toString());
                khhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
        }
        if (cc || dd){
            map.put("khjcds",decf.format(khzds));
            map.put("khhgds",decf.format(khhgds));
            map.put("khhgl",khzds!=0 ? df.format(khhgds/khzds*100) : 0);
        }*/

        commonInfoVo.setFbgc("隧道工程");
        List<Map<String, Object>> mapList = jjgFbgcSdgcLmgzsdsgpsfService.lookJdbjgPage(commonInfoVo);
        if (mapList!=null && mapList.size()>0) {
            double khzds = 0;
            double khhgds = 0;
            for (Map<String, Object> objectMap : mapList) {
                khzds += Double.valueOf(objectMap.get("检测点数").toString());
                khhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("khjcds",decf.format(khzds));
            map.put("khhgds",decf.format(khhgds));
            map.put("khhgl",khzds!=0 ? df.format(khhgds/khzds*100) : 0);
        }
        commonInfoVo.setFbgc("");

        //厚度
        List<Map<String, Object>> list5 = jjgFbgcSdgcSdhntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list5!=null && list5.size()>0){
            double hzds = 0;
            double hhgds = 0;
            e = true;
            for (Map<String, Object> objectMap : list5) {
                hzds += Double.valueOf(objectMap.get("检测点数").toString());
                hhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("hnthdjcds",decf.format(hzds));
            map.put("hnthdhgds",decf.format(hhgds));
            map.put("hnthdhgl",hzds!=0 ? df.format(hhgds/hzds*100) : 0);
        }

        /*List<Map<String, Object>> list11 = jjgFbgcSdgcZdhldhdService.lookJdbjg(commonInfoVo);
        double hzds = 0;
        double hhgds = 0;
        if (list5!=null && list5.size()>0){
            e = true;
            for (Map<String, Object> objectMap : list5) {
                hzds += Double.valueOf(objectMap.get("检测点数").toString());
                hhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("hdjcds",decf.format(hzds));
            map.put("hdhgds",decf.format(hhgds));
            map.put("hdhgl",hzds!=0 ? df.format(hhgds/hzds*100) : 0);
        }*/
        List<Map<String, Object>> list6 = jjgFbgcSdgcGssdlqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list6!=null && list6.size()>0){
            double hzds = 0;
            double hhgds = 0;
            e2 = true;
            for (Map<String, Object> objectMap : list6) {
                hzds += Double.valueOf(objectMap.get("上面层厚度检测点数").toString());
                hzds += Double.valueOf(objectMap.get("总厚度检测点数").toString());
                hhgds += Double.valueOf(objectMap.get("上面层厚度合格点数").toString());
                hhgds += Double.valueOf(objectMap.get("总厚度合格点数").toString());
            }
            map.put("hdjcds",decf.format(hzds));
            map.put("hdhgds",decf.format(hhgds));
            map.put("hdhgl",hzds!=0 ? df.format(hhgds/hzds*100) : 0);

        }
        /*if (list11!=null && list11.size()>0){
            ee = true;
            for (Map<String, Object> objectMap : list11) {
                hzds += Double.valueOf(objectMap.get("总点数").toString());
                hhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
        }*/


        //横坡
        List<Map<String, Object>> list7 = jjgFbgcSdgcSdhpService.lookJdbjgPage(commonInfoVo);
        if (list7!=null && list7.size()>0){
            f = true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list7) {
                zds += Double.valueOf(objectMap.get("总点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("hpjcds",decf.format(zds));
            map.put("hphgds",decf.format(hgds));
            map.put("hphgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }
        if (a || b || c || d || e || f || aa || bb || cc || dd  || e2){
            map.put("htd",commonInfoVo.getHtd());
            resultList.add(map);
        }
        return resultList;

    }


    /**
     * 表4.1.4-2总体检测结果汇总表
     * 还有个净空
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getztData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        List<Map<String, Object>> list = jjgFbgcSdgcZtkdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> jk = jjgFbgcSdgcJkService.lookJdbjg(commonInfoVo);
        double kdzds = 0,kdhgds = 0;
        double jkzds = 0,jkhgds = 0;
        boolean a = false,b = false;
        Map map1 = new LinkedHashMap();
        if (list!=null && list.size()>0){
            a = true;
            for (Map<String, Object> map : list) {
                kdzds += Double.valueOf(map.get("总点数").toString());
                kdhgds += Double.valueOf(map.get("合格点数").toString());
            }

            map1.put("htd",commonInfoVo.getHtd());
            map1.put("kdjcds",kdzds);
            map1.put("kdhgds",kdhgds);
            map1.put("kdhgl",kdzds!=0 ? df.format(kdhgds/kdzds*100) : 0);
        }
        if (jk!=null && jk.size()>0){
            b =true;
            for (Map<String, Object> map : jk) {
                jkzds += Double.valueOf(map.get("总点数").toString());
                jkhgds += Double.valueOf(map.get("合格点数").toString());
            }
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("jkjcds",jkzds);
            map1.put("jkhgds",jkhgds);
            map1.put("jkhgl",jkzds!=0 ? df.format(jkhgds/jkzds*100) : 0);

        }
        if (a || b){
            resultList.add(map1);
        }

        return resultList;

    }

    /**
     * 表4.1.4-1 衬砌检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getcqData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //衬砌强度
        List<Map<String, Object>> cqqd = jjgFbgcSdgcCqtqdService.lookJdbjg(commonInfoVo);

        //衬砌厚度  多个
        List<Map<String, Object>> cqhd = jjgFbgcSdgcCqhdService.lookJdbjg(commonInfoVo);

        List<Map<String, Object>> dmpzd = jjgFbgcSdgcDmpzdService.lookJdbjg(commonInfoVo);

        Map remap = new HashMap();


        if (cqqd!=null && cqqd.size()>0){
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("cqqdjcds",cqqd.get(0).get("总点数"));
            map.put("cqqdhgds",cqqd.get(0).get("合格点数"));
            map.put("cqqdhgl",cqqd.get(0).get("合格率"));
            remap.putAll(map);
        }


        if (cqhd!=null && cqhd.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> stringObjectMap : cqhd) {
                zds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                hgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
            }
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("cqhdjcds",decf.format(zds));
            map.put("cqhdhgds",decf.format(hgds));
            map.put("cqhdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            remap.putAll(map);
        }
        if (dmpzd!=null && dmpzd.size()>0){
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("dmpzdjcds",dmpzd.get(0).get("总点数"));
            map.put("dmpzdhgds",dmpzd.get(0).get("合格点数"));
            map.put("dmpzdhgl",dmpzd.get(0).get("合格率"));
            remap.putAll(map);
        }
        resultList.add(remap);
        return resultList;

    }

    /**
     * 表4.1.3-3  桥梁工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getqlgcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        List<Map<String, Object>> list1 = getqlxbData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getqlsbData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getqmxData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getqmxData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        boolean a = false, b =false,c = false,d = false,e= false,f = false;
        Map resultmap = new LinkedHashMap();

        //平整度
        List<Map<String, Object>> list = jjgFbgcQlgcQmpzdService.lookJdbjg(commonInfoVo);
        //List<Map<String, Object>> list1 = jjgFbgcQlgcZdhpzdService.lookJdbjg(commonInfoVo);
        double pzdzds = 0;
        double pzdhgds = 0;
        if (CollectionUtils.isNotEmpty(list)){
            c = true;
            for (Map<String, Object> map : list) {
                pzdzds += Double.valueOf(map.get("检测点数").toString());
                pzdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            /*resultmap.put("qmpzdzds",decf.format(zds));
            resultmap.put("qmpzdhgds",decf.format(hgds));
            resultmap.put("qmpzdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);*/

        }
        /*if (CollectionUtils.isNotEmpty(list1)){
            f = true;
            for (Map<String, Object> map : list1) {
                pzdzds += Double.valueOf(map.get("总点数").toString());
                pzdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            *//*resultmap.put("qmpzdzds",decf.format(zds));
            resultmap.put("qmpzdhgds",decf.format(hgds));
            resultmap.put("qmpzdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);*//*

        }*/
        if (c || f) {
            resultmap.put("qmpzdzds",decf.format(pzdzds));
            resultmap.put("qmpzdhgds",decf.format(pzdhgds));
            resultmap.put("qmpzdhgl",pzdzds!=0 ? df.format(pzdhgds/pzdzds*100) : 0);
        }

        //横坡
        List<Map<String, Object>> list2 = jjgFbgcQlgcQmhpService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            double zds = 0;
            double hgds = 0;
            a = true;

            for (Map<String, Object> map : list2) {
                zds += Double.valueOf(map.get("总点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("qmhpzds",decf.format(zds));
            resultmap.put("qmhphgds",decf.format(hgds));
            resultmap.put("qmhphgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //抗滑
        List<Map<String, Object>> list3 = jjgFbgcQlgcQmgzsdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> list4 = jjgFbgcQlgcZdhgzsdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> list5 = jjgFbgcQlgcZdhmcxsService.lookJdbjg(commonInfoVo);
        double zdskh = 0;
        double hgdskh = 0;
        if (CollectionUtils.isNotEmpty(list3)){
            b = true;
            for (Map<String, Object> map : list3) {
                zdskh += Double.valueOf(map.get("检测点数").toString());
                hgdskh += Double.valueOf(map.get("合格点数").toString());
            }
        }
        if (b){
            resultmap.put("qmkhgzsdzds",decf.format(zdskh));
            resultmap.put("qmkhgzsdhgds",decf.format(hgdskh));
            resultmap.put("qmkhgzsdhgl",zdskh!=0 ? df.format(hgdskh/zdskh*100) : 0);
            resultmap.put("htd",commonInfoVo.getHtd());
            resultList.add(resultmap);
        }
        /*if (CollectionUtils.isNotEmpty(list4)){
            d = true;
            for (Map<String, Object> map : list4) {
                zdskh += Double.valueOf(map.get("总点数").toString());
                hgdskh += Double.valueOf(map.get("合格点数").toString());
            }
           *//* resultmap.put("qmkhgzsdzds",decf.format(zds));
            resultmap.put("qmkhgzsdhgds",decf.format(hgds));
            resultmap.put("qmkhgzsdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);*//*
        }
        if (CollectionUtils.isNotEmpty(list5)){
            e = true;
            for (Map<String, Object> map : list5) {
                zdskh += Double.valueOf(map.get("总点数").toString());
                hgdskh += Double.valueOf(map.get("合格点数").toString());
            }
        }
        if (b || d || e) {
            resultmap.put("qmkhgzsdzds",decf.format(zdskh));
            resultmap.put("qmkhgzsdhgds",decf.format(hgdskh));
            resultmap.put("qmkhgzsdhgl",zdskh!=0 ? df.format(hgdskh/zdskh*100) : 0);
        }

        if (a || b || c || d || e || f){
            resultmap.put("htd",commonInfoVo.getHtd());
            resultList.add(resultmap);

        }*/

        return resultList;

    }

    /**
     * 表4.1.3-2  桥梁上部检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqlsbData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();

        //墩台混凝土强度
        List<Map<String, Object>> list1 = jjgFbgcQlgcSbTqdService.lookJdbjg(commonInfoVo);

        //主要结构尺寸
        List<Map<String, Object>> list2 = jjgFbgcQlgcSbJgccService.lookJdbjg(commonInfoVo);

        //钢筋保护层厚度
        List<Map<String, Object>> list3 = jjgFbgcQlgcSbBhchdService.lookJdbjg(commonInfoVo);

        Map remap = new HashMap();

        if (list1!=null && list1.size()>0){
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("qlsbqdjcds",list1.get(0).get("总点数"));
            map.put("qlsbqdhgds",list1.get(0).get("合格点数"));
            map.put("qlsbqdhgl",list1.get(0).get("合格率"));
            remap.putAll(map);

        }
        if (list2!=null && list2.size()>0){
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("qlsbjgccjcds",list2.get(0).get("总点数"));
            map.put("qlsbjgcchgds",list2.get(0).get("合格点数"));
            map.put("qlsbjgcchgl",list2.get(0).get("合格率"));
            remap.putAll(map);
        }

        if (list3!=null && list3.size()>0){
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("qlsbbhchdjcds",list3.get(0).get("总点数"));
            map.put("qlsbbhchdhgds",list3.get(0).get("合格点数"));
            map.put("qlsbbhchdhgl",list3.get(0).get("合格率"));
            remap.putAll(map);
        }

        resultList.add(remap);
        return resultList;
    }

    /**
     * 表4.1.3-1  桥梁下部检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqlxbData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();

        //墩台混凝土强度
        List<Map<String, Object>> tqd = jjgFbgcQlgcXbTqdService.lookJdbjg(commonInfoVo);

        //主要结构尺寸
        List<Map<String, Object>> jgcc = jjgFbgcQlgcXbJgccService.lookJdbjg(commonInfoVo);

        //钢筋保护层厚度
        List<Map<String, Object>> bhchd = jjgFbgcQlgcXbBhchdService.lookJdbjg(commonInfoVo);

        //墩台垂直度
        List<Map<String, Object>> szd = jjgFbgcQlgcXbSzdService.lookJdbjg(commonInfoVo);

        Map remap = new HashMap();
        if (tqd!=null && tqd.size()>0){
            Map map = new LinkedHashMap();
            map.put("tqdjcds",tqd.get(0).get("总点数"));
            map.put("tqdhgds",tqd.get(0).get("合格点数"));
            map.put("tqdhgl",tqd.get(0).get("合格率"));
            map.put("htd",commonInfoVo.getHtd());
            remap.putAll(map);
        }

        if (jgcc!=null && jgcc.size()>0){
            Map map = new LinkedHashMap();
            map.put("jgccjcds",jgcc.get(0).get("总点数"));
            map.put("jgcchgds",jgcc.get(0).get("合格点数"));
            map.put("jgcchgl",jgcc.get(0).get("合格率"));
            map.put("htd",commonInfoVo.getHtd());
            remap.putAll(map);


        }

        if (bhchd!=null && bhchd.size()>0){
            Map map = new LinkedHashMap();
            map.put("bhchdjcds",bhchd.get(0).get("总点数"));
            map.put("bhchdhgds",bhchd.get(0).get("合格点数"));
            map.put("bhchdhgl",bhchd.get(0).get("合格率"));
            map.put("htd",commonInfoVo.getHtd());
            remap.putAll(map);
        }
        if (szd!=null && szd.size()>0){
            Map map = new LinkedHashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("szdjcds",szd.get(0).get("总点数"));
            map.put("szdhgds",szd.get(0).get("合格点数"));
            map.put("szdhgl",szd.get(0).get("合格率"));
            remap.putAll(map);

        }

        resultList.add(remap);

        return resultList;

    }


    /**
     * 表4.1.2-24  路面面层、桥面系、隧道路面抽查项目汇总表
     * 路面工程下所有桥梁和隧道 路面的数据汇总
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getxmccData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        resultmap.put("htd",commonInfoVo.getHtd());
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //压实度
        List<Map<String, Object>> list1 = jjgFbgcLmgcLqlmysdService.lookJdbjgbgzbg(commonInfoVo, 1);
        if (CollectionUtils.isNotEmpty(list1)){
            double ysdzds = 0;
            double ysdhgds = 0;
            for (Map<String, Object> map : list1) {
                ysdzds += Double.valueOf(map.get("检测点数").toString());
                ysdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("ysdzds",decf.format(ysdzds));
            resultmap.put("ysdhgds",decf.format(ysdhgds));
            resultmap.put("ysdhgl",ysdzds!=0 ? df.format(ysdhgds/ysdzds*100) : 0);
        }
        //路面弯沉
        List<Map<String, Object>> list2 = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo,1);
        List<Map<String, Object>> listlcf = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo, 1);
        double wczds = 0;
        double wchgds = 0;
        boolean a =false,b = false;
        if (CollectionUtils.isNotEmpty(list2)){
            a = true;
            for (Map<String, Object> map : list2) {
                wczds += Double.valueOf(map.get("检测单元数").toString());
                wchgds += Double.valueOf(map.get("合格单元数").toString());
            }

        }
        if (CollectionUtils.isNotEmpty(listlcf)){
            b = true;
            for (Map<String, Object> map : listlcf) {
                wczds += Double.valueOf(map.get("检测单元数").toString());
                wchgds += Double.valueOf(map.get("合格单元数").toString());
            }
        }
        if (a || b){
            resultmap.put("wczds",decf.format(wczds));
            resultmap.put("wchgds",decf.format(wchgds));
            resultmap.put("wchgl",wczds!=0 ? df.format(wchgds/wczds*100) : 0);
        }

        //车辙
        List<Map<String, Object>> list3 = jjgZdhCzService.lookJdbjg(commonInfoVo, 1);
        if (CollectionUtils.isNotEmpty(list3)){
            double czzds = 0;
            double czhgds = 0;
            for (Map<String, Object> map : list3) {
                String lmlx = map.get("路面类型").toString();
                if (!lmlx.contains("桥")){
                    czzds += Double.valueOf(map.get("总点数").toString());
                    czhgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("czzds",decf.format(czzds));
            resultmap.put("czhgds",decf.format(czhgds));
            resultmap.put("czhgl",czzds!=0 ? df.format(czhgds/czzds*100) : 0);
        }
        List<Map<String, Object>> list4 = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo,1);
        if (CollectionUtils.isNotEmpty(list4)){
            double ssxszds = 0;
            double ssxshgds = 0;
            for (Map<String, Object> map : list4) {
                ssxszds += Double.valueOf(map.get("检测点数").toString());
                ssxshgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("ssxszds",decf.format(ssxszds));
            resultmap.put("ssxshgds",decf.format(ssxshgds));
            resultmap.put("ssxshgl",ssxszds!=0 ? df.format(ssxshgds/ssxszds*100) : 0);
        }
        List<Map<String, Object>> list5 = jjgZdhPzdService.lookJdbjg(commonInfoVo, 1);
        if (CollectionUtils.isNotEmpty(list5)){
            double lqpzdzds = 0;
            double lqpzdhgds = 0;
            double hntlqpzdzds = 0;
            double hntlqpzdhgds = 0;
            for (Map<String, Object> map : list5) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("沥青")){
                    lqpzdzds += Double.valueOf(map.get("总点数").toString());
                    lqpzdhgds += Double.valueOf(map.get("合格点数").toString());
                }else if (lmlx.contains("混凝土")){
                    hntlqpzdzds += Double.valueOf(map.get("总点数").toString());
                    hntlqpzdhgds += Double.valueOf(map.get("合格点数").toString());

                }
            }
            resultmap.put("lqpzdzds",decf.format(lqpzdzds));
            resultmap.put("lqpzdhgds",decf.format(lqpzdhgds));
            resultmap.put("lqpzdhgl",lqpzdzds!=0 ? df.format(lqpzdhgds/lqpzdzds*100) : 0);
            resultmap.put("hntpzdzds",decf.format(hntlqpzdzds));
            resultmap.put("hntpzdhgds",decf.format(hntlqpzdhgds));
            resultmap.put("hntpzdhgl",hntlqpzdzds!=0 ? df.format(hntlqpzdhgds/hntlqpzdzds*100) : 0);
        }


        List<Map<String, Object>> list6 = jjgZdhMcxsService.lookJdbjg(commonInfoVo,1);
        if (CollectionUtils.isNotEmpty(list6)){
            double mcxszds = 0;
            double mcxshgds = 0;
            for (Map<String, Object> map : list6) {
                mcxszds += Double.valueOf(map.get("总点数").toString());
                mcxshgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("mcxszds",decf.format(mcxszds));
            resultmap.put("mcxshgds",decf.format(mcxshgds));
            resultmap.put("mcxshgl",mcxszds!=0 ? df.format(mcxshgds/mcxszds*100) : 0);
        }

        List<Map<String, Object>> list7 = jjgZdhGzsdService.lookJdbjg(commonInfoVo, 1);
        if (CollectionUtils.isNotEmpty(list7)){
            double gzsdzds = 0;
            double gzsdhgds = 0;
            for (Map<String, Object> map : list7) {
                gzsdzds += Double.valueOf(map.get("总点数").toString());
                gzsdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("gzsdzds",decf.format(gzsdzds));
            resultmap.put("gzsdhgds",decf.format(gzsdhgds));
            resultmap.put("gzsdhgl",gzsdzds!=0 ? df.format(gzsdhgds/gzsdzds*100) : 0);
        }

        //List<Map<String, Object>> list8 = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> ldhd = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        double lqhdzds = 0;
        double lqhdhgds = 0;
        boolean aa =false,bb = false;
        /*if (CollectionUtils.isNotEmpty(list8)){
            aa = true;
            for (Map<String, Object> map : list8) {
                lqhdzds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                lqhdzds += Double.valueOf(map.get("总厚度检测点数").toString());
                lqhdhgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                lqhdhgds += Double.valueOf(map.get("总厚度合格点数").toString());
            }

        }*/
        if (CollectionUtils.isNotEmpty(ldhd)){
            bb=true;
            for (Map<String, Object> map : ldhd) {
                String lmlx = map.get("路面类型").toString();
                if (!lmlx.contains("桥")){
                    lqhdzds += Double.valueOf(map.get("总点数").toString());
                    lqhdhgds += Double.valueOf(map.get("合格点数").toString());
                }

            }
        }
        if (aa || bb){
            resultmap.put("lqhdzds",decf.format(lqhdzds));
            resultmap.put("lqhdhgds",decf.format(lqhdhgds));
            resultmap.put("lqhdhgl",lqhdzds!=0 ? df.format(lqhdhgds/lqhdzds*100) : 0);
        }

        List<Map<String, String>> list9 = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list9)){
            double lqhpzds = 0;
            double lqhphgds = 0;
            for (Map<String, String> map : list9) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("沥青")){
                    lqhpzds += Double.valueOf(map.get("检测点数"));
                    lqhphgds += Double.valueOf(map.get("合格点数"));
                }
            }
            resultmap.put("lqhpzds",decf.format(lqhpzds));
            resultmap.put("lqhphgds",decf.format(lqhphgds));
            resultmap.put("lqhphgl",lqhpzds!=0 ? df.format(lqhphgds/lqhpzds*100) : 0);
        }

        List<Map<String, Object>> list10 = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list10)){
            double hntqdzds = 0;
            double hntqdhgds = 0;
            for (Map<String, Object> map : list10) {
                hntqdzds += Double.valueOf(map.get("总点数").toString());
                hntqdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("hntqdzds",decf.format(hntqdzds));
            resultmap.put("hntqdhgds",decf.format(hntqdhgds));
            resultmap.put("hntqdhgl",hntqdzds!=0 ? df.format(hntqdhgds/hntqdzds*100) : 0);
        }
        List<Map<String, Object>> list11 = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list11)){
            double tlmxlbgczds = 0;
            double tlmxlbgchgds = 0;
            for (Map<String, Object> map : list11) {
                tlmxlbgczds += Double.valueOf(map.get("总点数").toString());
                tlmxlbgchgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("tlmxlbgczds",decf.format(tlmxlbgczds));
            resultmap.put("tlmxlbgchgds",decf.format(tlmxlbgchgds));
            resultmap.put("tlmxlbgchgl",tlmxlbgczds!=0 ? df.format(tlmxlbgchgds/tlmxlbgczds*100) : 0);
        }
        //混凝土路面平整度
        /*if (CollectionUtils.isNotEmpty(list5)){
            double hntpzdzds = 0;
            double hntpzdhgds = 0;
            for (Map<String, Object> map : list5) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("混凝土路面")){
                    hntpzdzds += Double.valueOf(map.get("总点数").toString());
                    hntpzdhgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("hntpzdzds",decf.format(hntpzdzds));
            resultmap.put("hntpzdhgds",decf.format(hntpzdhgds));
            resultmap.put("hntpzdhgl",hntpzdzds!=0 ? df.format(hntpzdhgds/hntpzdzds*100) : 0);
        }*/

        //抗滑
        List<Map<String, Object>> list12 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (CollectionUtils.isNotEmpty(list12)){
            double khzds = 0;
            double khhgds = 0;
            for (Map<String, Object> map : list12) {
                khzds += Double.valueOf(map.get("检测点数").toString());
                khhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("khzds",decf.format(khzds));
            resultmap.put("khhgds",decf.format(khhgds));
            resultmap.put("khhgl",khzds!=0 ? df.format(khhgds/khzds*100) : 0);
        }

        List<Map<String, Object>> list13 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list13)){
            double hnthdzds = 0;
            double hnthgds = 0;
            for (Map<String, Object> map : list13) {
                hnthdzds += Double.valueOf(map.get("检测点数").toString());
                hnthgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("hnthdzds",decf.format(hnthdzds));
            resultmap.put("hnthgds",decf.format(hnthgds));
            resultmap.put("hnthgl",hnthdzds!=0 ? df.format(hnthgds/hnthdzds*100) : 0);
        }

        if (CollectionUtils.isNotEmpty(list9)){
            double hnthpzds = 0;
            double hnthphgds = 0;
            for (Map<String, String> map : list9) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("混凝土路面")){
                    hnthpzds += Double.valueOf(map.get("检测点数"));
                    hnthphgds += Double.valueOf(map.get("合格点数"));
                }
            }
            resultmap.put("hnthpzds",decf.format(hnthpzds));
            resultmap.put("hnthphgds",decf.format(hnthphgds));
            resultmap.put("hnthphdl",hnthpzds!=0 ? df.format(hnthphgds/hnthpzds*100) : 0);
        }

        resultList.add(resultmap);
        return resultList;
    }

    /**
     *
     * 表4.1.2-23 隧道路面工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdlmhzData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap();
        //List<Map<String, Object>> list1 = getsdysdData(commonInfoVo);
        List<Map<String, Object>> list1 = jjgFbgcLmgcLqlmysdService.lookJdbjgbgzbg(commonInfoVo, 1);

        if (CollectionUtils.isNotEmpty(list1)){
            double zds = 0,hgds = 0;
            for (Map<String, Object> map : list1) {
                String lmlc = map.get("路面类型").toString();
                if (lmlc.contains("隧道")){
                    zds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
                //resultmap.putAll(map);
            }
            Map map = new HashMap<>();
            map.put("htd", commonInfoVo.getHtd());
            map.put("ysdzs",decf.format(zds));
            map.put("ysdhgs",decf.format(hgds));
            map.put("ysdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultmap.putAll(map);

        }
        List<Map<String, Object>> list2 = getsdczData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            double zds = 0,hgds = 0;
            for (Map<String, Object> map : list2) {
                zds += Double.valueOf(map.get("sdczjcds").toString());
                hgds += Double.valueOf(map.get("sdczhgds").toString());
                //resultmap.putAll(map);
            }
            Map map = new HashMap<>();
            map.put("htd", commonInfoVo.getHtd());
            map.put("sdczjcds",decf.format(zds));
            map.put("sdczhgds",decf.format(hgds));
            map.put("sdczhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultmap.putAll(map);
        }
        List<Map<String, Object>> list3 = getsdssxsData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list4 = getsdhntqdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            for (Map<String, Object> map : list4) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list5 = getsdhntxlbgcData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list5)){
            for (Map<String, Object> map : list5) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list6 = getsdpzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list6)){
            double zds = 0,hgds = 0;
            for (Map<String, Object> map : list6) {
                String sdpzdlmlx = map.get("sdpzdlmlx").toString();
                if (sdpzdlmlx.contains("混凝土")){
                    resultmap.put("sdpzdhntjcds",map.get("sdpzdjcds"));
                    resultmap.put("sdpzdhnthgds",map.get("sdpzdhgds"));
                    resultmap.put("sdpzdhnthgl",map.get("sdpzdhgl"));
                }else {
                    zds += Double.valueOf(map.get("sdpzdjcds").toString());
                    hgds += Double.valueOf(map.get("sdpzdhgds").toString());

                    /*resultmap.put("sdpzdlqjcds",map.get("sdpzdjcds"));
                    resultmap.put("sdpzdlqhgds",map.get("sdpzdhgds"));
                    resultmap.put("sdpzdlqhgl",map.get("sdpzdhgl"));*/
                }
            }
            resultmap.put("sdpzdlqjcds",decf.format(zds));
            resultmap.put("sdpzdlqhgds",decf.format(hgds));
            resultmap.put("sdpzdlqhgl",zds!=0 ? df.format(hgds/zds*100) : 0);

        }

        List<Map<String, Object>> list7 = getsdkhData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list7)){
            double gzsdzds = 0,gzsdhgds = 0;
            double mcxszds = 0,mcxshgds = 0;
            for (Map<String, Object> map : list7) {
                String hplmlx = map.get("sdkhlmlx").toString();
                if (hplmlx.contains("隧道路面")){
                    String khkpzb = map.get("sdkhzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        mcxszds += Double.valueOf(map.get("sdkhjcds").toString());
                        mcxshgds += Double.valueOf(map.get("sdkhhgds").toString());

                        /*resultmap.put("khlqmcxsjcds",map.get("sdkhjcds"));
                        resultmap.put("khlqmcxshgds",map.get("sdkhhgds"));
                        resultmap.put("khlqmcxshgl",map.get("sdkhhgl"));*/
                    }else if (khkpzb.equals("构造深度")){
                        gzsdzds += Double.valueOf(map.get("sdkhjcds").toString());
                        gzsdhgds += Double.valueOf(map.get("sdkhhgds").toString());

                        /*resultmap.put("khlqgzsdjcds",map.get("sdkhjcds"));
                        resultmap.put("khlqgzsdhgds",map.get("sdkhhgds"));
                        resultmap.put("khlqgzsdhgl",map.get("sdkhhgl"));*/
                    }
                }else if (hplmlx.contains("混凝土隧道")){
                    String khkpzb = map.get("sdkhzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        resultmap.put("khhntmcxsjcds",map.get("sdkhjcds"));
                        resultmap.put("khhntmcxshgds",map.get("sdkhhgds"));
                        resultmap.put("khhntmcxshgl",map.get("sdkhhgl"));
                    }else if (khkpzb.equals("构造深度")){
                        resultmap.put("khhntgzsdjcds",map.get("sdkhjcds"));
                        resultmap.put("khhntgzsdhgds",map.get("sdkhhgds"));
                        resultmap.put("khhntgzsdhgl",map.get("sdkhhgl"));
                    }

                }
            }

            resultmap.put("khlqmcxsjcds",df.format(mcxszds));
            resultmap.put("khlqmcxshgds",df.format(mcxshgds));
            resultmap.put("khlqmcxshgl",mcxszds!=0 ? df.format(mcxshgds/mcxszds*100) : 0);

            resultmap.put("khlqgzsdjcds",df.format(gzsdzds));
            resultmap.put("khlqgzsdhgds",df.format(gzsdhgds));
            resultmap.put("khlqgzsdhgl",gzsdzds!=0 ? df.format(gzsdhgds/gzsdzds*100) : 0);

        }

        List<Map<String, Object>> list10 = getsdhdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list10)){
            for (Map<String, Object> map : list10) {
                String hplmlx = map.get("sdlmlx").toString();
                if (hplmlx.contains("沥青路面")){
                    resultmap.put("hdlqzds",map.get("sdjcds"));
                    resultmap.put("hdlqhgds",map.get("sdhgs"));
                    resultmap.put("hdlqhgl",map.get("sdhgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("hdhntzds",map.get("sdjcds"));
                    resultmap.put("hdhnthgds",map.get("sdhgs"));
                    resultmap.put("hdhnthgl",map.get("sdhgl"));
                }
            }
        }
        List<Map<String, Object>> list11 = getsdhpData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list11)){
            for (Map<String, Object> map : list11) {
                String hplmlx = map.get("sdhplmlx").toString();
                if (hplmlx.contains("沥青")){
                    resultmap.put("hplqzds",map.get("sdhpzds"));
                    resultmap.put("hplqhgds",map.get("sdhphgds"));
                    resultmap.put("hplqhgl",map.get("sdhphgl"));
                }else if (hplmlx.equals("混凝土")){
                    resultmap.put("hphntzds",map.get("sdhpzds"));
                    resultmap.put("hphnthgds",map.get("sdhphgds"));
                    resultmap.put("hphnthgl",map.get("sdhphgl"));
                }
            }

        }
        resultList.add(resultmap);
        return resultList;

    }



    /**
     * 表4.1.2-22 隧道路面横坡检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhpData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, String>> list = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        if (list!=null && list.size()>0){
            double lmzds = 0.0;
            double lmhgds = 0.0;

            double lmzds1 = 0.0;
            double lmhgds1 = 0.0;
            boolean a = false;
            boolean b = false;
            for (Map<String, String> map : list) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("沥青隧道")){
                    a = true;
                    lmzds +=Double.valueOf(map.get("检测点数"));
                    lmhgds +=Double.valueOf(map.get("合格点数"));

                }
                if (lmlx.contains("混凝土隧道")){
                    b = true;
                    lmzds1 +=Double.valueOf(map.get("检测点数"));
                    lmhgds1 +=Double.valueOf(map.get("合格点数"));
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdhpzds",decf.format(lmzds));
                map1.put("sdhphgds",decf.format(lmhgds));
                map1.put("sdhphgl",lmzds!=0 ? df.format(lmhgds/lmzds*100) : 0);
                map1.put("sdhplmlx","沥青隧道");
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdhpzds",decf.format(lmzds1));
                map2.put("sdhphgds",decf.format(lmhgds1));
                map2.put("sdhphgl",lmzds1!=0 ? df.format(lmhgds1/lmzds1*100) : 0);
                map2.put("sdhplmlx","混凝土隧道");
                resultList.add(map2);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-21（3）  隧道路面厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青

        double smcjcds = 0,smchgds = 0;
        double zhdjcds = 0,zhdhgds = 0;

        double jcds = 0,hgds = 0;
        double ljxjcds = 0,ljxhgds = 0;
        boolean a = false,b = false;

        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道") ){
                    a = true;
                    smcjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    zhdjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    smchgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    zhdhgds += Double.valueOf(map.get("总厚度合格点数").toString());
                }

            }
            if (a){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdlmlx","沥青路面上面层厚度（高速沥青厚度钻芯法-隧道）");
                map2.put("sdjcds",decf.format(smcjcds));
                map2.put("sdhgs",decf.format(smchgds));
                map2.put("sdhgl",smcjcds!=0 ? df.format(smchgds/smcjcds*100) : 0);
                resultList.add(map2);
                Map map3= new HashMap();
                map3.put("htd",commonInfoVo.getHtd());
                map3.put("sdlmlx","沥青路面总厚度（高速沥青厚度钻芯法-隧道）");
                map3.put("sdjcds",decf.format(zhdjcds));
                map3.put("sdhgs",decf.format(zhdhgds));
                map3.put("sdhgl",zhdjcds!=0 ? df.format(zhdhgds/zhdjcds*100) : 0);
                resultList.add(map3);
            }
        }
        //混凝土  无隧道

        //自动化的雷达法有
        List<Map<String, Object>> list1 = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            for (Map<String, Object> map : list1) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线") && lmlx.contains("隧道")){
                    b = true;
                    ljxjcds += Double.valueOf(map.get("总点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());

                }else {
                    if (lmlx.contains("隧道") ){
                        a = true;
                        jcds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }

            }
        }
        if (a){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("sdlmlx","沥青路面（路面雷达厚度-隧道）");
            map2.put("sdjcds",decf.format(jcds));
            map2.put("sdhgs",decf.format(hgds));
            map2.put("sdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            resultList.add(map2);
        }
        if (b){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("sdlmlx","沥青路面（路面雷达厚度-连接线隧道）");
            map2.put("sdjcds",decf.format(ljxjcds));
            map2.put("sdhgs",decf.format(ljxhgds));
            map2.put("sdhgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
            resultList.add(map2);
        }

        return resultList;

    }

    /**
     * 表4.1.2-21（2）  隧道路面雷达厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdldhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> resultList = new ArrayList<>();
        Double sdMax = 0.0,ljxsdMax = 0.0;
        Double sdMin = Double.MAX_VALUE,ljxsdMin = Double.MAX_VALUE;
        double jcds = 0,ljxjcds = 0;
        double hgs = 0,ljxhgs = 0;
        String ljxsjz = "";
        if (list!=null && list.size()>0){
            boolean a = false,b =false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("隧道")){
                        b = true;
                        double max = Double.valueOf(map.get("Max").toString());
                        ljxsdMax = (max > ljxsdMax) ? max : ljxsdMax;
                        double min = Double.valueOf(map.get("Min").toString());
                        ljxsdMin = (min < ljxsdMin) ? min : ljxsdMin;

                        ljxjcds += Double.valueOf(map.get("总点数").toString());
                        ljxhgs += Double.valueOf(map.get("合格点数").toString());

                        ljxsjz = map.get("设计值").toString();
                    }

                }else {
                    if (lmlx.contains("隧道")){
                        a = true;
                        double max = Double.valueOf(map.get("Max").toString());
                        sdMax = (max > sdMax) ? max : sdMax;
                        double min = Double.valueOf(map.get("Min").toString());
                        sdMin = (min < sdMin) ? min : sdMin;

                        jcds += Double.valueOf(map.get("总点数").toString());
                        hgs += Double.valueOf(map.get("合格点数").toString());

                    }
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdldhdlb","总厚度");
                map1.put("sdldhdlmlx","沥青路面（主线/匝道）");
                map1.put("sdldhdsjz",list.get(0).get("设计值"));
                map1.put("sdldhdjcds",decf.format(jcds));
                map1.put("sdldhdhgds",decf.format(hgs));
                map1.put("sdldhdhgl",jcds!=0 ? df.format(hgs/jcds*100) : 0);
                map1.put("sdldhdbhfw",df.format(sdMin)+"~"+df.format(sdMax));
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdldhdlb","总厚度");
                map1.put("sdldhdlmlx","沥青路面（连接线）");
                map1.put("sdldhdsjz",ljxsjz);
                map1.put("sdldhdjcds",decf.format(ljxjcds));
                map1.put("sdldhdhgds",decf.format(ljxhgs));
                map1.put("sdldhdhgl",ljxjcds!=0 ? df.format(ljxhgs/ljxjcds*100) : 0);
                map1.put("sdldhdbhfw",df.format(ljxsdMin)+"~"+df.format(ljxsdMax));
                resultList.add(map1);
            }
        }

        return resultList;

    }

    /**
     * 表4.1.2-21（1）  隧道路面钻芯法厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdzxfhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list!=null && list.size()>0){
            List<Map<String, Object>> sdsmc = new ArrayList<>();
            List<Map<String, Object>> sdzhd = new ArrayList<>();
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    Map map3 = new HashMap();
                    Map map4 = new HashMap();
                    map3.put("上面层厚度合格点数", map.get("上面层厚度合格点数").toString());
                    map3.put("上面层平均值最小值", map.get("上面层平均值最小值").toString());
                    map3.put("上面层设计值", map.get("上面层设计值").toString());
                    map3.put("上面层代表值", map.get("上面层代表值").toString());
                    map3.put("上面层厚度合格率", map.get("上面层厚度合格率").toString());
                    map3.put("上面层平均值最大值", map.get("上面层平均值最大值").toString());
                    map3.put("上面层厚度检测点数", map.get("上面层厚度检测点数").toString());
                    map3.put("路面类型", map.get("路面类型").toString());
                    sdsmc.add(map3);

                    map4.put("总厚度合格点数", map.get("总厚度合格点数").toString());
                    map4.put("总厚度平均值最小值", map.get("总厚度平均值最小值").toString());
                    map4.put("总厚度检测点数", map.get("总厚度检测点数").toString());
                    map4.put("总厚度代表值", map.get("总厚度代表值").toString());
                    map4.put("总厚度平均值最大值", map.get("总厚度平均值最大值").toString());
                    map4.put("总厚度设计值", map.get("总厚度设计值").toString());
                    map4.put("总厚度合格率", map.get("总厚度合格率").toString());
                    map4.put("路面类型", map.get("路面类型").toString());
                    sdzhd.add(map4);

                }
            }
            Map<String, List<Map<String, Object>>> sdsmcgrouped =
                    sdsmc.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("上面层设计值").toString()
                            ));
            Map<String, List<Map<String, Object>>> sdzhdgrouped =
                    sdzhd.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("总厚度设计值").toString()
                            ));
            for (List<Map<String, Object>> slist : sdsmcgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : slist) {
                    zds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("上面层厚度合格点数").toString());

                    double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("上面层代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("上面层代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("sdzxfhdlmlx","沥青路面（隧道）");
                mapzhd.put("sdzxfhdlb","上面层厚度");
                mapzhd.put("sdzxfhdjcds", zds);
                mapzhd.put("sdzxfhdhgs", hgds);
                mapzhd.put("sdzxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("sdzxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("sdzxfhdsjz", slist.get(0).get("上面层设计值"));
                mapzhd.put("sdzxfhdzdbz", zzhdbz);
                mapzhd.put("sdzxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }
            for (List<Map<String, Object>> zlist : sdzhdgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : zlist) {
                    zds += Double.valueOf(map.get("总厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("总厚度合格点数").toString());

                    double max = Double.valueOf(map.get("总厚度平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("总厚度平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("总厚度代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("总厚度代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("sdzxfhdlmlx","沥青路面（隧道）");
                mapzhd.put("sdzxfhdlb","总厚度");
                mapzhd.put("sdzxfhdjcds", decf.format(zds));
                mapzhd.put("sdzxfhdhgs", decf.format(hgds));
                mapzhd.put("sdzxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("sdzxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("sdzxfhdsjz", zlist.get(0).get("总厚度设计值"));
                mapzhd.put("sdzxfhdzdbz", zzhdbz);
                mapzhd.put("sdzxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }
            /*Double sMax = 0.0;
            Double sMin = Double.MAX_VALUE;
            Double zMax = 0.0;
            Double zMin = Double.MAX_VALUE;
            String zdbz = "";
            String ydbz = "";
            String zzdbz = "";
            String yzdbz = "";
            String smcsjz = "";
            String zhdsjz = "";
            double smcjcds = 0;
            double smchgs = 0;
            double zhdjcds = 0;
            double zhdhgs = 0;
            boolean a = false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("隧道左幅") || lmlx.equals("隧道右幅")){
                    a = true;
                    double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;

                    double max1 = Double.valueOf(map.get("总厚度平均值最大值").toString());
                    zMax = (max1 > zMax) ? max1 : zMax;

                    double min1 = Double.valueOf(map.get("总厚度平均值最小值").toString());
                    zMin = (min1 < zMin) ? min1 : zMin;

                    if (lmlx.equals("隧道左幅")){
                        zdbz = map.get("上面层代表值").toString();
                        zzdbz = map.get("总厚度代表值").toString();
                    }
                    if (lmlx.equals("隧道右幅")){
                        ydbz = map.get("上面层代表值").toString();
                        yzdbz = map.get("总厚度代表值").toString();
                    }
                    smcsjz = map.get("上面层设计值").toString();
                    zhdsjz = map.get("总厚度设计值").toString();
                    smcjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    smchgs += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    zhdjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    zhdhgs += Double.valueOf(map.get("总厚度合格点数").toString());

                }
            }
            if (a){
                Map map = new HashMap();
                map.put("htd",commonInfoVo.getHtd());
                map.put("sdzxfhdlmlx","沥青路面（隧道）");
                map.put("sdzxfhdlb","上面层厚度");
                map.put("sdzxfhdpjzfw",sMin+"~"+sMax);
                map.put("sdzxfhdzdbz",zdbz);
                map.put("sdzxfhdydbz",ydbz);
                map.put("sdzxfhdsjz",smcsjz);
                map.put("sdzxfhdjcds",decf.format(smcjcds));
                map.put("sdzxfhdhgs",decf.format(smchgs));
                map.put("sdzxfhdhgl",smcjcds!=0 ? df.format(smchgs/smcjcds*100) : 0);
                resultList.add(map);
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdzxfhdlmlx","沥青路面（隧道）");
                map1.put("sdzxfhdlb","总厚度");
                map1.put("sdzxfhdpjzfw",zMin+"~"+zMax);
                map1.put("sdzxfhdzdbz",zzdbz);
                map1.put("sdzxfhdydbz",yzdbz);
                map1.put("sdzxfhdsjz",zhdsjz);
                map1.put("sdzxfhdjcds",decf.format(zhdjcds));
                map1.put("sdzxfhdhgs",decf.format(zhdhgs));
                map1.put("sdzxfhdhgl",zhdjcds!=0 ? df.format(zhdhgs/zhdjcds*100) : 0);
                resultList.add(map1);
            }*/



        }
        //混凝土
        /*List<Map<String, Object>> list1 = jjgFbgcSdgcSdhntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            double jcds = 0;
            double hgds = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    a = true;
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;
                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            if (a){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdzxfhdlmlx","混凝土路面");
                map2.put("sdzxfhdlb","总厚度");
                map2.put("sdzxfhdpjzfw",lmMin+"~"+lmMax);
                map2.put("sdzxfhdzdbz",list1.get(0).get("代表值"));
                map2.put("sdzxfhdydbz",list1.get(0).get("代表值"));
                map2.put("sdzxfhdsjz",list1.get(0).get("设计值"));
                map2.put("sdzxfhdjcds",decf.format(jcds));
                map2.put("sdzxfhdhgs",decf.format(hgds));
                map2.put("sdzxfhdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map2);
            }
        }*/

        return resultList;


    }

    /**
     * 表4.1.2-20  隧道路面抗滑检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdkhData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhGzsdService.lookJdbjg(commonInfoVo, 1);//自动化
        String sjz = "",ljxsjz = "";
        double jcds = 0.0;
        double hgds = 0.0;
        double ljxjcds = 0.0;
        double ljxhgds = 0.0;

        if (list!=null && list.size()>0){
            boolean a = false,b = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("隧道")){
                        b = true;
                        ljxsjz = map.get("设计值").toString();
                        ljxjcds += Double.valueOf(map.get("总点数").toString());
                        ljxhgds += Double.valueOf(map.get("合格点数").toString());

                        double max = Double.valueOf(map.get("最大值").toString());
                        ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                        double min = Double.valueOf(map.get("最小值").toString());
                        ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;

                    }

                }else {
                    if (lmlx.contains("隧道")){
                        a = true;
                        sjz = map.get("设计值").toString();
                        jcds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                        double max = Double.valueOf(map.get("最大值").toString());
                        lmMax = (max > lmMax) ? max : lmMax;

                        double min = Double.valueOf(map.get("最小值").toString());
                        lmMin = (min < lmMin) ? min : lmMin;
                    }
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","隧道路面（主线/匝道）");
                map1.put("sdkhzb","构造深度");
                map1.put("sdkhsjz",sjz);
                map1.put("sdkhbhfw",lmMin+"~"+lmMax);
                map1.put("sdkhjcds",decf.format(jcds));
                map1.put("sdkhhgds",decf.format(hgds));
                map1.put("sdkhhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","隧道路面（连接线）");
                map1.put("sdkhzb","构造深度");
                map1.put("sdkhsjz",ljxsjz);
                map1.put("sdkhbhfw",ljxlmMin+"~"+ljxlmMax);
                map1.put("sdkhjcds",decf.format(ljxjcds));
                map1.put("sdkhhgds",decf.format(ljxhgds));
                map1.put("sdkhhgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list1 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (list1!=null && list1.size()>0){
            String gdz = "";
            double szds = 0;
            double shgs = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("混凝土隧道")){
                    a = true;
                    szds += Double.valueOf(map.get("检测点数").toString());
                    shgs += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                }else {
                    continue;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","混凝土隧道");
                map1.put("sdkhzb","构造深度");
                map1.put("sdkhsjz",gdz);
                map1.put("sdkhbhfw",lmMin+"~"+lmMax);
                map1.put("sdkhjcds",decf.format(szds));
                map1.put("sdkhhgds",decf.format(shgs));
                map1.put("sdkhhgl",szds!=0 ? df.format(shgs/szds*100) : 0);
                resultList.add(map1);
            }

        }
        List<Map<String, Object>> list2 = jjgZdhMcxsService.lookJdbjg(commonInfoVo,1);
        if (list2!=null && list2.size()>0){
            boolean a = false,b = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            double zds = 0;
            double hgs = 0;
            double ljxzds = 0;
            double ljxhgs = 0;
            String sjz2 = "",mcxsljxsjz = "";
            for (Map<String, Object> map : list2) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("隧道")){
                        b = true;
                        mcxsljxsjz = map.get("设计值").toString();
                        ljxzds += Double.valueOf(map.get("总点数").toString());
                        ljxhgs += Double.valueOf(map.get("合格点数").toString());

                        double max = Double.valueOf(map.get("Max").toString());
                        ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                        double min = Double.valueOf(map.get("Min").toString());
                        ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;

                    }
                }else {
                    if (lmlx.contains("隧道")){
                        a = true;
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgs += Double.valueOf(map.get("合格点数").toString());
                        sjz2 = map.get("设计值").toString();

                        double max = Double.valueOf(map.get("Max").toString());
                        lmMax = (max > lmMax) ? max : lmMax;

                        double min = Double.valueOf(map.get("Min").toString());
                        lmMin = (min < lmMin) ? min : lmMin;
                    }
                }


            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","隧道路面（主线/匝道）");
                map1.put("sdkhzb","摩擦系数");
                map1.put("sdkhsjz",sjz2);
                map1.put("sdkhbhfw",lmMin+"~"+lmMax);
                map1.put("sdkhjcds",decf.format(zds));
                map1.put("sdkhhgds",decf.format(hgs));
                map1.put("sdkhhgl",zds!=0 ? df.format(hgs/zds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","隧道路面（连接线）");
                map1.put("sdkhzb","摩擦系数");
                map1.put("sdkhsjz",mcxsljxsjz);
                map1.put("sdkhbhfw",df.format(ljxlmMin)+"~"+df.format(ljxlmMax));
                map1.put("sdkhjcds",decf.format(ljxzds));
                map1.put("sdkhhgds",decf.format(ljxhgs));
                map1.put("sdkhhgl",ljxzds!=0 ? df.format(ljxhgs/ljxzds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;


    }

    /**
     * 表4.1.2-19  隧道路面平整度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdpzdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.0");
        DecimalFormat decf = new DecimalFormat("0.#");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhPzdService.lookJdbjg(commonInfoVo, 1);
        if (list!=null && list.size()>0){
            boolean a = false,b =false,c = false;
            String ljxsdlmlx = "",ljxsdsjz = "";
            String sdlmlx = "",sdsjz = "";
            String hntsdlmlx = "",hntsdsjz = "";
            double ljxsdzds = 0,ljxsdhgds = 0;
            double sdzds = 0,sdhgds = 0;
            double hntsdzds = 0,hntsdhgds = 0;
            List<Double> ljxsdbhfw = new ArrayList<>();
            List<Double> sdbhfw = new ArrayList<>();
            List<Double> hntsdbhfw = new ArrayList<>();

            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("隧道")){
                        a =true;
                        ljxsdlmlx = lmlx;
                        ljxsdsjz = map.get("设计值").toString();
                        ljxsdzds += Double.valueOf(map.get("总点数").toString());
                        ljxsdhgds += Double.valueOf(map.get("合格点数").toString());
                        ljxsdbhfw.add(Double.valueOf(map.get("Min").toString()));
                        ljxsdbhfw.add(Double.valueOf(map.get("Max").toString()));

                        /*Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("sdpzdzb","IRI");
                        map1.put("lmlx",lmlx);
                        map1.put("sdpzdlmlx","沥青路面（连接线隧道）");
                        map1.put("sdpzdgdz",map.get("设计值"));
                        map1.put("sdpzdjcds",map.get("总点数"));
                        map1.put("sdpzdhgds",map.get("合格点数"));
                        map1.put("sdpzdhgl",map.get("合格率"));
                        map1.put("sdpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);*/
                    }

                }else {
                    if (lmlx.contains("隧道")){
                        b = true;
                        sdlmlx = lmlx;
                        sdsjz = map.get("设计值").toString();
                        sdzds += Double.valueOf(map.get("总点数").toString());
                        sdhgds += Double.valueOf(map.get("合格点数").toString());
                        sdbhfw.add(Double.valueOf(map.get("Min").toString()));
                        sdbhfw.add(Double.valueOf(map.get("Max").toString()));

                        /*Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("sdpzdzb","IRI");
                        map1.put("sdpzdlmlx","沥青路面（隧道）");
                        map1.put("sdpzdgdz",map.get("设计值"));
                        map1.put("sdpzdjcds",map.get("总点数"));
                        map1.put("sdpzdhgds",map.get("合格点数"));
                        map1.put("sdpzdhgl",map.get("合格率"));
                        map1.put("sdpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);*/
                    }
                    if (lmlx.equals("混凝土隧道")){
                        c = true;
                        hntsdlmlx = lmlx;
                        hntsdsjz = map.get("设计值").toString();
                        hntsdzds += Double.valueOf(map.get("总点数").toString());
                        hntsdhgds += Double.valueOf(map.get("合格点数").toString());
                        hntsdbhfw.add(Double.valueOf(map.get("Min").toString()));
                        hntsdbhfw.add(Double.valueOf(map.get("Max").toString()));

                        /*Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("sdpzdzb","IRI");
                        map1.put("sdpzdlmlx","混凝土路面（隧道）");
                        map1.put("sdpzdgdz",map.get("设计值"));
                        map1.put("sdpzdjcds",map.get("总点数"));
                        map1.put("sdpzdhgds",map.get("合格点数"));
                        map1.put("sdpzdhgl",map.get("合格率"));
                        map1.put("sdpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);*/
                    }
                }
            }

            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdpzdzb","IRI");
                map1.put("lmlx",ljxsdlmlx);
                map1.put("sdpzdlmlx","隧道路面（连接线隧道）");
                map1.put("sdpzdgdz",ljxsdsjz);
                map1.put("sdpzdjcds",decf.format(ljxsdzds));
                map1.put("sdpzdhgds",decf.format(ljxsdhgds));
                map1.put("sdpzdhgl",ljxsdzds!=0 ? df.format(ljxsdhgds/ljxsdzds*100) : 0);
                map1.put("sdpzdbhfw",Collections.min(ljxsdbhfw)+ "~" + Collections.max(ljxsdbhfw));
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdpzdzb","IRI");
                map1.put("lmlx",sdlmlx);
                map1.put("sdpzdlmlx","隧道路面（主线/匝道）");
                map1.put("sdpzdgdz",sdsjz);
                map1.put("sdpzdjcds",decf.format(sdzds));
                map1.put("sdpzdhgds",decf.format(sdhgds));
                map1.put("sdpzdhgl",sdzds!=0 ? df.format(sdhgds/sdzds*100) : 0);
                map1.put("sdpzdbhfw",Collections.min(sdbhfw)+ "~" + Collections.max(sdbhfw));
                resultList.add(map1);
            }
            if (c){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdpzdzb","IRI");
                map1.put("lmlx",hntsdlmlx);
                map1.put("sdpzdlmlx","混凝土隧道");
                map1.put("sdpzdgdz",hntsdsjz);
                map1.put("sdpzdjcds",decf.format(hntsdzds));
                map1.put("sdpzdhgds",decf.format(hntsdhgds));
                map1.put("sdpzdhgl",hntsdzds!=0 ? df.format(hntsdhgds/hntsdzds*100) : 0);
                map1.put("sdpzdbhfw",Collections.min(hntsdbhfw)+ "~" + Collections.max(hntsdbhfw));
                resultList.add(map1);
            }
        }
        return resultList;

    }

    /**
     * 和getsdhntqdData一样
     * 有待确认的事项
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhntxlbgcData(CommonInfoVo commonInfoVo) throws IOException {
        /*List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcSdgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        double jcds = 0;
        double hgds = 0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        boolean a = false;
        if (list.size()>0){
            a = true;
            for (Map<String, Object> map : list) {
                jcds += Double.valueOf(map.get("检测点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());

                double max = Double.valueOf(map.get("最大值").toString());
                lmMax = (max > lmMax) ? max : lmMax;
                double min = Double.valueOf(map.get("最小值").toString());
                lmMin = (min < lmMin) ? min : lmMin;

            }

        }
        if (a){
            Map map = new HashMap();
            map.put("jcds",decf.format(jcds));
            map.put("hgds",decf.format(hgds));
            map.put("hgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            map.put("max",lmMax);
            map.put("min",lmMin);
            map.put("gdz",list.get(0).get("规定值"));
            map.put("pjz",list.get(0).get("平均值"));
            resultList.add(map);
        }
        return resultList;*/
        return null;
    }

    /**
     * 有待确认的事项
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhntqdData(CommonInfoVo commonInfoVo) throws IOException {
        /*List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcSdgcHntlmqdService.lookJdbjg(commonInfoVo);

        *//**
         * 这块会有多个隧道的数据，平均值如何取？
         * 最大值和最小值是取全部比较后的结构
         * 还有规定值
         *//*
        double jcds = 0;
        double hgds = 0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                jcds += Double.valueOf(map.get("检测点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());

                double max = Double.valueOf(map.get("最大值").toString());
                lmMax = (max > lmMax) ? max : lmMax;
                double min = Double.valueOf(map.get("最小值").toString());
                lmMin = (min < lmMin) ? min : lmMin;

            }
            Map map = new HashMap();
            map.put("jcds",decf.format(jcds));
            map.put("hgds",decf.format(hgds));
            map.put("hgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            map.put("max",lmMax);
            map.put("min",lmMin);
            map.put("gdz",list.get(0).get("规定值"));
            map.put("pjz",list.get(0).get("平均值"));
            resultList.add(map);
        }

        return resultList;*/
        return null;
    }

    /**
     * 表4.1.2-16  隧道路面渗水系数检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdssxsData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo,1);
        double jcds =0.0;
        double hgds =0.0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        boolean a = false;
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                String lmlx = map.get("检测项目").toString();
                if (lmlx.contains("隧道路面")){
                    a = true;
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }

            }
            if (a){
                Map<String,Object> map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("sdssxsjcds",jcds);
                map.put("sdssxsmax",decf.format(lmMax));
                map.put("sdssxsmin",decf.format(lmMin));
                map.put("sdssxshgds",hgds);
                map.put("sdssxshgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                map.put("sdssxsgdz",list.get(0).get("规定值"));
                resultList.add(map);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-19  隧道路面车辙检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdczData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> finalList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhCzService.lookJdbjg(commonInfoVo, 1);
        if (list!=null && list.size()>0){
            String ljxsjz = "";
            double ljxzds =0.0,ljxhgds = 0;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list) {
                Map map1 = new HashMap();
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("隧道")){
                        a = true;
                        ljxsjz = map.get("设计值").toString();
                        ljxzds += Double.valueOf(map.get("总点数").toString());
                        ljxhgds += Double.valueOf(map.get("合格点数").toString());
                        double max = Double.valueOf(map.get("Max").toString());
                        ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                        double min = Double.valueOf(map.get("Min").toString());
                        ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;
                    }
                }else {
                    if (lmlx.contains("隧道")){
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("sdczzb","RDI");
                        map1.put("sdczlmlx","隧道路面（主线）");
                        map1.put("sdczgdz",map.get("设计值"));
                        map1.put("sdczjcds",map.get("总点数"));
                        map1.put("sdczhgds",map.get("合格点数"));
                        map1.put("sdczhgl",map.get("合格率"));
                        map1.put("sdczbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);
                    }
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdczzb","RDI");
                map1.put("sdczlmlx","隧道路面（连接线）");
                map1.put("sdczgdz",ljxsjz);
                map1.put("sdczjcds",decf.format(ljxzds));
                map1.put("sdczhgds",decf.format(ljxhgds));
                map1.put("sdczhgl",ljxzds!=0 ? df.format(ljxhgds/ljxzds*100) : 0);
                map1.put("sdczbhfw",ljxlmMin+"~"+ljxlmMax);
                finalList.add(map1);
            }
        }


        if (resultList != null && resultList.size()>0){
            double hgds = 0,zds = 0;
            for (Map<String, Object> stringObjectMap : resultList) {
                hgds += Double.valueOf(stringObjectMap.get("sdczhgds").toString());
                zds += Double.valueOf(stringObjectMap.get("sdczjcds").toString());
            }

            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("sdczzb","RDI");
            map1.put("sdczlmlx","隧道路面（主线/匝道）");
            map1.put("sdczgdz",resultList.get(0).get("sdczgdz"));
            map1.put("sdczjcds",decf.format(zds));
            map1.put("sdczhgds",decf.format(hgds));
            map1.put("sdczhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            map1.put("sdczbhfw",resultList.get(0).get("sdczbhfw"));
            finalList.add(map1);
        }

        return finalList;

    }

    /**
     * 表4.1.2-15  隧道路面面层压实度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdysdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        //Double lmMax = 0.0;
        //Double lmMin = Double.MAX_VALUE;
        String bhfw = "";
        String gdz = "";
        String zdbz = "";
        String ydbz = "";
        Double jcds = 0.0;
        Double hgds = 0.0;
        boolean a= false;
        //还有连接线隧道
        /*if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                String lx = map.get("路面类型").toString();
                if (lx.contains("隧道")){
                    a = true;
                    *//*double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;*//*
                    bhfw = map.get("实测值变化范围").toString();
                    gdz = map.get("密度规定值").toString();
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                    if (lx.contains("隧道左幅")){
                        zdbz = map.get("密度代表值").toString();

                    }else if (lx.contains("隧道右幅")){
                        ydbz = map.get("密度代表值").toString();
                    }

                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("ysdlx","主线");
                //map1.put("ysdsczbh",lmMin+"~"+lmMax);
                map1.put("ysdsczbh",bhfw);
                map1.put("ysdzdbz",zdbz);
                map1.put("ysdydbz",ydbz);
                map1.put("ysdgdz",gdz);
                map1.put("ysdzs",decf.format(jcds));
                map1.put("ysdhgs",decf.format(hgds));
                map1.put("ysdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }
        }*/
        if (list!=null && list.size()>0){
            List<Map<String, Object>> sdbzmdList = new ArrayList<>();
            List<Map<String, Object>> ljxsdbzmdList = new ArrayList<>();
            List<Map<String, Object>> zdsdbzmdList = new ArrayList<>();


            List<Map<String, Object>> sdzdmdList = new ArrayList<>();
            List<Map<String, Object>> ljxsdzdmdList = new ArrayList<>();
            List<Map<String, Object>> zdsdzdmdList = new ArrayList<>();
            for (Map<String, Object> map : list) {
                String lx = map.get("路面类型").toString();
                String bz = map.get("标准").toString();
                if (bz.equals("标准密度")){
                    if (lx.contains("连接线隧道")){
                        ljxsdbzmdList.add(map);
                    }else if (lx.contains("匝道隧道")){
                        zdsdbzmdList.add(map);
                    }else if (lx.contains("隧道") && !lx.contains("连接线隧道") && !lx.contains("匝道隧道")){
                        sdbzmdList.add(map);
                    }
                }else if (bz.equals("最大理论密度")){
                    if (lx.contains("连接线隧道")){
                        ljxsdzdmdList.add(map);

                    }else if (lx.contains("匝道隧道")){
                        zdsdzdmdList.add(map);

                    }else if (lx.contains("隧道") && !lx.contains("连接线隧道") && !lx.contains("匝道隧道")){
                        sdzdmdList.add(map);
                    }

                }
            }
            if (sdbzmdList!=null && sdbzmdList.size()>0){
                List<Map<String, Object>> extractedmethod = sdextractedmethod(commonInfoVo, sdbzmdList, df,"主线隧道","试验室标准密度");
                resultList.addAll(extractedmethod);
            }

            if (ljxsdbzmdList!=null && ljxsdbzmdList.size()>0){
                List<Map<String, Object>> extractedmethod = sdextractedmethod(commonInfoVo, ljxsdbzmdList, df,"连接线隧道","试验室标准密度");
                resultList.addAll(extractedmethod);
            }

            if (zdsdbzmdList!=null && zdsdbzmdList.size()>0){
                List<Map<String, Object>> extractedmethod = sdextractedmethod(commonInfoVo, zdsdbzmdList, df,"匝道隧道","试验室标准密度");
                resultList.addAll(extractedmethod);
            }

            if (sdzdmdList!=null && sdzdmdList.size()>0){
                List<Map<String, Object>> extractedmethod = sdextractedmethod(commonInfoVo, sdzdmdList, df,"主线隧道","最大理论密度");
                resultList.addAll(extractedmethod);
            }

            if (ljxsdzdmdList!=null && ljxsdzdmdList.size()>0){
                List<Map<String, Object>> extractedmethod = sdextractedmethod(commonInfoVo, ljxsdzdmdList, df,"连接线隧道","最大理论密度");
                resultList.addAll(extractedmethod);
            }

            if (zdsdzdmdList!=null && zdsdzdmdList.size()>0){
                List<Map<String, Object>> extractedmethod = sdextractedmethod(commonInfoVo, zdsdzdmdList, df,"匝道隧道","最大理论密度");
                resultList.addAll(extractedmethod);
            }
        }

        return resultList;

    }

    private static List<Map<String, Object>> sdextractedmethod(CommonInfoVo commonInfoVo, List<Map<String, Object>> bzmdzxList, DecimalFormat df,String zx,String bz) {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map<String, List<Map<String, Object>>> result = bzmdzxList.stream()
                .collect(Collectors.groupingBy(map -> (String) map.get("密度规定值")));
        result.forEach((group, grouphtdData) -> {
            double jcds = 0,hgds = 0;
            String zxzfdbz="",zxyfdbz="";
            List<Double> bhfwz = new ArrayList<>();
            for (Map<String, Object> grouphtdDatum : grouphtdData) {
                jcds += Double.valueOf(grouphtdDatum.get("检测点数").toString());
                hgds += Double.valueOf(grouphtdDatum.get("合格点数").toString());
                String lmlx = grouphtdDatum.get("路面类型").toString();
                if (lmlx.contains("左幅")){
                    zxzfdbz = grouphtdDatum.get("密度代表值").toString();
                }else if(lmlx.contains("右幅")) {
                    zxyfdbz = grouphtdDatum.get("密度代表值").toString();
                }else {
                    zxzfdbz = grouphtdDatum.get("密度代表值").toString();
                    zxyfdbz = grouphtdDatum.get("密度代表值").toString();
                }
                String sczbhfw = grouphtdDatum.get("实测值变化范围").toString();
                bhfwz.add(Double.valueOf(StringUtils.substringBefore(sczbhfw,"~")));
                bhfwz.add(Double.valueOf(StringUtils.substringAfter(sczbhfw,"~")));
            }
            String finbzsczbhfw = "";
            if (!bhfwz.isEmpty()) {
                double min = Collections.min(bhfwz);
                double max = Collections.max(bhfwz);
                finbzsczbhfw = min + "~" + max;
            }
            Map map1 = new HashMap();
            map1.put("htd", commonInfoVo.getHtd());
            map1.put("ysdlx",zx);
            map1.put("bz",bz);
            map1.put("ysdsczbh",finbzsczbhfw);
            map1.put("ysdzdbz",zxzfdbz);
            map1.put("ysdydbz",zxyfdbz);
            map1.put("ysdgdz",group);
            map1.put("ysdzs",jcds);
            map1.put("ysdhgs",hgds);
            map1.put("ysdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            resultList.add(map1);
        });
        return resultList;
    }

    /**
     * 表4.1.2-14  桥面系检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmqlhzData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap();

        boolean a = false,b = false,c = false;
        List<Map<String, Object>> list1 = getqmpzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            a = true;
            double qmpzdjcds = 0;
            double qmpzdhgds = 0;
            for (Map<String, Object> map : list1) {
                qmpzdjcds += Double.valueOf(map.get("qmpzdjcds").toString());
                qmpzdhgds += Double.valueOf(map.get("qmpzdhgds").toString());
            }
            resultmap.put("qmpzdjcds",decf.format(qmpzdjcds));
            resultmap.put("qmpzdhgds",decf.format(qmpzdhgds));
            resultmap.put("qmpzdhgl",qmpzdjcds!=0 ? df.format(qmpzdhgds/qmpzdjcds*100) : 0);
        }
        List<Map<String, Object>> list2 = getqmhpData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            b =true;
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> map : list2) {
                zds += Double.valueOf(map.get("qmhpzds").toString());
                hgds += Double.valueOf(map.get("qmhphgds").toString());
            }
            resultmap.put("qmhpzds",decf.format(zds));
            resultmap.put("qmhphgds", decf.format(hgds));
            resultmap.put("qmhphgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }
        List<Map<String, Object>> list3 = getqmkhData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            c = true;
            double mcxszds = 0;
            double mcxshgds = 0;
            double gzsdzds = 0;
            double gzsdhgds = 0;
            for (Map<String, Object> map : list3) {
                String qmkhzb = map.get("qmkhzb").toString();
                if (qmkhzb.equals("摩擦系数")){
                    mcxszds += Double.valueOf(map.get("qmkhjcds").toString());
                    mcxshgds += Double.valueOf(map.get("qmkhhgds").toString());
                }else if (qmkhzb.equals("构造深度")){
                    gzsdzds += Double.valueOf(map.get("qmkhjcds").toString());
                    gzsdhgds += Double.valueOf(map.get("qmkhhgds").toString());
                }
            }
            resultmap.put("mcxszds",decf.format(mcxszds));
            resultmap.put("mcxshgds", decf.format(mcxshgds));
            resultmap.put("mcxshgl",mcxszds!=0 ? df.format(mcxshgds/mcxszds*100) : 0);

            resultmap.put("gzsdzds",decf.format(gzsdzds));
            resultmap.put("gzsdhgds", decf.format(gzsdhgds));
            resultmap.put("gzsdhgl",gzsdzds!=0 ? df.format(gzsdhgds/gzsdzds*100) : 0);
        }
        if (a || b || c){
            resultmap.put("htd",commonInfoVo.getHtd());
            resultList.add(resultmap);
        }
        System.out.println(resultList);
        return resultList;
    }

    /**
     * 表4.1.2-13  桥面抗滑检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqmkhData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhGzsdService.lookJdbjg(commonInfoVo, 1);//自动化
        String sjz = "",ljxsjz = "";
        double jcds = 0.0;
        double hgds = 0.0;
        double ljxjcds = 0.0;
        double ljxhgds = 0.0;

        if (list!=null && list.size()>0){
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            boolean a = false,b = false;
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("桥")){
                        b = true;
                        ljxsjz = map.get("设计值").toString();
                        ljxjcds += Double.valueOf(map.get("总点数").toString());
                        ljxhgds += Double.valueOf(map.get("合格点数").toString());

                        double max = Double.valueOf(map.get("最大值").toString());
                        ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                        double min = Double.valueOf(map.get("最小值").toString());
                        ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;
                    }
                }else {
                    if (lmlx.contains("桥")){
                        a = true;
                        sjz = map.get("设计值").toString();
                        jcds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                        double max = Double.valueOf(map.get("最大值").toString());
                        lmMax = (max > lmMax) ? max : lmMax;

                        double min = Double.valueOf(map.get("最小值").toString());
                        lmMin = (min < lmMin) ? min : lmMin;
                    }
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","沥青桥（主线/匝道）");
                map1.put("qmkhzb","构造深度");
                map1.put("qmkhsjz",sjz);
                map1.put("qmkhbhfw",lmMin+"~"+lmMax);
                map1.put("qmkhjcds",decf.format(jcds));
                map1.put("qmkhhgds",decf.format(hgds));
                map1.put("qmkhhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","沥青桥(连接线)");
                map1.put("qmkhzb","构造深度");
                map1.put("qmkhsjz",ljxsjz);
                map1.put("qmkhbhfw",ljxlmMin+"~"+ljxlmMax);
                map1.put("qmkhjcds",decf.format(ljxjcds));
                map1.put("qmkhhgds",decf.format(ljxhgds));
                map1.put("qmkhhgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list1 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (list1!=null && list1.size()>0){
            String gdz = "";
            double szds = 0;
            double shgs = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("桥")){
                    a = true;
                    szds += Double.valueOf(map.get("检测点数").toString());
                    shgs += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                }else {
                    continue;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","混凝土桥");
                map1.put("qmkhzb","构造深度");
                map1.put("qmkhsjz",gdz);
                map1.put("qmkhbhfw",lmMin+"~"+lmMax);
                map1.put("qmkhjcds",decf.format(szds));
                map1.put("qmkhhgds",decf.format(shgs));
                map1.put("qmkhhgl",szds!=0 ? df.format(shgs/szds*100) : 0);
                resultList.add(map1);
            }


        }
        List<Map<String, Object>> list2 = jjgZdhMcxsService.lookJdbjg(commonInfoVo,1);
        if (list2!=null && list2.size()>0){
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            double zds = 0;
            double hgs = 0;
            double ljxzds = 0;
            double ljxhgs = 0;
            boolean a = false,b = false;
            String sjz2 = "",mcxsljxsjz = "";
            for (Map<String, Object> map : list2) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("桥")){
                        b = true;
                        mcxsljxsjz = map.get("设计值").toString();
                        ljxzds += Double.valueOf(map.get("总点数").toString());
                        ljxhgs += Double.valueOf(map.get("合格点数").toString());

                        double max = Double.valueOf(map.get("Max").toString());
                        ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                        double min = Double.valueOf(map.get("Min").toString());
                        ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;
                    }

                }else {
                    if (lmlx.contains("桥")){
                        a = true;
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgs += Double.valueOf(map.get("合格点数").toString());
                        sjz2 = map.get("设计值").toString();

                        double max = Double.valueOf(map.get("Max").toString());
                        lmMax = (max > lmMax) ? max : lmMax;

                        double min = Double.valueOf(map.get("Min").toString());
                        lmMin = (min < lmMin) ? min : lmMin;
                    }
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","沥青桥（主线/匝道）");
                map1.put("qmkhzb","摩擦系数");
                map1.put("qmkhsjz",sjz2);
                map1.put("qmkhbhfw",lmMin+"~"+lmMax);
                map1.put("qmkhjcds",decf.format(zds));
                map1.put("qmkhhgds",decf.format(hgs));
                map1.put("qmkhhgl",zds!=0 ? df.format(hgs/zds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","沥青桥（连接线）");
                map1.put("qmkhzb","摩擦系数");
                map1.put("qmkhsjz",mcxsljxsjz);
                map1.put("qmkhbhfw",df.format(ljxlmMin)+"~"+df.format(ljxlmMax));
                map1.put("qmkhjcds",decf.format(ljxzds));
                map1.put("qmkhhgds",decf.format(ljxhgs));
                map1.put("qmkhhgl",ljxzds!=0 ? df.format(ljxhgs/ljxzds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;
    }

    /**
     *表4.1.2-12  桥面系横坡检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqmhpData(CommonInfoVo commonInfoVo) throws IOException {

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, String>> list = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);

        if (list!=null && list.size()>0){
            double lmzds = 0.0;
            double lmhgds = 0.0;
            double ljxlmzds = 0.0;
            double ljxlmhgds = 0.0;

            double lmzds1 = 0.0;
            double lmhgds1 = 0.0;
            boolean a = false;
            boolean b = false,c = false;
            for (Map<String, String> map : list) {
                String jcxm = map.get("检测项目");
                String lmlx = map.get("路面类型");
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("沥青桥")){
                        c = true;
                        ljxlmzds +=Double.valueOf(map.get("检测点数"));
                        ljxlmhgds +=Double.valueOf(map.get("合格点数"));
                    }
                }else {
                    if (lmlx.contains("沥青桥面")){
                        a = true;
                        lmzds +=Double.valueOf(map.get("检测点数"));
                        lmhgds +=Double.valueOf(map.get("合格点数"));

                    }
                    if (lmlx.contains("混凝土桥面")){
                        b = true;
                        lmzds1 +=Double.valueOf(map.get("检测点数"));
                        lmhgds1 +=Double.valueOf(map.get("合格点数"));
                    }
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmhpzds",decf.format(lmzds));
                map1.put("qmhphgds",decf.format(lmhgds));
                map1.put("qmhphgl",lmzds!=0 ? df.format(lmhgds/lmzds*100) : 0);
                map1.put("qmhplmlx","沥青桥面(主线/匝道)");
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("qmhphtd",commonInfoVo.getHtd());
                map2.put("qmhpzds",decf.format(lmzds1));
                map2.put("qmhphgds",decf.format(lmhgds1));
                map2.put("qmhphgl",lmzds1!=0 ? df.format(lmhgds1/lmzds1*100) : 0);
                map2.put("qmhplmlx","混凝土桥面");
                resultList.add(map2);
            }
            if (c){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("hpzds",decf.format(ljxlmzds));
                map2.put("hphgds",decf.format(ljxlmhgds));
                map2.put("hphgl",ljxlmzds!=0 ? df.format(ljxlmhgds/ljxlmzds*100) : 0);
                map2.put("hplmlx","沥青桥面(连接线)");
                resultList.add(map2);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-11  桥面铺装平整度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqmpzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.0");
        List<Map<String, Object>> list = jjgZdhPzdService.lookJdbjg(commonInfoVo, 1);
        if (list != null && list.size()>0){
            boolean a= false;
            double zdzds = 0,zdhgds = 0;
            double zxz = 0,zdz = 0;
            String zdlmlx = "", zdsjz = "";
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                if (jcxm.contains("连接线")){
                    if (lmlx.contains("桥")){
                        Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("qmpzdzb","IRI");
                        map1.put("lmlx",lmlx);
                        map1.put("qmpzdlmlx","沥青路面（连接线）");
                        map1.put("qmpzdgdz",map.get("设计值"));
                        map1.put("qmpzdjcds",map.get("总点数"));
                        map1.put("qmpzdhgds",map.get("合格点数"));
                        map1.put("qmpzdhgl",map.get("合格率"));
                        map1.put("qmpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);
                    }
                }else if (!jcxm.contains("主线") && !jcxm.contains("连接线")){
                    if (lmlx.contains("桥")){
                        //匝道
                        a = true;
                        zdzds += Double.valueOf(map.get("总点数").toString());
                        zdhgds += Double.valueOf(map.get("合格点数").toString());
                        zdlmlx = lmlx;
                        zdsjz = map.get("设计值").toString();
                        zxz = Double.valueOf(map.get("Min").toString());
                        zdz = Double.valueOf(map.get("Max").toString());

                        double zxz1 = Double.valueOf(map.get("Min").toString());
                        double zdz1 = Double.valueOf(map.get("Max").toString());
                        if (zxz1<zxz){
                            zxz = zxz1;
                        }
                        if (zdz1>zdz){
                            zdz = zdz1;
                        }
                    }
                }else {
                    if (lmlx.equals("沥青桥")){
                        Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("qmpzdzb","IRI");
                        map1.put("qmpzdlmlx","沥青路面(主线)");
                        map1.put("qmpzdgdz",map.get("设计值"));
                        map1.put("qmpzdjcds",map.get("总点数"));
                        map1.put("qmpzdhgds",map.get("合格点数"));
                        map1.put("qmpzdhgl",map.get("合格率"));
                        map1.put("qmpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);
                    }
                    if (lmlx.equals("混凝土桥")){
                        Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("qmpzdzb","IRI");
                        map1.put("qmpzdlmlx","混凝土路面");
                        map1.put("qmpzdgdz",map.get("设计值"));
                        map1.put("qmpzdjcds",map.get("总点数"));
                        map1.put("qmpzdhgds",map.get("合格点数"));
                        map1.put("qmpzdhgl",map.get("合格率"));
                        map1.put("qmpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);
                    }
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmpzdzb","IRI");
                map1.put("lmlx",zdlmlx);
                map1.put("qmpzdlmlx","沥青路面（匝道）");
                map1.put("qmpzdgdz",zdsjz);
                map1.put("qmpzdjcds",zdzds);
                map1.put("qmpzdhgds",zdhgds);
                map1.put("qmpzdbhfw",df.format(zxz)+"~"+df.format(zdz));
                map1.put("qmpzdhgl",zdzds!=0 ? df.format(zdhgds/zdzds*100) : 0);
                resultList.add(map1);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-10  路面工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getlmhzData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        //处理一下list1，将所有的检测点数和合格点数相加，算合格率
        //List<Map<String, Object>> list1 = getlmysdData(commonInfoVo);
        List<Map<String, Object>> list1 = jjgFbgcLmgcLqlmysdService.lookJdbjgbgzbg(commonInfoVo, 1);
        if (CollectionUtils.isNotEmpty(list1)){
            double zds = 0,ljxzds = 0;
            double hgds = 0,ljxhgds = 0;
            boolean a = false,b = false;
            for (Map<String, Object> map : list1) {
                String lmlc = map.get("路面类型").toString();
                if (!lmlc.contains("隧道") && !lmlc.contains("连接线") ){
                    a = true;
                    zds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
                if (lmlc.contains("连接线")){
                    b = true;
                    ljxzds += Double.valueOf(map.get("检测点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            if (a){
                resultmap.put("ysdzds",decf.format(zds));
                resultmap.put("ysdhgds",decf.format(hgds));
                resultmap.put("ysdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
                resultmap.put("htd",list1.get(0).get("htd"));
            }
            if (b){
                resultmap.put("ljxysdzds",decf.format(ljxzds));
                resultmap.put("ljxysdhgds",decf.format(ljxhgds));
                resultmap.put("ljxysdhgl",ljxzds!=0 ? df.format(ljxhgds/ljxzds*100) : 0);
                resultmap.put("htd",list1.get(0).get("htd"));
            }


        }
        List<Map<String, Object>> list2 = getlmwcallData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        //弯沉的话，待确认
        /*List<Map<String, Object>> list2 = getlmwcData(commonInfoVo);
        List<Map<String, Object>> list3 = getlmwclcfData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list2);
        }else {
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list3);
        }*/

        //车辙是分桥 隧道 路面  这块只拿了路面的数据
        List<Map<String, Object>> list4 = getczData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            for (Map<String, Object> map : list4) {
                String czlmlx = map.get("czlmlx").toString();
                if (czlmlx.contains("连接线")){
                    resultmap.put("htd",map.get("htd").toString());
                    resultmap.put("ljxczzds",map.get("czzds").toString());
                    resultmap.put("ljxczhgds",map.get("czhgds").toString());
                    resultmap.put("ljxczhgl",map.get("czhgl").toString());
                }else {
                    resultmap.putAll(map);
                }
            }
        }

        List<Map<String, Object>> list5 = getlmssxsData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list5)){
            for (Map<String, Object> map : list5) {
                String lxmc = map.get("lxmc").toString();
                if (lxmc.contains("连接线")){
                    resultmap.put("htd",map.get("htd").toString());
                    resultmap.put("ljxlmssxsssjcds",map.get("lmssxsssjcds").toString());
                    resultmap.put("ljxlmssxssshgds",map.get("lmssxssshgds").toString());
                    resultmap.put("ljxlmssxssshgl",map.get("lmssxssshgl").toString());
                }else {
                    resultmap.putAll(map);
                }
            }
            /*double sszds = 0,sshgds = 0;
            for (Map<String, Object> map : list5) {
                sszds += Double.valueOf(map.get("lmssxsssjcds").toString());
                sshgds += Double.valueOf(map.get("lmssxssshgds").toString());
            }
            resultmap.put("lmssxsssjcds",decf.format(sszds));
            resultmap.put("lmssxssshgds",decf.format(sshgds));
            resultmap.put("lmssxssshgl",sszds!=0 ? df.format(sshgds/sszds*100) : 0);
            resultmap.put("htd",list5.get(0).get("htd"));*/
        }
        List<Map<String, Object>> list6 = gethntqdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list6)){
            for (Map<String, Object> map : list6) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list6);
        }
        List<Map<String, Object>> list7 = gethntxlbgcData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list7)){
            for (Map<String, Object> map : list7) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list7);
        }
        //平整度需要分沥青和水泥 list8分了 根据lmlx判断
        List<Map<String, Object>> list8 = getpzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list8)){
            double zds = 0,hgds = 0;
            boolean a =false;
            for (Map<String, Object> map : list8) {
                String hplmlx = map.get("pzdlmlx").toString();
                if (hplmlx.contains("连接线")){
                    resultmap.put("ljxpzdlqjcds",map.get("pzdjcds"));
                    resultmap.put("ljxpzdlqhgds",map.get("pzdhgds"));
                    resultmap.put("ljxpzdlqhgl",map.get("pzdhgl"));
                }else {
                    if (hplmlx.contains("沥青路面")){
                        a = true;
                        zds += Double.valueOf(map.get("pzdjcds").toString());
                        hgds += Double.valueOf(map.get("pzdhgds").toString());

                    }else if (hplmlx.contains("混凝土路面")){
                        resultmap.put("pzdhntjcds",map.get("pzdjcds"));
                        resultmap.put("pzdhnthgds",map.get("pzdhgds"));
                        resultmap.put("pzdhnthgl",map.get("pzdhgl"));
                    }
                }

            }
            if (a){
                resultmap.put("pzdlqjcds",zds);
                resultmap.put("pzdlqhgds",hgds);
                resultmap.put("pzdlqhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            }
            //resultList.addAll(list8);
        }
        List<Map<String, Object>> list9 = getkhData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list9)){
            for (Map<String, Object> map : list9) {
                String hplmlx = map.get("khlmlx").toString();
                if (hplmlx.contains("连接线")){
                    String khkpzb = map.get("khkpzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        resultmap.put("ljxkhlqmcxsjcds",map.get("khjcds"));
                        resultmap.put("ljxkhlqmcxshgds",map.get("khhgds"));
                        resultmap.put("ljxkhlqmcxshgl",map.get("khhgl"));
                    }else if (khkpzb.equals("构造深度")){
                        resultmap.put("ljxkhlqgzsdjcds",map.get("khjcds"));
                        resultmap.put("ljxkhlqgzsdhgds",map.get("khhgds"));
                        resultmap.put("ljxkhlqgzsdhgl",map.get("khhgl"));
                    }

                }else {
                    if (hplmlx.contains("沥青路面")){
                        String khkpzb = map.get("khkpzb").toString();
                        if (khkpzb.equals("摩擦系数")){
                            resultmap.put("khlqmcxsjcds",map.get("khjcds"));
                            resultmap.put("khlqmcxshgds",map.get("khhgds"));
                            resultmap.put("khlqmcxshgl",map.get("khhgl"));
                        }else if (khkpzb.equals("构造深度")){
                            resultmap.put("khlqgzsdjcds",map.get("khjcds"));
                            resultmap.put("khlqgzsdhgds",map.get("khhgds"));
                            resultmap.put("khlqgzsdhgl",map.get("khhgl"));
                        }
                    }else if (hplmlx.equals("混凝土路面")){
                        String khkpzb = map.get("khkpzb").toString();
                        if (khkpzb.equals("摩擦系数")){
                            resultmap.put("khhntmcxsjcds",map.get("khjcds"));
                            resultmap.put("khhntmcxshgds",map.get("khhgds"));
                            resultmap.put("khhntmcxshgl",map.get("khhgl"));
                        }else if (khkpzb.equals("构造深度")){
                            resultmap.put("khhntgzsdjcds",map.get("khjcds"));
                            resultmap.put("khhntgzsdhgds",map.get("khhgds"));
                            resultmap.put("khhntgzsdhgl",map.get("khhgl"));
                        }

                    }
                }

            }
        }

        List<Map<String, Object>> list10 = gethpData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list10)){
            for (Map<String, Object> map : list10) {
                String hplmlx = map.get("hplmlx").toString();

                if (hplmlx.equals("沥青路面(主线/匝道)")){
                    resultmap.put("hplqzds",map.get("hpzds"));
                    resultmap.put("hplqhgds",map.get("hphgds"));
                    resultmap.put("hplqhgl",map.get("hphgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("hphntzds",map.get("hpzds"));
                    resultmap.put("hphnthgds",map.get("hphgds"));
                    resultmap.put("hphnthgl",map.get("hphgl"));
                }else if (hplmlx.equals("沥青路面(连接线)")){
                    resultmap.put("ljxhphntzds",map.get("hpzds"));
                    resultmap.put("ljxhphnthgds",map.get("hphgds"));
                    resultmap.put("ljxhphnthgl",map.get("hphgl"));
                }
            }
        }
        List<Map<String, Object>> list13 = getlmhdData(commonInfoVo);

        if (CollectionUtils.isNotEmpty(list13)){
            double ljxzds = 0,ljxhgds = 0;
            double lqzds = 0,lqhgds = 0;
            boolean a = false;
            boolean b = false;
            for (Map<String, Object> map : list13) {
                String hplmlx = map.get("lmhdlmlx").toString();
                if (hplmlx.contains("连接线")){
                    a = true;
                    ljxzds += Double.valueOf(map.get("lmhdjcds").toString());
                    ljxhgds += Double.valueOf(map.get("lmhdhgs").toString());

                }else {
                    if (hplmlx.contains("沥青路面")){
                        b=true;
                        lqzds += Double.valueOf(map.get("lmhdjcds").toString());
                        lqhgds += Double.valueOf(map.get("lmhdhgs").toString());

                    }else if (hplmlx.equals("混凝土路面")){
                        resultmap.put("lmhdhntjcds",map.get("lmhdjcds"));
                        resultmap.put("lmhdhnthgs",map.get("lmhdhgs"));
                        resultmap.put("lmhdhnthgl",map.get("lmhdhgl"));
                    }
                }

            }
            if (a){
                resultmap.put("ljxlmhdlqjcds",ljxzds);
                resultmap.put("ljxlmhdlqhgs",ljxhgds);
                resultmap.put("ljxlmhdlqhgl",ljxzds!=0 ? df.format(ljxhgds/ljxzds*100) : 0);
            }
            if (b){
                resultmap.put("lmhdlqjcds",lqzds);
                resultmap.put("lmhdlqhgs",lqhgds);
                resultmap.put("lmhdlqhgl",lqzds!=0 ? df.format(lqhgds/lqzds*100) : 0);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     * 表4.1.2-8（3）  厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        //雷达法
        List<Map<String, Object>> list2 = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        boolean a = false,b = false,c= false,d = false,e = false,f = false,g = false;
        double smcgsjcds = 0,smcgshgds = 0,zdsmcgsjcds = 0,zdsmcgshgds = 0,zdzhdgsjcds = 0,zdzhdgshgds = 0,zhdgsjcds = 0,zhdgshgds = 0,ldjcds = 0,ldhgds = 0,smcljxgsjcds = 0,smcljxgshgds = 0,zhdljxgsjcds = 0,zhdljxgshgds = 0,fhlmldjcds = 0,fhlmldhgds = 0,ljxldjcds = 0,ljxldhgds = 0,zaldjcds = 0,zaldhgds = 0;
        if (list!=null && list.size()>0) {

            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("连接线")){
                    c = true;
                    smcljxgsjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    zhdljxgsjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    smcljxgshgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    zhdljxgshgds += Double.valueOf(map.get("总厚度合格点数").toString());

                }else if (lmlx.equals("路面匝道")){
                    f = true;
                    zdsmcgsjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    zdzhdgsjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    zdsmcgshgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    zdzhdgshgds += Double.valueOf(map.get("总厚度合格点数").toString());

                }else if (lmlx.equals("路面左幅") || lmlx.equals("路面右幅")){
                    a = true;
                    smcgsjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    zhdgsjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    smcgshgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    zhdgshgds += Double.valueOf(map.get("总厚度合格点数").toString());
                }
            }
            if (a){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面上面层厚度(高速沥青厚度钻芯法-主线)");
                map2.put("lmhdjcds",decf.format(smcgsjcds));
                map2.put("lmhdhgs",decf.format(smcgshgds));
                map2.put("lmhdhgl",smcgsjcds!=0 ? df.format(smcgshgds/smcgsjcds*100) : 0);
                resultList.add(map2);
                Map map3 = new HashMap();
                map3.put("htd",commonInfoVo.getHtd());
                map3.put("lmhdlmlx","沥青路面总厚度(高速沥青厚度钻芯法-主线)");
                map3.put("lmhdjcds",decf.format(zhdgsjcds));
                map3.put("lmhdhgs",decf.format(zhdgshgds));
                map3.put("lmhdhgl",zhdgsjcds!=0 ? df.format(zhdgshgds/zhdgsjcds*100) : 0);
                resultList.add(map3);
            }
            if (f){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面上面层厚度(高速沥青厚度钻芯法-匝道)");
                map2.put("lmhdjcds",decf.format(zdsmcgsjcds));
                map2.put("lmhdhgs",decf.format(zdsmcgshgds));
                map2.put("lmhdhgl",zdsmcgsjcds!=0 ? df.format(zdsmcgshgds/zdsmcgsjcds*100) : 0);
                resultList.add(map2);
                Map map3 = new HashMap();
                map3.put("htd",commonInfoVo.getHtd());
                map3.put("lmhdlmlx","沥青路面总厚度(高速沥青厚度钻芯法-匝道)");
                map3.put("lmhdjcds",decf.format(zdzhdgsjcds));
                map3.put("lmhdhgs",decf.format(zdzhdgshgds));
                map3.put("lmhdhgl",zdzhdgsjcds!=0 ? df.format(zdzhdgshgds/zdzhdgsjcds*100) : 0);
                resultList.add(map3);
            }
            if (c){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面上面层厚度(高速沥青厚度钻芯法-连接线)");
                map2.put("lmhdjcds",decf.format(smcljxgsjcds));
                map2.put("lmhdhgs",decf.format(smcljxgshgds));
                map2.put("lmhdhgl",smcljxgsjcds!=0 ? df.format(smcljxgshgds/smcljxgsjcds*100) : 0);
                resultList.add(map2);
                Map map3 = new HashMap();
                map3.put("htd",commonInfoVo.getHtd());
                map3.put("lmhdlmlx","沥青路面总厚度(高速沥青厚度钻芯法-连接线)");
                map3.put("lmhdjcds",decf.format(zhdljxgsjcds));
                map3.put("lmhdhgs",decf.format(zhdljxgshgds));
                map3.put("lmhdhgl",zhdljxgsjcds!=0 ? df.format(zhdljxgshgds/zhdljxgsjcds*100) : 0);
                resultList.add(map3);
            }
        }
        if (list2!=null && list2.size()>0) {
            for (Map<String, Object> map : list2) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                if (jcxm.contains("连接线")){
                    e = true;
                    ljxldjcds += Double.valueOf(map.get("总点数").toString());
                    ljxldhgds += Double.valueOf(map.get("合格点数").toString());
                }else if (jcxm.equals("主线")){
                    if (lmlx.contains("复合路面")){
                        d = true;
                        fhlmldjcds += Double.valueOf(map.get("总点数").toString());
                        fhlmldhgds += Double.valueOf(map.get("合格点数").toString());
                    }else {
                        if (lmlx.equals("左幅") || lmlx.equals("右幅")){
                            b = true;
                            ldjcds += Double.valueOf(map.get("总点数").toString());
                            ldhgds += Double.valueOf(map.get("合格点数").toString());
                        }
                    }
                }else {
                    if (lmlx.equals("左幅") ||lmlx.equals("右幅") ){
                        g = true;
                        zaldjcds += Double.valueOf(map.get("总点数").toString());
                        zaldhgds += Double.valueOf(map.get("合格点数").toString());
                    }
                }


            }
            if (b){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面(路面雷达厚度-主线)");
                map2.put("lmhdjcds",decf.format(ldjcds));
                map2.put("lmhdhgs",decf.format(ldhgds));
                map2.put("lmhdhgl",ldjcds!=0 ? df.format(ldhgds/ldjcds*100) : 0);
                resultList.add(map2);
            }
            if (g){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面(路面雷达厚度-匝道)");
                map2.put("lmhdjcds",decf.format(zaldjcds));
                map2.put("lmhdhgs",decf.format(zaldhgds));
                map2.put("lmhdhgl",zaldjcds!=0 ? df.format(zaldhgds/zaldjcds*100) : 0);
                resultList.add(map2);
            }
            if (d){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面(路面雷达厚度-复合路面)");
                map2.put("lmhdjcds",decf.format(fhlmldjcds));
                map2.put("lmhdhgs",decf.format(fhlmldhgds));
                map2.put("lmhdhgl",fhlmldjcds!=0 ? df.format(fhlmldhgds/fhlmldjcds*100) : 0);
                resultList.add(map2);
            }
            if (e){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("lmhdlmlx","沥青路面(路面雷达厚度-连接线)");
                map2.put("lmhdjcds",decf.format(ljxldjcds));
                map2.put("lmhdhgs",decf.format(ljxldhgds));
                map2.put("lmhdhgl",ljxldjcds!=0 ? df.format(ljxldhgds/ljxldjcds*100) : 0);
                resultList.add(map2);
            }
        }
        //混凝土
        List<Map<String, Object>> list1 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("lmhdlmlx","混凝土路面");
            map2.put("lmhdjcds",list1.get(0).get("检测点数"));
            map2.put("lmhdhgs",list1.get(0).get("合格点数"));
            map2.put("lmhdhgl",list1.get(0).get("合格率"));
            resultList.add(map2);
        }
        return resultList;

    }

    /**
     * 表4.1.2-8（1）  钻芯法厚度检测结果汇总表
     * 待解决  根据设计值分开，鉴定表生成的时候，要根据设计值分工作簿
     * 钻芯法厚度（总厚度是两行），是匝道的？
     *
     * 关于沥青上面层，需要将桥面左右幅，路面左右幅，和路面匝道的
     * 且设计值相同的检测点数相加，
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getzxfhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list)) {
            List<Map<String, Object>> smc = new ArrayList<>();
            List<Map<String, Object>> zhd = new ArrayList<>();
            List<Map<String, Object>> ljxsmc = new ArrayList<>();
            List<Map<String, Object>> ljxzhd = new ArrayList<>();
            /*List<Map<String, Object>> sdsmc = new ArrayList<>();
            List<Map<String, Object>> sdzhd = new ArrayList<>();*/
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (!lmlx.contains("隧道") && !lmlx.contains("连接线")) {
                    Map map1 = new HashMap();
                    Map map2 = new HashMap();
                    map1.put("上面层厚度合格点数", map.get("上面层厚度合格点数").toString());
                    map1.put("上面层平均值最小值", map.get("上面层平均值最小值").toString());
                    map1.put("上面层设计值", map.get("上面层设计值").toString());
                    map1.put("上面层代表值", map.get("上面层代表值").toString());
                    map1.put("上面层厚度合格率", map.get("上面层厚度合格率").toString());
                    map1.put("上面层平均值最大值", map.get("上面层平均值最大值").toString());
                    map1.put("上面层厚度检测点数", map.get("上面层厚度检测点数").toString());
                    map1.put("路面类型", map.get("路面类型").toString());
                    smc.add(map1);

                    map2.put("总厚度合格点数", map.get("总厚度合格点数").toString());
                    map2.put("总厚度平均值最小值", map.get("总厚度平均值最小值").toString());
                    map2.put("总厚度检测点数", map.get("总厚度检测点数").toString());
                    map2.put("总厚度代表值", map.get("总厚度代表值").toString());
                    map2.put("总厚度平均值最大值", map.get("总厚度平均值最大值").toString());
                    map2.put("总厚度设计值", map.get("总厚度设计值").toString());
                    map2.put("总厚度合格率", map.get("总厚度合格率").toString());
                    map2.put("路面类型", map.get("路面类型").toString());
                    zhd.add(map2);
                }
                if (lmlx.contains("连接线")){
                    Map map3 = new HashMap();
                    Map map4 = new HashMap();
                    map3.put("上面层厚度合格点数", map.get("上面层厚度合格点数").toString());
                    map3.put("上面层平均值最小值", map.get("上面层平均值最小值").toString());
                    map3.put("上面层设计值", map.get("上面层设计值").toString());
                    map3.put("上面层代表值", map.get("上面层代表值").toString());
                    map3.put("上面层厚度合格率", map.get("上面层厚度合格率").toString());
                    map3.put("上面层平均值最大值", map.get("上面层平均值最大值").toString());
                    map3.put("上面层厚度检测点数", map.get("上面层厚度检测点数").toString());
                    map3.put("路面类型", map.get("路面类型").toString());
                    ljxsmc.add(map3);

                    map4.put("总厚度合格点数", map.get("总厚度合格点数").toString());
                    map4.put("总厚度平均值最小值", map.get("总厚度平均值最小值").toString());
                    map4.put("总厚度检测点数", map.get("总厚度检测点数").toString());
                    map4.put("总厚度代表值", map.get("总厚度代表值").toString());
                    map4.put("总厚度平均值最大值", map.get("总厚度平均值最大值").toString());
                    map4.put("总厚度设计值", map.get("总厚度设计值").toString());
                    map4.put("总厚度合格率", map.get("总厚度合格率").toString());
                    map4.put("路面类型", map.get("路面类型").toString());
                    ljxzhd.add(map4);
                }
                /*if (lmlx.contains("隧道")){
                    Map map3 = new HashMap();
                    Map map4 = new HashMap();
                    map3.put("上面层厚度合格点数", map.get("上面层厚度合格点数").toString());
                    map3.put("上面层平均值最小值", map.get("上面层平均值最小值").toString());
                    map3.put("上面层设计值", map.get("上面层设计值").toString());
                    map3.put("上面层代表值", map.get("上面层代表值").toString());
                    map3.put("上面层厚度合格率", map.get("上面层厚度合格率").toString());
                    map3.put("上面层平均值最大值", map.get("上面层平均值最大值").toString());
                    map3.put("上面层厚度检测点数", map.get("上面层厚度检测点数").toString());
                    map3.put("路面类型", map.get("路面类型").toString());
                    sdsmc.add(map3);

                    map4.put("总厚度合格点数", map.get("总厚度合格点数").toString());
                    map4.put("总厚度平均值最小值", map.get("总厚度平均值最小值").toString());
                    map4.put("总厚度检测点数", map.get("总厚度检测点数").toString());
                    map4.put("总厚度代表值", map.get("总厚度代表值").toString());
                    map4.put("总厚度平均值最大值", map.get("总厚度平均值最大值").toString());
                    map4.put("总厚度设计值", map.get("总厚度设计值").toString());
                    map4.put("总厚度合格率", map.get("总厚度合格率").toString());
                    map4.put("路面类型", map.get("路面类型").toString());
                    sdzhd.add(map4);

                }*/
            }
            Map<String, List<Map<String, Object>>> smcgrouped =
                    smc.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("上面层设计值").toString()
                            ));
            Map<String, List<Map<String, Object>>> ljxsmcgrouped =
                    ljxsmc.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("上面层设计值").toString()
                            ));

            /*Map<String, List<Map<String, Object>>> sdsmcgrouped =
                    sdsmc.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("上面层设计值").toString()
                            ));*/

            Map<String, List<Map<String, Object>>> zhdgrouped =
                    zhd.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("总厚度设计值").toString()
                            ));
            Map<String, List<Map<String, Object>>> ljxzhdgrouped =
                    ljxzhd.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("总厚度设计值").toString()
                            ));
            /*Map<String, List<Map<String, Object>>> sdzhdgrouped =
                    sdzhd.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("总厚度设计值").toString()
                            ));*/
            for (List<Map<String, Object>> slist : smcgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : slist) {
                    zds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("上面层厚度合格点数").toString());

                    double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("上面层代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("上面层代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("zxfhdlmlx","沥青路面(主线/匝道)");
                mapzhd.put("zxfhdlb","上面层厚度");
                mapzhd.put("zxfhdjcds", zds);
                mapzhd.put("zxfhdhgs", hgds);
                mapzhd.put("zxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("zxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("zxfhdsjz", slist.get(0).get("上面层设计值"));
                mapzhd.put("zxfhdzdbz", zzhdbz);
                mapzhd.put("zxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }
            for (List<Map<String, Object>> slist : ljxsmcgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : slist) {
                    zds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("上面层厚度合格点数").toString());

                    double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("上面层代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("上面层代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("zxfhdlmlx","沥青路面（连接线）");
                mapzhd.put("zxfhdlb","上面层厚度");
                mapzhd.put("zxfhdjcds", zds);
                mapzhd.put("zxfhdhgs", hgds);
                mapzhd.put("zxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("zxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("zxfhdsjz", slist.get(0).get("上面层设计值"));
                mapzhd.put("zxfhdzdbz", zzhdbz);
                mapzhd.put("zxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }

            for (List<Map<String, Object>> zlist : zhdgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : zlist) {
                    zds += Double.valueOf(map.get("总厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("总厚度合格点数").toString());

                    double max = Double.valueOf(map.get("总厚度平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("总厚度平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("总厚度代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("总厚度代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("zxfhdlmlx","沥青路面(主线/匝道)");
                mapzhd.put("zxfhdlb","总厚度");
                mapzhd.put("zxfhdjcds", zds);
                mapzhd.put("zxfhdhgs", hgds);
                mapzhd.put("zxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("zxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("zxfhdsjz", zlist.get(0).get("总厚度设计值"));
                mapzhd.put("zxfhdzdbz", zzhdbz);
                mapzhd.put("zxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }
            for (List<Map<String, Object>> zlist : ljxzhdgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : zlist) {
                    zds += Double.valueOf(map.get("总厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("总厚度合格点数").toString());

                    double max = Double.valueOf(map.get("总厚度平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("总厚度平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("总厚度代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("总厚度代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("zxfhdlmlx","沥青路面(连接线)");
                mapzhd.put("zxfhdlb","总厚度");
                mapzhd.put("zxfhdjcds", zds);
                mapzhd.put("zxfhdhgs", hgds);
                mapzhd.put("zxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("zxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("zxfhdsjz", zlist.get(0).get("总厚度设计值"));
                mapzhd.put("zxfhdzdbz", zzhdbz);
                mapzhd.put("zxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }

        }

        //混凝土
        List<Map<String, Object>> list1 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("zxfhdlmlx","混凝土路面");
            map2.put("zxfhdlb","总厚度");
            map2.put("zxfhdpjzfw",list1.get(0).get("最大值")+"~"+list1.get(0).get("最小值"));
            map2.put("zxfhdzdbz",list1.get(0).get("代表值"));
            map2.put("zxfhdydbz",list1.get(0).get("代表值"));
            map2.put("zxfhdsjz",list1.get(0).get("设计值"));
            map2.put("zxfhdjcds",list1.get(0).get("检测点数"));
            map2.put("zxfhdhgs",list1.get(0).get("合格点数"));
            map2.put("zxfhdhgl",list1.get(0).get("合格率"));
            resultList.add(map2);
        }
        return resultList;
    }


    /**
     * 表4.1.2-8（2）  路面雷达厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getldhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.0");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> resultList = new ArrayList<>();
        Double sdMax = 0.0,ljxsdMax = 0.0,fhlmsdMax = 0.0;
        Double sdMin = Double.MAX_VALUE,ljxsdMin = Double.MAX_VALUE,fhlmsdMin = Double.MAX_VALUE;
        double jcds = 0,ljxjcds = 0,fhlmjcds = 0;
        double hgs = 0,ljxhgs = 0,fhlmhgs = 0;
        boolean a =false,b = false,c = false;
        String ljxsjz = "", fhlmsjz = "",sjz = "";
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线")){
                    a = true;
                    double max = Double.valueOf(map.get("Max").toString());
                    ljxsdMax = (max > ljxsdMax) ? max : ljxsdMax;
                    double min = Double.valueOf(map.get("Min").toString());
                    ljxsdMin = (min < ljxsdMin) ? min : ljxsdMin;

                    ljxjcds += Double.valueOf(map.get("总点数").toString());
                    ljxhgs += Double.valueOf(map.get("合格点数").toString());

                    ljxsjz = map.get("设计值").toString();
                }else {
                    if (lmlx.contains("复合路面")){
                        b = true;
                        double max = Double.valueOf(map.get("Max").toString());
                        fhlmsdMax = (max > fhlmsdMax) ? max : fhlmsdMax;
                        double min = Double.valueOf(map.get("Min").toString());
                        fhlmsdMin = (min < fhlmsdMin) ? min : fhlmsdMin;

                        fhlmjcds += Double.valueOf(map.get("总点数").toString());
                        fhlmhgs += Double.valueOf(map.get("合格点数").toString());
                        fhlmsjz = map.get("设计值").toString();
                    }else {
                        if (lmlx.equals("左幅") || lmlx.equals("右幅")){
                            c = true;
                            double max = Double.valueOf(map.get("Max").toString());
                            sdMax = (max > sdMax) ? max : sdMax;
                            double min = Double.valueOf(map.get("Min").toString());
                            sdMin = (min < sdMin) ? min : sdMin;

                            jcds += Double.valueOf(map.get("总点数").toString());
                            hgs += Double.valueOf(map.get("合格点数").toString());

                            sjz = map.get("设计值").toString();
                        }
                    }

                }
            }
            if (c){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("ldhdlb","总厚度");
                map1.put("ldhdlmlx","沥青路面(主线/匝道)");
                map1.put("ldhdsjz",sjz);
                map1.put("ldhdjcds",decf.format(jcds));
                map1.put("ldhdhgds",decf.format(hgs));
                map1.put("ldhdhgl",jcds!=0 ? df.format(hgs/jcds*100) : 0);
                map1.put("ldhdbhfw",df.format(sdMin)+"~"+df.format(sdMax));
                resultList.add(map1);
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("ldhdlb","总厚度");
                map1.put("ldhdlmlx","沥青路面(连接线)");
                map1.put("ldhdsjz",ljxsjz);
                map1.put("ldhdjcds",decf.format(ljxjcds));
                map1.put("ldhdhgds",decf.format(ljxhgs));
                map1.put("ldhdhgl",ljxjcds!=0 ? df.format(ljxhgs/ljxjcds*100) : 0);
                map1.put("ldhdbhfw",df.format(ljxsdMin)+"~"+df.format(ljxsdMax));
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("ldhdlb","总厚度");
                map1.put("ldhdlmlx","沥青路面(复合路面)");
                map1.put("ldhdsjz",fhlmsjz);
                map1.put("ldhdjcds",decf.format(fhlmjcds));
                map1.put("ldhdhgds",decf.format(fhlmhgs));
                map1.put("ldhdhgl",fhlmjcds!=0 ? df.format(fhlmhgs/fhlmjcds*100) : 0);
                map1.put("ldhdbhfw",df.format(fhlmsdMin)+"~"+df.format(fhlmsdMax));
                resultList.add(map1);
            }

        }


        return resultList;

    }

    /**
     * 表4.1.2-9 横坡检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethpData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, String>> list = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        double lmzds = 0.0;
        double lmhgds = 0.0;
        double ljxlmzds = 0.0;
        double ljxlmhgds = 0.0;

        double lmzds1 = 0.0;
        double lmhgds1 = 0.0;
        if (list!=null && list.size()>0){
            boolean a = false;
            boolean b = false,c = false;
            for (Map<String, String> map : list) {
                String jcxm = map.get("检测项目");
                String lmlx = map.get("路面类型");
                if (jcxm.contains("连接线")){
                    c = true;
                    ljxlmzds +=Double.valueOf(map.get("检测点数"));
                    ljxlmhgds +=Double.valueOf(map.get("合格点数"));
                }else {
                    if (lmlx.contains("沥青路面")){
                        a = true;
                        lmzds +=Double.valueOf(map.get("检测点数"));
                        lmhgds +=Double.valueOf(map.get("合格点数"));

                    }
                    if (lmlx.contains("混凝土路面")){
                        b = true;
                        lmzds1 +=Double.valueOf(map.get("检测点数"));
                        lmhgds1 +=Double.valueOf(map.get("合格点数"));
                    }
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("hpzds",decf.format(lmzds));
                map1.put("hphgds",decf.format(lmhgds));
                map1.put("hphgl",lmzds!=0 ? df.format(lmhgds/lmzds*100) : 0);
                map1.put("hplmlx","沥青路面(主线/匝道)");
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("hpzds",decf.format(lmzds1));
                map2.put("hphgds",decf.format(lmhgds1));
                map2.put("hphgl",lmzds1!=0 ? df.format(lmhgds1/lmzds1*100) : 0);
                map2.put("hplmlx","混凝土路面");
                resultList.add(map2);
            }
            if (c){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("hpzds",decf.format(ljxlmzds));
                map2.put("hphgds",decf.format(ljxlmhgds));
                map2.put("hphgl",ljxlmzds!=0 ? df.format(ljxlmhgds/ljxlmzds*100) : 0);
                map2.put("hplmlx","沥青路面(连接线)");
                resultList.add(map2);
            }
        }
        return resultList;

    }


    /**
     * 表4.1.2-7  抗滑检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getkhData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhGzsdService.lookJdbjg(commonInfoVo, 1);//自动化
        String sjz = "",ljxsjz = "";
        double jcds = 0.0;
        double hgds = 0.0;
        double ljxjcds = 0.0;
        double ljxhgds = 0.0;

        if (list!=null && list.size()>0){
            boolean a = false,b = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();
                if (lmlx.contains("路面") && !jcxm.contains("连接线")){
                    a = true;
                    sjz = map.get("设计值").toString();
                    jcds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
                if (jcxm.contains("连接线")){
                    b = true;
                    ljxsjz = map.get("设计值").toString();
                    ljxjcds += Double.valueOf(map.get("总点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;

                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","沥青路面（主线/匝道）");
                map1.put("khkpzb","构造深度");
                map1.put("khsjz",sjz);
                map1.put("khbhfw",lmMin+"~"+lmMax);
                map1.put("khjcds",decf.format(jcds));
                map1.put("khhgds",decf.format(hgds));
                map1.put("khhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","沥青路面(连接线)");
                map1.put("khkpzb","构造深度");
                map1.put("khsjz",ljxsjz);
                map1.put("khbhfw",ljxlmMin+"~"+ljxlmMax);
                map1.put("khjcds",decf.format(ljxjcds));
                map1.put("khhgds",decf.format(ljxhgds));
                map1.put("khhgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list1 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (list1!=null && list1.size()>0){
            String gdz = "";
            double szds = 0;
            double shgs = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("混凝土路面")){
                    a = true;
                    szds += Double.valueOf(map.get("检测点数").toString());
                    shgs += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                }else {
                    continue;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","混凝土路面");
                map1.put("khkpzb","构造深度");
                map1.put("khsjz",gdz);
                map1.put("khbhfw",df.format(lmMin)+"~"+df.format(lmMax));
                map1.put("khjcds",decf.format(szds));
                map1.put("khhgds",decf.format(shgs));
                map1.put("khhgl",szds!=0 ? df.format(shgs/szds*100) : 0);
                resultList.add(map1);
            }

        }
        List<Map<String, Object>> list2 = jjgZdhMcxsService.lookJdbjg(commonInfoVo,1);
        if (list2!=null && list2.size()>0){
            boolean a = false,b = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Double ljxlmMax = 0.0;
            Double ljxlmMin = Double.MAX_VALUE;
            double zds = 0;
            double hgs = 0;
            double ljxzds = 0;
            double ljxhgs = 0;
            String sjz2 = "",mcxsljxsjz = "";
            for (Map<String, Object> map : list2) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("路面") && !jcxm.contains("连接线")){
                    a = true;
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgs += Double.valueOf(map.get("合格点数").toString());
                    sjz2 = map.get("设计值").toString();

                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
                if (jcxm.contains("连接线")){
                    b = true;
                    mcxsljxsjz = map.get("设计值").toString();
                    ljxzds += Double.valueOf(map.get("总点数").toString());
                    ljxhgs += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("Max").toString());
                    ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;

                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","沥青路面(主线/匝道)");
                map1.put("khkpzb","摩擦系数");
                map1.put("khsjz",sjz2);
                map1.put("khbhfw",df.format(lmMin)+"~"+df.format(lmMax));
                map1.put("khjcds",decf.format(zds));
                map1.put("khhgds",decf.format(hgs));
                map1.put("khhgl",zds!=0 ? df.format(hgs/zds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","沥青路面（连接线）");
                map1.put("khkpzb","摩擦系数");
                map1.put("khsjz",mcxsljxsjz);
                map1.put("khbhfw",df.format(ljxlmMin)+"~"+df.format(ljxlmMax));
                map1.put("khjcds",decf.format(ljxzds));
                map1.put("khhgds",decf.format(ljxhgs));
                map1.put("khhgl",ljxzds!=0 ? df.format(ljxhgs/ljxzds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;
    }


    /**
     * 表4.1.2-6  平整度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getpzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.0");
        List<Map<String, Object>> list = jjgZdhPzdService.lookJdbjg(commonInfoVo, 1);
        if (list!=null && list.size()>0){
            boolean a= false,b = false;
            double zdzds = 0,zdhgds = 0;
            double zxz = 0,zdz = 0;
            double hntzxz = 0,hntzdz = 0;
            double hntzds = 0,hnthgds = 0;
            String zdlmlx = "", zdsjz = "";
            String hntlmlx = "", hntsjz = "";
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                String jcxm = map.get("检测项目").toString();

                if (lmlx.contains("混凝土")){
                    if (!b){
                        hntzxz = Double.valueOf(map.get("Min").toString());
                        hntzdz = Double.valueOf(map.get("Max").toString());
                    }
                    b = true;
                    hntzds += Double.valueOf(map.get("总点数").toString());
                    hnthgds += Double.valueOf(map.get("合格点数").toString());
                    hntlmlx = lmlx;
                    hntsjz = map.get("设计值").toString();

                    double zxz1 = Double.valueOf(map.get("Min").toString());
                    double zdz1 = Double.valueOf(map.get("Max").toString());
                    if (zxz1<hntzxz){
                        hntzxz = zxz1;
                    }
                    if (zdz1>hntzdz){
                        hntzdz = zdz1;
                    }
                   /* Map map1 = new HashMap();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("pzdzb","IRI");
                    map1.put("lmlx",lmlx);
                    map1.put("pzdlmlx","混凝土路面");
                    map1.put("pzdgdz",map.get("设计值"));
                    map1.put("pzdjcds",map.get("总点数"));
                    map1.put("pzdhgds",map.get("合格点数"));
                    map1.put("pzdhgl",map.get("合格率"));
                    map1.put("pzdbhfw",map.get("Min")+"~"+map.get("Max"));
                    resultList.add(map1);*/
                }else {
                    if (jcxm.contains("连接线")){
                        Map map1 = new HashMap();
                        map1.put("htd",commonInfoVo.getHtd());
                        map1.put("pzdzb","IRI");
                        map1.put("lmlx",lmlx);
                        map1.put("pzdlmlx","沥青路面（连接线）");
                        map1.put("pzdgdz",map.get("设计值"));
                        map1.put("pzdjcds",map.get("总点数"));
                        map1.put("pzdhgds",map.get("合格点数"));
                        map1.put("pzdhgl",map.get("合格率"));
                        map1.put("pzdbhfw",map.get("Min")+"~"+map.get("Max"));
                        resultList.add(map1);
                    }else if (!jcxm.contains("主线") && !jcxm.contains("连接线")){
                        if (lmlx.equals("沥青匝道")){
                            if (!a){
                                zxz = Double.valueOf(map.get("Min").toString());
                                zdz = Double.valueOf(map.get("Max").toString());
                            }
                            //匝道
                            a = true;
                            zdzds += Double.valueOf(map.get("总点数").toString());
                            zdhgds += Double.valueOf(map.get("合格点数").toString());
                            zdlmlx = lmlx;
                            zdsjz = map.get("设计值").toString();


                            double zxz1 = Double.valueOf(map.get("Min").toString());
                            double zdz1 = Double.valueOf(map.get("Max").toString());
                            if (zxz1<zxz){
                                zxz = zxz1;
                            }
                            if (zdz1>zdz){
                                zdz = zdz1;
                            }
                        }
                    } else {
                        if (lmlx.equals("沥青路面")){
                            Map map1 = new HashMap();
                            map1.put("htd",commonInfoVo.getHtd());
                            map1.put("pzdzb","IRI");
                            map1.put("lmlx",lmlx);
                            map1.put("pzdlmlx","沥青路面（主线）");
                            map1.put("pzdgdz",map.get("设计值"));
                            map1.put("pzdjcds",map.get("总点数"));
                            map1.put("pzdhgds",map.get("合格点数"));
                            map1.put("pzdhgl",map.get("合格率"));
                            map1.put("pzdbhfw",map.get("Min")+"~"+map.get("Max"));
                            resultList.add(map1);
                        }

                    }
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("pzdzb","IRI");
                map1.put("lmlx",zdlmlx);
                map1.put("pzdlmlx","沥青路面（匝道）");
                map1.put("pzdgdz",zdsjz);
                map1.put("pzdjcds",zdzds);
                map1.put("pzdhgds",zdhgds);
                map1.put("pzdbhfw",df.format(zxz)+"~"+df.format(zdz));
                map1.put("pzdhgl",zdzds!=0 ? df.format(zdhgds/zdzds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("pzdzb","IRI");
                map1.put("lmlx",hntlmlx);
                map1.put("pzdlmlx","混凝土路面");
                map1.put("pzdgdz",hntsjz);
                map1.put("pzdjcds",df.format(hntzds));
                map1.put("pzdhgds",df.format(hnthgds));
                map1.put("pzdbhfw",df.format(hntzxz)+"~"+df.format(hntzdz));
                map1.put("pzdhgl",hntzds!=0 ? df.format(hnthgds/hntzds*100) : 0);
                resultList.add(map1);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-5  混凝土路面相邻板高差检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethntxlbgcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.0");
        List<Map<String, Object>> list = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                map.put("htd",commonInfoVo.getHtd());
                map.put("hntxlbgczds",map.get("总点数"));
                map.put("hntxlbgchgds",map.get("合格点数"));
                map.put("hntxlbgchgl",map.get("合格率"));
                map.put("hntxlbgcgdz",map.get("规定值"));
                double zdz = Double.valueOf(map.get("max").toString());
                map.put("hntxlbgcmax",df.format(zdz));
                double zxz = Double.valueOf(map.get("min").toString());
                map.put("hntxlbgcmin",df.format(zxz));
                resultList.add(map);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-4  混凝土路面强度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethntqdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                map.put("htd",commonInfoVo.getHtd());
                map.put("hntqdzds",map.get("总点数"));
                map.put("hntqdhgds",map.get("合格点数"));
                map.put("hntqdhgl",map.get("合格率"));
                map.put("hntqdgdz",map.get("规定值"));
                double zdz = Double.valueOf(map.get("最大值").toString());
                map.put("hntqdmax",decf.format(zdz));
                double zxz = Double.valueOf(map.get("最小值").toString());
                map.put("hntqdmin",decf.format(zxz));
                map.put("hntqdpjz",map.get("平均值"));

                resultList.add(map);
            }
        }
        return resultList;
    }


    /**
     * 表4.1.2-3  渗水系数检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmssxsData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.0");
        DecimalFormat decf = new DecimalFormat("0.#");
        List<Map<String, Object>> list = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo,1);
        double jcds =0.0;
        double hgds =0.0;
        double ljxjcds =0.0;
        double ljxhgds =0.0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        Double ljxMax = 0.0;
        Double ljxMin = Double.MAX_VALUE;
        boolean a= false,b = false;
        String lmgdz = "",ljxgdz = "";
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                String lmlx = map.get("检测项目").toString();
                if (lmlx.contains("沥青路面") || lmlx.contains("匝道路面")){
                    a = true;
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                    lmgdz = map.get("规定值").toString();
                }
                if (lmlx.contains("连接线")){
                    b = true;
                    ljxjcds += Double.valueOf(map.get("检测点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("最大值").toString());
                    ljxMax = (max > ljxMax) ? max : ljxMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    ljxMin = (min < ljxMin) ? min : ljxMin;
                    ljxgdz = map.get("规定值").toString();
                }
            }
            if (a){
                Map<String,Object> map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("lxmc","主线/匝道");
                map.put("lmssxsssjcds",jcds);
                map.put("lmssxsmax",decf.format(lmMax));
                map.put("lmssxsmin",decf.format(lmMin));
                map.put("lmssxssshgds",hgds);
                map.put("lmssxssshgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                map.put("lmssxsgdz",lmgdz);
                resultList.add(map);
            }
            if (b){
                Map<String,Object> map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("lxmc","连接线");
                map.put("lmssxsssjcds",ljxjcds);
                map.put("lmssxsmax",decf.format(ljxMax));
                map.put("lmssxsmin",decf.format(ljxMin));
                map.put("lmssxssshgds",ljxhgds);
                map.put("lmssxssshgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
                map.put("lmssxsgdz",ljxgdz);
                resultList.add(map);
            }
        }
        return resultList;

    }


    /**
     * 表4.1.2-6  车辙检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getczData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.0");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhCzService.lookJdbjg(commonInfoVo, 1);
        if (list!=null && list.size()>0){
            double zds =0.0,ljxzds = 0;
            double hgds =0.0,ljxhgds = 0;
            Double lmMax = 0.0,ljxlmMax = 0.0;
            Double lmMin = Double.MAX_VALUE,ljxlmMin = Double.MAX_VALUE;
            String sjz = "",ljxsjz = "";
            boolean a = false,b = false;
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                String lmlx = map.get("路面类型").toString();
                if (jcxm.contains("连接线")){
                    a = true;
                    ljxsjz = map.get("设计值").toString();
                    ljxzds += Double.valueOf(map.get("总点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    ljxlmMax = (max > ljxlmMax) ? max : ljxlmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    ljxlmMin = (min < ljxlmMin) ? min : ljxlmMin;
                }else {
                    if (lmlx.contains("路面")){
                        b = true;
                        sjz = map.get("设计值").toString();
                        zds += Double.valueOf(map.get("总点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());

                        double max = Double.valueOf(map.get("Max").toString());
                        lmMax = (max > lmMax) ? max : lmMax;

                        double min = Double.valueOf(map.get("Min").toString());
                        lmMin = (min < lmMin) ? min : lmMin;
                    }
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("czzb","RDI");
                map1.put("czlmlx","沥青路面(连接线)");
                map1.put("czgdz",ljxsjz);
                map1.put("czbhfw",ljxlmMin+"~"+ljxlmMax);
                map1.put("czzds",decf.format(ljxzds));
                map1.put("czhgds",decf.format(ljxhgds));
                map1.put("czhgl",ljxzds!=0 ? df.format(ljxhgds/ljxzds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("czzb","RDI");
                map1.put("czlmlx","沥青路面(主线/匝道)");
                map1.put("czgdz",sjz);
                map1.put("czbhfw",lmMin+"~"+lmMax);
                map1.put("czzds",decf.format(zds));
                map1.put("czhgds",decf.format(hgds));
                map1.put("czhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;
    }


    /**
     * 表4.1.2-2  沥青路面弯沉检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmwcallData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.#");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> listlcf = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo,1);
        List<Map<String, Object>> listbrml = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo,1);
        if (listlcf != null || listbrml != null){
            if (listlcf != null && listbrml != null) {
                listlcf.addAll(listbrml);
            }else if (listlcf == null && listbrml != null) {
                // 如果listlcf为空而listbrml不为空，你可以选择将listbrml赋值给listlcf
                // 或者根据业务需求做其他处理
                listlcf = listbrml;
            }

            Map<String, List<Map<String, Object>>> result = listlcf.stream()
                    .collect(Collectors.groupingBy(map -> (String) map.get("规定值")));
            result.forEach((group, grouphtdData) -> {
                double zds = 0.0,hgds = 0.0;
                for (Map<String, Object> grouphtdDatum : grouphtdData) {
                    zds += Double.valueOf(grouphtdDatum.get("检测单元数").toString());
                    hgds += Double.valueOf(grouphtdDatum.get("合格单元数").toString());
                }
                Map map = new HashMap();
                map.put("htd",commonInfoVo.getHtd());
                map.put("lmwcdbz",group);
                map.put("lmwcmax",decf.format(findMaxValue(grouphtdData, "max")));
                map.put("lmwcmin",decf.format(findMinValue(grouphtdData, "min")));
                map.put("lmwczds",df.format(zds));
                map.put("lmwchgds",df.format(hgds));
                map.put("lmwchgl",(zds != 0) ? df.format(hgds/zds*100) : "0");
                resultList.add(map);
            });
        }

        /**
         * [{listtemp=[{wcjl=合格, yswcz=10.0, wcdbz=6.273656531157691}, {wcjl=合格, yswcz=10.0, wcdbz=6.338587958229803}], 合格单元数=2, 规定值=10.0, 检测单元数=2, 合格率=100.00, 代表值=[6.273656531157691, 6.273656531157691]}]
         * [{listtemp=[{wcjl=合格, yswcz=10.0, wcdbz=7.952}, {wcjl=合格, yswcz=10.0, wcdbz=8.09312}], 合格单元数=2, 规定值=10.0, 检测单元数=2, 合格率=100.00, 代表值=[7.952, 7.952]}]
         */
        /*if (listlcf!=null && listlcf.size()>0){
            for (Map<String, Object> map : listlcf) {

                List<String,Object> retrievedList = map.get("listtemp");
                String s = map.get("listtemp").toString();


            }
        }*/


        /*double zds =0;
        double hgds =0;
        boolean a =false,b=false;
        if (listlcf!=null && listlcf.size()>0){
            a = true;
            for (Map<String, Object> map : listlcf) {
                zds += Double.valueOf(map.get("检测单元数").toString());
                hgds += Double.valueOf(map.get("合格单元数").toString());
            }
           *//* Map map = new HashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("lmwcdbz",listlcf.get(0).get("规定值"));
            map.put("lmwcmax",getmaxvalue(listlcf.get(0).get("代表值").toString()));
            map.put("lmwcmin",getminvalue(listlcf.get(0).get("代表值").toString()));
            map.put("lmwczds",list.get(0).get("检测单元数"));
            map.put("lmwchgds",list.get(0).get("合格单元数"));
            map.put("lmwchgl",list.get(0).get("合格率"));
            resultList.add(map);*//*
        }
        if (listbrml!=null && listbrml.size()>0) {
            b = true;
            for (Map<String, Object> map : listbrml) {
                zds += Double.valueOf(map.get("检测单元数").toString());
                hgds += Double.valueOf(map.get("合格单元数").toString());
            }
        }



        if (a || b){
            Map map = new HashMap();
            map.put("htd",commonInfoVo.getHtd());
            if (a && b){
                if (listlcf !=null && listlcf.size()>0){

                }
                if (listlcf !=null && listlcf.size()>0){

                }

            }
            if (a && !b){
                map.put("lmwcdbz",listlcf.get(0).get("规定值"));
                map.put("lmwcmax",getmaxvalue(listlcf.get(0).get("代表值").toString()));
                map.put("lmwcmin",getminvalue(listlcf.get(0).get("代表值").toString()));
            }
            if (b && !a){
                map.put("lmwcdbz",listbrml.get(0).get("规定值"));
                map.put("lmwcmax",getmaxvalue(listbrml.get(0).get("代表值").toString()));
                map.put("lmwcmin",getminvalue(listbrml.get(0).get("代表值").toString()));
            }

            map.put("lmwczds",zds);
            map.put("lmwchgds",hgds);
            map.put("lmwchgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultList.add(map);
        }*/
        return resultList;

    }

    // 查找最大值
    public static double findMaxValue(List<Map<String, Object>> list, String key) {
        double maxValue = Double.valueOf(list.get(0).get("max").toString()); // 初始化为Double的最小值
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
        double minValue = Double.valueOf(list.get(0).get("min").toString()); // 初始化为Double的最大值
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
     * 表4.1.2-2 弯沉(落锤法)
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getlmwclcfData(CommonInfoVo commonInfoVo) throws IOException {

        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo,1);
        if (list!=null && list.size()>0){
            for (Map<String, Object> stringObjectMap : list) {
                Map map = new HashMap();
                map.put("htd",commonInfoVo.getHtd());
                map.put("lmwclcfdbz",stringObjectMap.get("规定值"));
                map.put("lmwclcfmax",stringObjectMap.get("max"));
                map.put("lmwclcfmin",stringObjectMap.get("min"));
                map.put("lmwclcfzds",stringObjectMap.get("检测单元数"));
                map.put("lmwclcfhgds",stringObjectMap.get("合格单元数"));
                map.put("lmwclcfhgl",stringObjectMap.get("合格率"));
                resultList.add(map);
            }

        }
        return resultList;

    }


    /**
     * 表4.1.2-2 弯沉(贝尔曼梁法)
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmwcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo,1);
        if (list!=null && list.size()>0){
            for (Map<String, Object> stringObjectMap : list) {
                Map map = new HashMap();
                map.put("htd",commonInfoVo.getHtd());
                map.put("lmwcdbz",stringObjectMap.get("规定值"));
                map.put("lmwcmax",stringObjectMap.get("max"));
                map.put("lmwcmin",stringObjectMap.get("min"));
                map.put("lmwczds",stringObjectMap.get("检测单元数"));
                map.put("lmwchgds",stringObjectMap.get("合格单元数"));
                map.put("lmwchgl",stringObjectMap.get("合格率"));
                resultList.add(map);
            }
        }
        return resultList;
    }


    private String getminvalue(String myList) {
        DecimalFormat df = new DecimalFormat("0.00");
        myList = myList.replace("[", "").replace("]", ""); // 去除方括号
        String[] valus = myList.split(",");
        double min = Integer.MAX_VALUE; // 初始化最大值为最小整数
        for (String num : valus) {
            Double n = Double.valueOf(num);
            if (n < min) {
                min = n;
            }
        }
        return String.valueOf(df.format(min));
    }


    private String getmaxvalue(String myList) {
        DecimalFormat df = new DecimalFormat("0.00");
        myList = myList.replace("[", "").replace("]", ""); // 去除方括号
        String[] valus = myList.split(",");
        double max = Integer.MIN_VALUE; // 初始化最大值为最小整数
        for (String num : valus) {
            Double n = Double.valueOf(num);
            if (n > max) {
                max = n;
            }
        }
        return String.valueOf(df.format(max));
    }

    /**
     * 表4.1.2-1  沥青路面面层压实度检测结果
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmysdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        System.out.println(list);

        /**
         * [{密度规定值=98.0, 标准=标准密度, 合格点数=0, 实测值变化范围=98.8~99.6, 检测点数=2, 密度代表值=96.7, 合格率=0.00, 路面类型=沥青路面压实度匝道-上面层},
         * {密度规定值=94.0, 标准=最大理论密度, 合格点数=0, 实测值变化范围=95.0~95.8, 检测点数=2, 密度代表值=92.9, 合格率=0.00, 路面类型=沥青路面压实度匝道-上面层},
         * {密度规定值=96.0, 标准=标准密度, 合格点数=4, 实测值变化范围=95.7~98.4, 检测点数=4, 密度代表值=96.1, 合格率=100.00, 路面类型=沥青路面压实度匝道-中下面层},
         * {密度规定值=92.0, 标准=最大理论密度, 合格点数=4, 实测值变化范围=91.6~94.1, 检测点数=4, 密度代表值=92.0, 合格率=100.00, 路面类型=沥青路面压实度匝道-中下面层},
         *
         * {密度规定值=98.0, 标准=标准密度, 合格点数=2, 实测值变化范围=99.6~99.6, 检测点数=2, 密度代表值=99.6, 合格率=100.00, 路面类型=隧道右幅-上面层},
         * {密度规定值=94.0, 标准=最大理论密度, 合格点数=2, 实测值变化范围=95.8~95.8, 检测点数=2, 密度代表值=95.8, 合格率=100.00, 路面类型=隧道右幅-上面层},
         * {密度规定值=96.0, 标准=标准密度, 合格点数=0, 实测值变化范围=95.7~98.4, 检测点数=4, 密度代表值=95.2, 合格率=0.00, 路面类型=隧道右幅-中下面层},
         * {密度规定值=92.0, 标准=最大理论密度, 合格点数=0, 实测值变化范围=91.6~94.1, 检测点数=4, 密度代表值=91.2, 合格率=0.00, 路面类型=隧道右幅-中下面层},
         * {密度规定值=98.0, 标准=标准密度, 合格点数=2, 实测值变化范围=98.8~98.8, 检测点数=2, 密度代表值=98.8, 合格率=100.00, 路面类型=隧道左幅-上面层},
         * {密度规定值=94.0, 标准=最大理论密度, 合格点数=2, 实测值变化范围=95.0~95.0, 检测点数=2, 密度代表值=95.0, 合格率=100.00, 路面类型=隧道左幅-上面层},
         * {密度规定值=96.0, 标准=标准密度, 合格点数=4, 实测值变化范围=97.8~98.0, 检测点数=4, 密度代表值=97.8, 合格率=100.00, 路面类型=隧道左幅-中下面层},
         * {密度规定值=92.0, 标准=最大理论密度, 合格点数=4, 实测值变化范围=93.9~93.9, 检测点数=4, 密度代表值=93.9, 合格率=100.00, 路面类型=隧道左幅-中下面层},
         *
         * {密度规定值=98.0, 标准=标准密度, 合格点数=3, 实测值变化范围=99.6~99.6, 检测点数=3, 密度代表值=99.6, 合格率=100.00, 路面类型=沥青路面压实度右幅-上面层},
         * {密度规定值=96.0, 标准=标准密度, 合格点数=0, 实测值变化范围=95.7~98.4, 检测点数=6, 密度代表值=95.8, 合格率=0.00, 路面类型=沥青路面压实度右幅-中下面层},
         * {密度规定值=98.0, 标准=标准密度, 合格点数=3, 实测值变化范围=98.8~98.8, 检测点数=3, 密度代表值=98.8, 合格率=100.00, 路面类型=沥青路面压实度左幅-上面层},
         * {密度规定值=96.0, 标准=标准密度, 合格点数=6, 实测值变化范围=97.8~98.0, 检测点数=6, 密度代表值=97.8, 合格率=100.00, 路面类型=沥青路面压实度左幅-中下面层},
         *
         *
         *
         * {密度规定值=94.0, 标准=最大理论密度, 合格点数=3, 实测值变化范围=95.8~95.8, 检测点数=3, 密度代表值=95.8, 合格率=100.00, 路面类型=沥青路面压实度右幅-上面层},
         *
         * {密度规定值=92.0, 标准=最大理论密度, 合格点数=0, 实测值变化范围=91.6~94.1, 检测点数=6, 密度代表值=91.7, 合格率=0.00, 路面类型=沥青路面压实度右幅-中下面层},
         *
         * {密度规定值=94.0, 标准=最大理论密度, 合格点数=3, 实测值变化范围=95.0~95.0, 检测点数=3, 密度代表值=95.0, 合格率=100.00, 路面类型=沥青路面压实度左幅-上面层},
         *
         * {密度规定值=92.0, 标准=最大理论密度, 合格点数=6, 实测值变化范围=93.9~93.9, 检测点数=6, 密度代表值=93.9, 合格率=100.00, 路面类型=沥青路面压实度左幅-中下面层},
         *
         *
         * {密度规定值=98.0, 标准=标准密度, 合格点数=0, 实测值变化范围=98.8~99.6, 检测点数=2, 密度代表值=96.7, 合格率=0.00, 路面类型=连接线-上面层},
         * {密度规定值=94.0, 标准=最大理论密度, 合格点数=0, 实测值变化范围=95.0~95.8, 检测点数=2, 密度代表值=92.9, 合格率=0.00, 路面类型=连接线-上面层},
         * {密度规定值=96.0, 标准=标准密度, 合格点数=4, 实测值变化范围=95.7~98.4, 检测点数=4, 密度代表值=96.1, 合格率=100.00, 路面类型=连接线-中下面层},
         * {密度规定值=92.0, 标准=最大理论密度, 合格点数=4, 实测值变化范围=91.6~94.1, 检测点数=4, 密度代表值=92.0, 合格率=100.00, 路面类型=连接线-中下面层}]
         */


        if (list!=null && list.size()>0){
            double zxbzjcds = 0,zxbzhgds = 0,zxzdjcds = 0,zxzdhgds = 0;
            String bzmdgdz = "",bzsczbhfw="",zmddbz="",ymddbz="",zdmdgdz = "",zdsczbhfw="",zdzmddbz="",zdymddbz="";
            boolean a = false ,b = false;


            List<Map<String, Object>> bzmdzxList = new ArrayList<>();
            List<Map<String, Object>> bzmdzdList = new ArrayList<>();
            List<Map<String, Object>> bzmdljxList = new ArrayList<>();

            List<Map<String, Object>> zdmdzxList = new ArrayList<>();
            List<Map<String, Object>> zdmdzdList = new ArrayList<>();
            List<Map<String, Object>> zdmdljxList = new ArrayList<>();
            for (Map<String, Object> map : list) {
                String lx = map.get("路面类型").toString();
                String bz = map.get("标准").toString();
                if (bz.equals("标准密度")){
                    //主线
                    if (lx.contains("沥青路面压实度左幅") || lx.contains("沥青路面压实度右幅")){
                        bzmdzxList.add(map);
                        /*a  = true;
                        zxbzjcds += Double.valueOf(map.get("检测点数").toString());
                        zxbzhgds += Double.valueOf(map.get("合格点数").toString());
                        bzmdgdz = map.get("密度规定值").toString();
                        bzsczbhfw = map.get("实测值变化范围").toString();
                        if (lx.contains("沥青路面压实度左幅")){
                            zmddbz = map.get("密度代表值").toString();

                        }else if (lx.contains("沥青路面压实度右幅")){
                            ymddbz = map.get("密度代表值").toString();
                        }*/
                    }//匝道
                    else if (lx.contains("沥青路面压实度匝道")){
                        bzmdzdList.add(map);
                        /*double zdzxbzjcds = Double.valueOf(map.get("检测点数").toString());
                        double zdzxbzhgds = Double.valueOf(map.get("合格点数").toString());
                        double hgl = Double.valueOf(map.get("合格率").toString());
                        String zdbzmdgdz = map.get("密度规定值").toString();
                        String zdbzsczbhfw = map.get("实测值变化范围").toString();
                        String zdmddbz = map.get("密度代表值").toString();

                        Map map2 = new HashMap();
                        map2.put("htd",commonInfoVo.getHtd());
                        map2.put("ysdlx","匝道");
                        map2.put("bz","试验室标准密度");
                        map2.put("bzsczbhfw",zdbzsczbhfw);
                        map2.put("zmddbz",zdmddbz);
                        map2.put("ymddbz",zdmddbz);
                        map2.put("bzmdgdz",zdbzmdgdz);
                        map2.put("zxbzjcds",zdzxbzjcds);
                        map2.put("zxbzhgds",zdzxbzhgds);
                        map2.put("hgl",hgl);
                        resultList.add(map2);*/

                    }//连接线
                    else if (lx.contains("连接线")){
                        bzmdljxList.add(map);
                        /*double zdzxbzjcds = Double.valueOf(map.get("检测点数").toString());
                        double zdzxbzhgds = Double.valueOf(map.get("合格点数").toString());
                        double hgl = Double.valueOf(map.get("合格率").toString());
                        String zdbzmdgdz = map.get("密度规定值").toString();
                        String zdbzsczbhfw = map.get("实测值变化范围").toString();
                        String zdmddbz = map.get("密度代表值").toString();

                        Map map2 = new HashMap();
                        map2.put("htd",commonInfoVo.getHtd());
                        map2.put("ysdlx","连接线");
                        map2.put("bz","试验室标准密度");
                        map2.put("bzsczbhfw",zdbzsczbhfw);
                        map2.put("zmddbz",zdmddbz);
                        map2.put("ymddbz",zdmddbz);
                        map2.put("bzmdgdz",zdbzmdgdz);
                        map2.put("zxbzjcds",zdzxbzjcds);
                        map2.put("zxbzhgds",zdzxbzhgds);
                        map2.put("hgl",hgl);
                        resultList.add(map2);*/

                    }
                }else if (bz.equals("最大理论密度")){
                    if (lx.contains("沥青路面压实度左幅") || lx.contains("沥青路面压实度右幅")){
                        zdmdzxList.add(map);
                       /* b  = true;
                        zxzdjcds += Double.valueOf(map.get("检测点数").toString());
                        zxzdhgds += Double.valueOf(map.get("合格点数").toString());
                        zdmdgdz = map.get("密度规定值").toString();
                        zdsczbhfw = map.get("实测值变化范围").toString();
                        if (lx.contains("沥青路面压实度左幅")){
                            zdzmddbz = map.get("密度代表值").toString();

                        }else if (lx.contains("沥青路面压实度右幅")){
                            zdymddbz = map.get("密度代表值").toString();
                        }*/
                    }else if (lx.contains("沥青路面压实度匝道")){
                        zdmdzdList.add(map);
                        /*double zdzxbzjcds = Double.valueOf(map.get("检测点数").toString());
                        double zdzxbzhgds = Double.valueOf(map.get("合格点数").toString());
                        double hgl = Double.valueOf(map.get("合格率").toString());
                        String zdbzmdgdz = map.get("密度规定值").toString();
                        String zdbzsczbhfw = map.get("实测值变化范围").toString();
                        String zdmddbz = map.get("密度代表值").toString();

                        Map map2 = new HashMap();
                        map2.put("htd",commonInfoVo.getHtd());
                        map2.put("ysdlx","匝道");
                        map2.put("bz","最大理论密度");
                        map2.put("bzsczbhfw",zdbzsczbhfw);
                        map2.put("zmddbz",zdmddbz);
                        map2.put("ymddbz",zdmddbz);
                        map2.put("bzmdgdz",zdbzmdgdz);
                        map2.put("zxbzjcds",zdzxbzjcds);
                        map2.put("zxbzhgds",zdzxbzhgds);
                        map2.put("hgl",hgl);
                        resultList.add(map2);*/


                    }else if (lx.contains("连接线")){
                        zdmdljxList.add(map);
                        /*double zdzxbzjcds = Double.valueOf(map.get("检测点数").toString());
                        double zdzxbzhgds = Double.valueOf(map.get("合格点数").toString());
                        double hgl = Double.valueOf(map.get("合格率").toString());
                        String zdbzmdgdz = map.get("密度规定值").toString();
                        String zdbzsczbhfw = map.get("实测值变化范围").toString();
                        String zdmddbz = map.get("密度代表值").toString();

                        Map map2 = new HashMap();
                        map2.put("htd",commonInfoVo.getHtd());
                        map2.put("ysdlx","连接线");
                        map2.put("bz","最大理论密度");
                        map2.put("bzsczbhfw",zdbzsczbhfw);
                        map2.put("zmddbz",zdmddbz);
                        map2.put("ymddbz",zdmddbz);
                        map2.put("bzmdgdz",zdbzmdgdz);
                        map2.put("zxbzjcds",zdzxbzjcds);
                        map2.put("zxbzhgds",zdzxbzhgds);
                        map2.put("hgl",hgl);
                        resultList.add(map2);*/

                    }
                }
            }
            if (bzmdzxList!=null && bzmdzxList.size()>0){
                List<Map<String, Object>> extractedmethod = extractedmethod(commonInfoVo, bzmdzxList, df,"主线","试验室标准密度");
                resultList.addAll(extractedmethod);
            }

            if (bzmdzdList!=null && bzmdzdList.size()>0){
                List<Map<String, Object>> extractedmethod = extractedmethod(commonInfoVo, bzmdzdList, df,"匝道","试验室标准密度");
                resultList.addAll(extractedmethod);
            }

            if (bzmdljxList!=null && bzmdljxList.size()>0){
                List<Map<String, Object>> extractedmethod = extractedmethod(commonInfoVo, bzmdljxList, df,"连接线","试验室标准密度");
                resultList.addAll(extractedmethod);
            }

            if (zdmdzxList!=null && zdmdzxList.size()>0){
                List<Map<String, Object>> extractedmethod = extractedmethod(commonInfoVo, zdmdzxList, df,"主线","最大理论密度");
                resultList.addAll(extractedmethod);
            }

            if (zdmdzdList!=null && zdmdzdList.size()>0){
                List<Map<String, Object>> extractedmethod = extractedmethod(commonInfoVo, zdmdzdList, df,"匝道","最大理论密度");
                resultList.addAll(extractedmethod);
            }

            if (zdmdljxList!=null && zdmdljxList.size()>0){
                List<Map<String, Object>> extractedmethod = extractedmethod(commonInfoVo, zdmdljxList, df,"连接线","最大理论密度");
                resultList.addAll(extractedmethod);
            }



            /*if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("ysdlx","主线");
                map1.put("bz","试验室标准密度");
                map1.put("bzsczbhfw",bzsczbhfw);
                map1.put("zmddbz",zmddbz);
                map1.put("ymddbz",ymddbz);
                map1.put("bzmdgdz",bzmdgdz);
                map1.put("zxbzjcds",zxbzjcds);
                map1.put("zxbzhgds",zxbzhgds);
                map1.put("hgl",zxbzjcds!=0 ? df.format(zxbzhgds/zxbzjcds*100) : 0);
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("ysdlx","主线");
                map2.put("bz","最大理论密度");
                map2.put("bzsczbhfw",zdsczbhfw);
                map2.put("zmddbz",zdzmddbz);
                map2.put("ymddbz",zdymddbz);
                map2.put("bzmdgdz",zdmdgdz);
                map2.put("zxbzjcds",zxzdjcds);
                map2.put("zxbzhgds",zxzdhgds);
                map2.put("hgl",zxzdjcds!=0 ? df.format(zxzdhgds/zxzdjcds*100) : 0);
                resultList.add(map2);
            }*/
        }
        /*Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        String gdz = "";
        String zdbz = "";
        String ydbz = "";
        Double jcds = 0.0;
        Double hgds = 0.0;

        Double ljxMax = 0.0;
        Double ljxMin = Double.MAX_VALUE;
        String ljxgdz = "";
        String ljxdbz = "";
        Double ljxjcds = 0.0;
        Double ljxhgds = 0.0;
        boolean a = false;
        boolean b = false;
        if (list!=null && list.size()>0){
            for (Map<String, Object> map : list) {
                String lx = map.get("路面类型").toString();
                if (lx.contains("沥青路面压实度左幅") || lx.contains("沥青路面压实度右幅")){
                    a = true;
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                    if (lx.contains("沥青路面压实度左幅")){
                        zdbz = map.get("代表值").toString();
                        jcds += Double.valueOf(map.get("检测点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lx.contains("沥青路面压实度右幅")){
                        ydbz = map.get("代表值").toString();
                        jcds += Double.valueOf(map.get("检测点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }

                }else if (lx.contains("沥青路面压实度匝道")){
                    Map map1 = new HashMap<>();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("ysdlx","匝道");
                    map1.put("ysdsczbh",map.get("最小值").toString()+"~"+map.get("最大值").toString());
                    map1.put("ysdzdbz",map.get("代表值").toString());
                    map1.put("ysdydbz",map.get("代表值").toString());
                    map1.put("ysdgdz",map.get("规定值").toString());
                    map1.put("ysdzs",map.get("检测点数").toString());
                    map1.put("ysdhgs",map.get("合格点数").toString());
                    map1.put("ysdhgl",map.get("合格率").toString());
                    resultList.add(map1);

                }else if (lx.contains("连接线")){
                    b = true;
                    double max = Double.valueOf(map.get("最大值").toString());
                    ljxMax = (max > ljxMax) ? max : ljxMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    ljxMin = (min < ljxMin) ? min : ljxMin;
                    ljxgdz = map.get("规定值").toString();
                    ljxdbz = map.get("代表值").toString();

                    ljxjcds += Double.valueOf(map.get("检测点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());

                }

            }
        }
        if (a){
            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("ysdlx","主线");
            map1.put("ysdsczbh",lmMin+"~"+lmMax);
            map1.put("ysdzdbz",zdbz);
            map1.put("ysdydbz",ydbz);
            map1.put("ysdgdz",gdz);
            map1.put("ysdzs",decf.format(jcds));
            map1.put("ysdhgs",decf.format(hgds));
            map1.put("ysdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            resultList.add(map1);
        }
        if (b){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("ysdlx","连接线");
            map2.put("ysdsczbh",ljxMin+"~"+ljxMax);
            map2.put("ysdzdbz",ljxdbz);
            map2.put("ysdydbz",ljxdbz);
            map2.put("ysdgdz",ljxgdz);
            map2.put("ysdzs",decf.format(ljxjcds));
            map2.put("ysdhgs",decf.format(ljxhgds));
            map2.put("ysdhgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
            resultList.add(map2);
        }
        System.out.println(resultList);*/
        return resultList;
    }

    private static List<Map<String, Object>> extractedmethod(CommonInfoVo commonInfoVo, List<Map<String, Object>> bzmdzxList, DecimalFormat df,String zx,String bz) {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map<String, List<Map<String, Object>>> result = bzmdzxList.stream()
                .collect(Collectors.groupingBy(map -> (String) map.get("密度规定值")));
        result.forEach((group, grouphtdData) -> {
            double jcds = 0,hgds = 0;
            String zxzfdbz="",zxyfdbz="";
            List<Double> bhfwz = new ArrayList<>();
            for (Map<String, Object> grouphtdDatum : grouphtdData) {
                jcds += Double.valueOf(grouphtdDatum.get("检测点数").toString());
                hgds += Double.valueOf(grouphtdDatum.get("合格点数").toString());
                String lmlx = grouphtdDatum.get("路面类型").toString();
                if (lmlx.contains("左幅")){
                    zxzfdbz = grouphtdDatum.get("密度代表值").toString();
                }else if(lmlx.contains("右幅")) {
                    zxyfdbz = grouphtdDatum.get("密度代表值").toString();
                }else {
                    zxzfdbz = grouphtdDatum.get("密度代表值").toString();
                    zxyfdbz = grouphtdDatum.get("密度代表值").toString();
                }
                String sczbhfw = grouphtdDatum.get("实测值变化范围").toString();
                bhfwz.add(Double.valueOf(StringUtils.substringBefore(sczbhfw,"~")));
                bhfwz.add(Double.valueOf(StringUtils.substringAfter(sczbhfw,"~")));
            }
            String finbzsczbhfw = "";
            if (!bhfwz.isEmpty()) {
                double min = Collections.min(bhfwz);
                double max = Collections.max(bhfwz);
                finbzsczbhfw = min + "~" + max;
            }
            Map map1 = new HashMap();
            map1.put("htd", commonInfoVo.getHtd());
            map1.put("ysdlx",zx);
            map1.put("bz",bz);
            map1.put("bzsczbhfw",finbzsczbhfw);
            map1.put("zmddbz",zxzfdbz);
            map1.put("ymddbz",zxyfdbz);
            map1.put("bzmdgdz",group);
            map1.put("zxbzjcds",jcds);
            map1.put("zxbzhgds",hgds);
            map1.put("hgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            resultList.add(map1);
        });
        return resultList;
    }


    /**
     * 表4.1.1-6 路基工程检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getljhzData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultlist = new ArrayList<>();
        Map resultmap = new LinkedHashMap();
        List<Map<String, Object>> list1 = gettsfData(commonInfoVo);
        List<Map<String, Object>> list2 = getpsData(commonInfoVo);
        List<Map<String, Object>> list3 = getxqData(commonInfoVo);
        List<Map<String, Object>> list4 = gethdData(commonInfoVo);
        List<Map<String, Object>> list5 = getzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            resultmap.putAll(list1.get(0));
        }
        if (CollectionUtils.isNotEmpty(list2)){
            resultmap.putAll(list2.get(0));
        }
        if (CollectionUtils.isNotEmpty(list3)){
            resultmap.putAll(list3.get(0));
        }
        if (CollectionUtils.isNotEmpty(list4)){
            resultmap.putAll(list4.get(0));
        }
        if (CollectionUtils.isNotEmpty(list5)){
            resultmap.putAll(list5.get(0));
        }
        resultlist.add(resultmap);
        return resultlist;
    }


    /**
     * 表4.1.1-5  支挡工程检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> zdlist = new ArrayList<>();
        Map<String, Object> zdmap = new HashMap<>();
        List<Map<String, Object>> list12 = jjgFbgcLjgcZddmccService.lookJdbjg(commonInfoVo);
        if (list12!=null && list12.size()>0){
            zdmap.put("zddmccjcds",list12.get(0).get("检测总点数"));
            zdmap.put("zddmcchgds",list12.get(0).get("合格点数"));
            zdmap.put("zddmcchgl",list12.get(0).get("合格率"));
        }

        List<Map<String, Object>> list13 = jjgFbgcLjgcZdgqdService.lookJdbjg(commonInfoVo);
        if (list13!=null && list13.size()>0){
            zdmap.put("zdtqdjcds",list13.get(0).get("总点数"));
            zdmap.put("zdtqdhgds",list13.get(0).get("合格点数"));
            zdmap.put("zdtqdhgl",list13.get(0).get("合格率"));
        }
        if (list12!=null && list12.size()>0 || list13!=null && list13.size()>0){
            zdmap.put("htd", commonInfoVo.getHtd());
            zdmap.put("sheetname","表4.1.1-5");
            zdlist.add(zdmap);
        }

        return zdlist;
    }

    /**
     * 表4.1.1-4  涵洞检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> hdlist = new ArrayList<>();
        Map<String, Object> hdmap = new HashMap<>();
        List<Map<String, Object>> list10 = jjgFbgcLjgcHdgqdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> list11 = jjgFbgcLjgcHdjgccService.lookJdbjg(commonInfoVo);
        if (list10!=null && list10.size()>0){
            hdmap.put("hdtqdjcds",list10.get(0).get("总点数"));
            hdmap.put("hdtqdhgds",list10.get(0).get("合格点数"));
            hdmap.put("hdtqdhgl",list10.get(0).get("合格率"));
        }
        if (list11!=null && list11.size()>0){
            hdmap.put("hdjgccjcds",list11.get(0).get("总点数"));
            hdmap.put("hdjgcchgds",list11.get(0).get("合格点数"));
            hdmap.put("hdjgcchgl",list11.get(0).get("合格率"));
        }
        if (list10!=null && list10.size()>0 || list11!=null && list11.size()>0){
            hdmap.put("htd", commonInfoVo.getHtd());
            hdmap.put("sheetname","表4.1.1-4");
            hdlist.add(hdmap);
        }

        return hdlist;
    }

    /**
     * 表4.1.1-2  小桥检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>>  getxqData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> xqlist = new ArrayList<>();
        Map<String, Object> xqmap = new HashMap<>();
        List<Map<String, Object>> list8 = jjgFbgcLjgcXqgqdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list8)){
            xqmap.put("xqtqdjcds",list8.get(0).get("检测总点数"));
            xqmap.put("xqtqdhgds",list8.get(0).get("合格点数"));
            xqmap.put("xqtqdhgl",list8.get(0).get("合格率"));
        }
        List<Map<String, Object>> list9 = jjgFbgcLjgcXqjgccService.lookJdbjg(commonInfoVo);
        if (list9!=null && list9.size()>0){
            xqmap.put("xqjgccjcds",list9.get(0).get("检测总点数"));
            xqmap.put("xqjgcchgds",list9.get(0).get("合格点数"));
            xqmap.put("xqjgcchgl",list9.get(0).get("合格率"));
        }
        if (list8!=null && list8.size()>0 || list9!=null && list9.size()>0){
            xqmap.put("htd", commonInfoVo.getHtd());
            xqmap.put("sheetname","表4.1.1-3");
            xqlist.add(xqmap);
        }
        return xqlist;
    }

    /**
     * 表4.1.1-2  排水工程检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>>  getpsData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> pslist = new ArrayList<>();
        Map<String, Object> psmap = new LinkedHashMap<>();
        List<Map<String, Object>> list6 = jjgFbgcLjgcPsdmccService.lookJdbjg(commonInfoVo);
        if (list6!=null && list6.size()>0){
            psmap.put("psdmccjcds",list6.get(0).get("检测总点数"));
            psmap.put("psdmcchgds",list6.get(0).get("合格点数"));
            psmap.put("psdmcchgl",list6.get(0).get("合格率"));
        }
        List<Map<String, Object>> list7 = jjgFbgcLjgcPspqhdService.lookJdbjg(commonInfoVo);
        if (list7!=null && list7.size()>0){
            psmap.put("pspqhdjcds",list7.get(0).get("检测总点数"));
            psmap.put("pspqhdhgds",list7.get(0).get("合格点数"));
            psmap.put("pspqhdhgl",list7.get(0).get("合格率"));
        }
        if (list6!=null && list6.size()>0 || list7!=null && list7.size()>0){
            psmap.put("htd", commonInfoVo.getHtd());
            psmap.put("sheetname","表4.1.1-2");
            pslist.add(psmap);
        }
        return pslist;
    }

    /**
     * 表4.1.1-1  路基土石方检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gettsfData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.0");
        DecimalFormat decf = new DecimalFormat("0.#");
        List<Map<String,Object>> tsflist = new ArrayList<>();
        Map<String, Object> tsdmap = new HashMap<>();
        List<Map<String, Object>> list1 = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
        if (list1!=null && list1.size()>0){
            double j = 0;
            double h = 0;
            for (Map<String, Object> map : list1) {
                j += Double.valueOf(map.get("检测点数").toString());
                h += Double.valueOf(map.get("合格点数").toString());

            }
            tsdmap.put("ysdjcds", decf.format(j));
            tsdmap.put("ysdhgds", decf.format(h));
            tsdmap.put("ysdhgl",j!=0 ? df.format(h/j*100) : 0);
        }
        //沉降
        List<Map<String, Object>> list2 = jjgFbgcLjgcLjcjService.lookJdbjg(commonInfoVo);
        if (list2!=null && list2.size()>0){
            for (Map<String, Object> map : list2) {
                tsdmap.put("cjjcds",map.get("总点数"));
                tsdmap.put("cjhgds",map.get("合格点数"));
                tsdmap.put("cjhgl",map.get("合格率"));
            }
        }

        //弯沉
        List<Map<String, Object>> list3 = jjgFbgcLjgcLjwcLcfService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> list4 = jjgFbgcLjgcLjwcService.lookJdbjg(commonInfoVo);
        double d = 0 ,dd = 0;
        boolean f = false;
        if (list3!=null&&list3.size()>0){
            f = true;
            d += Double.valueOf(list3.get(0).get("检测单元数").toString());
            dd += Double.valueOf(list3.get(0).get("合格单元数").toString());
        }
        if (list4!=null&&list4.size()>0){
            f = true;
            d += Double.valueOf(list4.get(0).get("检测单元数").toString());
            dd += Double.valueOf(list4.get(0).get("合格单元数").toString());
        }
        /*double d = Double.valueOf(list3.get(0).get("检测单元数").toString()) + Double.valueOf(list4.get(0).get("检测单元数").toString());
        double dd = Double.valueOf(list3.get(0).get("合格单元数").toString()) + Double.valueOf(list4.get(0).get("合格单元数").toString());*/
        if (f){
            tsdmap.put("wcjcds", decf.format(d));
            tsdmap.put("wchgds", decf.format(dd));
            tsdmap.put("wchgl",d!=0 ? df.format(dd/d*100) : 0);
        }
        System.out.println(tsdmap);

        //边坡
        List<Map<String, Object>> list5 = jjgFbgcLjgcLjbpService.lookJdbjg(commonInfoVo);
        if (list5!=null && list5.size()>0){
            tsdmap.put("bpjcds",list5.get(0).get("总点数"));
            tsdmap.put("bphgds",list5.get(0).get("合格点数"));
            tsdmap.put("bphgl",list5.get(0).get("合格率"));
        }

        if (list1!=null && list1.size()>0 || list2!=null && list2.size()>0 || list3!=null&&list3.size()>0 || list4!=null&&list4.size()>0 || list5!=null && list5.size()>0){
            tsdmap.put("htd", commonInfoVo.getHtd());
            tsdmap.put("sheetname","表4.1.1-1");
            tsflist.add(tsdmap);
        }
        return tsflist;
    }

    @Autowired
    private JjgJanumService jjgJanumService;

    /**
     * 表3.4.3-1 交通安全设施抽检统计表
     * @param proname
     * @param htd
     * @return
     */
    private List<Map<String,Object>> getLjjaData(String proname, String htd) {
        DecimalFormat df = new DecimalFormat("0.0");
        List<Map<String,Object>> fhllist = new ArrayList<>();
        Map<String,Object> map = new HashMap<>();

        Map<String,Object> map1 = jjgFbgcJtaqssJabzService.selectchs(proname, htd);
        int bzccs = 0;
        if (map1.size()>0){
            bzccs = Integer.valueOf(map1.get("wz").toString());
        }
        map.put("bzccs",bzccs);
        int bznum = jjgJanumService.selectbznum(proname,htd);
        map.put("bzsys",bznum);
        map.put("bzcjpl",bznum!=0 ? df.format(bzccs/bznum*100) : 0);


        Map<String,Object> map2 = jjgFbgcJtaqssJabxService.selectchs(proname, htd);
        int bxccs = 0;
        if (map2.size()>0){
            bxccs = Integer.valueOf(map2.get("wz").toString());
        }
        map.put("bxccs",bxccs);
        int bxnum = jjgJanumService.selectbxnum(proname,htd);
        map.put("bxsys",bxnum);
        map.put("bxcjpl",bxnum!=0 ? df.format(bxccs/bxnum*100) : 0);

        Map<String,Object> map3 = jjgFbgcJtaqssJathlqdService.selectchs(proname, htd);
        int fhlccs = 0;
        if (map3.size()>0){
            fhlccs = Integer.valueOf(map3.get("zh").toString());
        }
        map.put("fhlccs",fhlccs);
        int fhlnum = jjgJanumService.selectfhlnum(proname,htd);
        map.put("fhlsys",fhlnum);
        map.put("fhlcjpl",fhlnum!=0 ? df.format(fhlccs/fhlnum*100) : 0);

        if (map1.size()>0 || map2.size()>0 || map3.size()>0){
            map.put("htd", htd);
            map.put("sheetname","表3.4.3-1");
            fhllist.add(map);
        }
        return fhllist;
    }

    /**
     * 表3.4.4-1  隧道工程抽检统计表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdcjData(CommonInfoVo commonInfoVo) {
        List<Map<String,Object>> resultlist = new ArrayList<>();
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<JjgLqsSd> list = jjgLqsSdService.getsdName(proname,htd);
        List<Map<String,Object>> sdcdlist = new ArrayList<>();//[{sdmc=qqq, length=特长隧道}, {sdmc=李家湾隧道左线, length=长隧道}, {sdmc=www, length=特长隧道}, {sdmc=李家湾隧道右线, length=长隧道}]
        if (list.size()>0){
            /**
             * 获取到路桥隧文件中当前项目合同段下的所有隧道
             * 然后判断每个隧道的length,结果存入sdcdlist
             */

            List<JjgLqsSd> uniqueList = list.stream()
                    .collect(Collectors.toMap(
                            JjgLqsSd::getSdname, // 假设getSdmc是获取sdmc属性的方法
                            Function.identity(),
                            (existing, replacement) -> existing // 如果key冲突，保留现有的
                    ))
                    .values()
                    .stream()
                    .collect(Collectors.toList());

            for (JjgLqsSd jjgLqsSd : uniqueList) {
                Map map = new HashMap();
                String sdqc = jjgLqsSd.getSdqc();
                String sdname = jjgLqsSd.getSdname();
                Double sdc = Double.valueOf(sdqc);
                if (sdc >= 3000){
                    map.put("length","特长隧道");

                }else if (sdc < 3000 && sdc >= 1000){
                    map.put("length","长隧道");

                }else if (sdc < 1000 && sdc >= 500){
                    map.put("length","中隧道");

                }else if (sdc < 500){
                    map.put("length","短隧道");
                }
                map.put("sdmc",sdname);
                sdcdlist.add(map);
            }
            /**
             * 根据length属性的值分组，然后得到分组后每个map的value长度，也就是某个length的数量
             */
            Map<String, List<Map<String, Object>>> groupedData =
                    sdcdlist.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("length").toString()
                            ));
            Map<String, Object> keyLengths = new HashMap<>();//{长隧道=2, 特长隧道=2}  实际的 路桥遂文件中的
            groupedData.forEach((key, value) -> keyLengths.put(key, value.size()));

            /**
             * 获取实测数据中的隧道
             */
            List<Map<String,Object>> sdnumList = jjgFbgcSdgcCqtqdService.getsdnum(proname,htd);//[{sdmc=李家湾隧道左线}, {sdmc=李家湾隧道右线}]
            List<Map<String,Object>> sdnumList1 = jjgFbgcSdgcCqhdService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList2 = jjgFbgcSdgcDmpzdService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList3 = jjgFbgcSdgcGssdlqlmhdzxfService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList4 = jjgFbgcSdgcHntlmqdService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList5 = jjgFbgcSdgcLmgzsdsgpsfService.getsdnum(proname,htd);
            //jjgFbgcSdgc
            List<Map<String,Object>> sdnumList6 = jjgFbgcSdgcLmssxsService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList7 = jjgFbgcSdgcSdhntlmhdzxfService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList8 = jjgFbgcSdgcSdhpService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList9 = jjgFbgcSdgcSdlqlmysdService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList10 = jjgFbgcSdgcTlmxlbgcService.getsdnum(proname,htd);
            List<Map<String,Object>> sdnumList11 = jjgFbgcSdgcZtkdService.getsdnum(proname,htd);
            sdnumList.addAll(sdnumList1);
            sdnumList.addAll(sdnumList2);
            sdnumList.addAll(sdnumList3);
            sdnumList.addAll(sdnumList4);
            sdnumList.addAll(sdnumList5);
            sdnumList.addAll(sdnumList6);
            sdnumList.addAll(sdnumList7);
            sdnumList.addAll(sdnumList8);
            sdnumList.addAll(sdnumList9);
            sdnumList.addAll(sdnumList10);
            sdnumList.addAll(sdnumList11);
            List<Map<String, Object>> distinctList = sdnumList.stream()
                    .collect(Collectors.toMap(
                            // 以qlmc作为key
                            m -> m.get("sdmc"),
                            // 如果有重复的key，则选择第一个Map
                            m -> m,
                            // 合并函数，当有重复的key时选择哪个Map
                            (existing, replacement) -> existing))
                    .values()
                    .stream()
                    .collect(Collectors.toList());

            Map resultMap = new HashMap<>();
            resultMap.put("htd",htd);
            resultMap.put("sheetname","表3.4.4-1");

            if (distinctList.size()>0){
                /**
                 * 先判断实测数据中每个隧道的length,将结果存入到cjnum
                 */
                Set<Map<String,Object>> cjnum = new HashSet<>();

                for (Map<String, Object> cjmap : distinctList) {//实测数据的
                    String cjsdmc = cjmap.get("sdmc").toString();
                    for (Map<String, Object> symap : sdcdlist) {//路桥隧的
                        String sysdmc = symap.get("sdmc").toString();
                        if (cjsdmc.equals(sysdmc)){
                            Map maps = new HashMap();
                            maps.put("sdmc",sysdmc);
                            maps.put("length",symap.get("length"));
                            cjnum.add(maps);
                        }
                    }
                }

                if (cjnum.size()>0){
                    Map<String, List<Map<String, Object>>> cjgroupedData =
                            cjnum.stream()
                                    .collect(Collectors.groupingBy(
                                            item -> item.get("length").toString()
                                    ));
                    Map<String, Object> cjkeyLengths = new HashMap<>();//{长隧道=2} 抽检的 实测文件中的
                    /**
                     * 然后就得到了实测数据中隧道的length情况，cjgroupedData
                     */
                    cjgroupedData.forEach((key, value) -> cjkeyLengths.put(key, value.size()));

                    for (String key : cjkeyLengths.keySet()) {
                        if (key.equals("特长隧道")){
                            resultMap.put("tccjs",cjkeyLengths.get(key));
                        }else if (key.equals("长隧道")){
                            resultMap.put("ccjs",cjkeyLengths.get(key));
                        }else if (key.equals("中隧道")){
                            resultMap.put("zcjs",cjkeyLengths.get(key));
                        }else if (key.equals("短隧道")){
                            resultMap.put("dcjs",cjkeyLengths.get(key));
                        }

                    }
                }
            }
            //实有数
            if (keyLengths.size()>0){
                for (String key : keyLengths.keySet()) {
                    if (key.equals("特长隧道")){
                        resultMap.put("tcccs",keyLengths.get(key));
                    } else if (key.equals("长隧道")){
                        resultMap.put("cccs",keyLengths.get(key));
                    } else if (key.equals("中隧道")){
                        resultMap.put("zccs",keyLengths.get(key));
                    } else if (key.equals("短隧道")){
                        resultMap.put("dccs",keyLengths.get(key));
                    }

                }
            }
            resultlist.add(resultMap);
        }
        return resultlist;
    }

    /**
     * 表3.4.2-1  桥梁工程抽检统计表
     * @param proname
     * @param htd
     * @return
     */
    private List<Map<String,Object>>  getLjqlcjData(String proname, String htd) {
        List<Map<String,Object>> qllist = new ArrayList<>();
        List<JjgLqsQl> list = jjgLqsQlService.getqlName(proname, htd);
        List<Map<String,Object>> qlxh = new ArrayList<>();
        if (list.size()>0){
            List<JjgLqsQl> distinctqlnameList = list.stream()
                    .collect(Collectors.toMap(
                            // 以qlmc作为key
                            m -> m.getQlname(),
                            // 如果有重复的key，则选择第一个Map
                            m -> m,
                            // 合并函数，当有重复的key时选择哪个Map
                            (existing, replacement) -> existing))
                    .values()
                    .stream()
                    .collect(Collectors.toList());
            for (JjgLqsQl jjgLqsQl : distinctqlnameList) {
                Map<String,Object> map = new HashMap<>();
                String qlname = jjgLqsQl.getQlname();
                String qlqc = jjgLqsQl.getQlqc();

                Double qc = Double.valueOf(qlqc);

                if (qc >= 8 && qc <= 30){
                    map.put("length","小桥");
                }else if (qc > 30 && qc < 100){
                    map.put("length","中桥");
                }else if (qc >= 100 && qc <= 1000){
                    map.put("length","大桥");
                }else if (qc > 1000){
                    map.put("length","特大桥");
                }
                map.put("qlname",qlname);
                qlxh.add(map);
            }
            System.out.println(qlxh);
            Map<String, List<Map<String, Object>>> groupedData =
                    qlxh.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("length").toString()
                            ));
            Map<String, Object> keyLengths = new HashMap<>();//路桥遂文件中的
            groupedData.forEach((key, value) -> keyLengths.put(key, value.size()));
            System.out.println(keyLengths);//{中桥=5, 大桥=4, 小桥=2, 特大桥=2}
            //获取全部实测数据得桥
            List<Map<String,Object>> qlnumList = jjgFbgcQlgcSbTqdService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList2 = jjgFbgcQlgcSbJgccService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList3 = jjgFbgcQlgcSbBhchdService.getqlname(proname,htd);

            List<Map<String,Object>> qlnumList4 = jjgFbgcQlgcXbTqdService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList5 = jjgFbgcQlgcXbSzdService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList6 = jjgFbgcQlgcXbJgccService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList7 = jjgFbgcQlgcXbBhchdService.getqlname(proname,htd);

            List<Map<String,Object>> qlnumList8 = jjgFbgcQlgcQmpzdService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList9 = jjgFbgcQlgcQmhpService.getqlname(proname,htd);
            List<Map<String,Object>> qlnumList10 = jjgFbgcQlgcQmgzsdService.getqlname(proname,htd);
            qlnumList.addAll(qlnumList2);
            qlnumList.addAll(qlnumList3);
            qlnumList.addAll(qlnumList4);
            qlnumList.addAll(qlnumList5);
            qlnumList.addAll(qlnumList6);
            qlnumList.addAll(qlnumList7);
            qlnumList.addAll(qlnumList8);
            qlnumList.addAll(qlnumList9);
            qlnumList.addAll(qlnumList10);

            //实测数据中的桥 也就是抽查数
            List<Map<String, Object>> distinctList = qlnumList.stream()
                    .collect(Collectors.toMap(
                            // 以qlmc作为key
                            m -> m.get("qlmc"),
                            // 如果有重复的key，则选择第一个Map
                            m -> m,
                            // 合并函数，当有重复的key时选择哪个Map
                            (existing, replacement) -> existing))
                    .values()
                    .stream()
                    .collect(Collectors.toList());

            List<Map<String,Object>> cjnum = new ArrayList<>();
            Map resultMap = new HashMap<>();
            resultMap.put("htd",htd);
            resultMap.put("sheetname","表3.4.2-1");
            if (distinctList.size()>0){
                for (Map<String, Object> cjmap : distinctList) {//实测数据的
                    String cjqlmc = cjmap.get("qlmc").toString();
                    for (Map<String, Object> symap : qlxh) {//路桥隧的
                        String syqlmc = symap.get("qlname").toString();
                        if (cjqlmc.equals(syqlmc)){
                            Map map = new HashMap();
                            map.put("qlmc",syqlmc);
                            map.put("length",symap.get("length"));
                            cjnum.add(map);
                        }
                    }
                }
                System.out.println(cjnum);
                //[{qlmc=西城大道中桥, length=中桥}, {qlmc=航空城互通E匝道桥, length=中桥}, {qlmc=群英渠中桥, length=中桥},
                // {qlmc=航空城互通A匝道桥, length=大桥}, {qlmc=富平立交DK0+200匝道桥, length=中桥}, {qlmc=罗家村中桥, length=特大桥},
                // {qlmc=东城大道中桥, length=大桥}, {qlmc=富平立交AK0+807匝大道桥, length=中桥}, {qlmc=李家村中桥, length=大桥}]
                if (cjnum.size()>0){
                    //分组
                    Map<String, List<Map<String, Object>>> cjgroupedData =
                            cjnum.stream()
                                    .collect(Collectors.groupingBy(
                                            item -> item.get("length").toString()
                                    ));
                    Map<String, Object> cjkeyLengths = new HashMap<>();
                    cjgroupedData.forEach((key, value) -> cjkeyLengths.put(key, value.size()));
                    System.out.println(cjkeyLengths);//{中桥=5, 大桥=3, 特大桥=1}
                    //抽查数
                    for (String key : cjkeyLengths.keySet()) {
                        if (key.equals("特大桥")){
                            resultMap.put("tdcjs",cjkeyLengths.get(key));
                        } else if (key.equals("大桥")){
                            resultMap.put("dcjs",cjkeyLengths.get(key));
                        } else if (key.equals("中桥")){
                            resultMap.put("zcjs",cjkeyLengths.get(key));
                        } else if (key.equals("小桥")){
                            resultMap.put("xcjs",cjkeyLengths.get(key));
                        }
                    }
                }
            }
            //实有数
            if (keyLengths.size()>0){
                for (String key : keyLengths.keySet()) {
                    if (key.equals("特大桥")){
                        resultMap.put("tdccs",keyLengths.get(key));
                    }else if (key.equals("大桥")){
                        resultMap.put("dccs",keyLengths.get(key));
                    } else if (key.equals("中桥")){
                        resultMap.put("zccs",keyLengths.get(key));
                    } else if (key.equals("小桥")){
                        resultMap.put("xccs",keyLengths.get(key));
                    }
                }
            }
            System.out.println(resultMap);
            qllist.add(resultMap);
        }
        return qllist;
    }

    @Autowired
    private JjgLjfrnumService jjgLjfrnumService;

    /**
     * 表3.4.1-1 涵洞、支挡工程抽检统计表
     * 实有数是用户自己输入，但是这块还要给抽检频率插入公式（12.23：代码写的是识别的，不是自己输入的）
     * @param proname
     * @param htd
     * @return
     */
    private List<Map<String,Object>> getLjcjtjbData(String proname, String htd) {
        DecimalFormat df = new DecimalFormat("0.0");
        List<Map<String,Object>> ljlist = new ArrayList<>();
        Map<String,Object> map = new LinkedHashMap<>();
        Map<String,Object> map1 = jjgFbgcLjgcHdgqdService.selectchs(proname, htd);
        // 涵洞
        int hdccs = 0;
        if (!"".equals(map1.get("ccs"))){
            hdccs = Integer.valueOf(map1.get("ccs").toString());
        }
        map.put("hdccs",hdccs);
        int hdnum = jjgLjfrnumService.selecthdnum(proname,htd);
        map.put("hdsys",hdnum);
        map.put("hdcjpl",hdnum!=0 ? df.format((double) hdccs/hdnum*100) : 0);

        //支挡工程
        Map<String,Object> map2 = jjgFbgcLjgcZddmccService.selectchs(proname, htd);
        int zdccs = 0;
        if (!"".equals(map2.get("ccs"))){
            zdccs = Integer.valueOf(map2.get("ccs").toString());
        }
        map.put("zdccs",zdccs);
        int zdnum = jjgLjfrnumService.selectzdnum(proname,htd);
        map.put("zdsys",zdnum);
        map.put("zdcjpl",zdnum!=0 ? df.format((double)zdccs/zdnum*100) : 0);

        //小桥
        Map<String,Object> map3 = jjgFbgcLjgcXqgqdService.selectchs(proname, htd);
        int xqccs = 0;
        if (!"".equals(map3.get("ccs"))){
            xqccs = Integer.valueOf(map3.get("ccs").toString());
        }
        map.put("xqccs",xqccs);
        int xqnum = jjgLjfrnumService.selectxqnum(proname,htd);
        map.put("xqsys",xqnum);
        map.put("xqcjpl",xqnum!=0 ? df.format((double)xqccs/xqnum*100) : 0);

        if (!map1.get("ccs").toString().equals("0") || !map2.get("ccs").toString().equals("0") || !map3.get("ccs").toString().equals("0")){
            map.put("htd", htd);
            map.put("sheetname","表3.4.1-1");
            ljlist.add(map);
        }
        return ljlist;
    }

    // 根据路面类型返回文件名
    private String getFileName(String lx) {
        if (lx.contains("隧道")) {
            return "沥青隧道路面厚度质量鉴定表（钻芯法）";
        } else if (lx.contains("路面") || lx.contains("路面")) {
            return "沥青路面厚度质量鉴定表（钻芯法）";
        } else if (lx.contains("连接线")) {
            return "沥青连接线路面厚度质量鉴定表（钻芯法）";
        }
        return "未知类型";
    }

    // 根据路面类型返回ccname2
    private String getCcname2(String lx) {
        if (lx.contains("隧道")) {
            return "隧道路面";
        } else if (lx.contains("路面") || lx.contains("路面")) {
            return "路面面层";
        } else if (lx.contains("连接线")) {
            return "连接线路面";
        }
        return "未知类型";
    }

    private static List<Map<String, Object>> processJgMap(List<Map<String, Object>> jgmap) {
        // 主线或者互通的路面相加， 其他的桥、隧左右幅相加
        DecimalFormat df = new DecimalFormat("#.00");
        List<Map<String, Object>> result = new ArrayList<>();

        // 处理条件1：主线路面类型
        List<Map<String, Object>> condition1Maps = jgmap.stream()
                .filter(map -> {
                    String lmType = (String) map.get("路面类型");
                    String jcxm = (String) map.get("检测项目");
                    return lmType != null && lmType.contains("路面") &&
                            jcxm != null && ("主线".equals(jcxm) || jcxm.contains("互通"));
                })
                .collect(Collectors.toList());

        if (!condition1Maps.isEmpty()) {
            int totalZds = condition1Maps.stream().mapToInt(m -> Integer.valueOf(m.get("总点数").toString())).sum();
            int totalHgds = condition1Maps.stream().mapToInt(m -> Integer.valueOf(m.get("合格点数").toString())).sum();
            double hgl = totalZds != 0 ? (totalHgds * 100.0 / totalZds) : 0;

            // 获取主线检测项目的设计值（优先取主线项目的设计值）
            String designValue = condition1Maps.stream()
                    .filter(m -> "主线".equals(m.get("检测项目")))
                    .map(m -> (String) m.get("设计值"))
                    .filter(Objects::nonNull)
                    .findFirst()
                    .orElseGet(() ->
                            condition1Maps.get(0).getOrDefault("设计值", "").toString()
                    );

            Map<String, Object> merged = new HashMap<>();
            merged.put("分部工程名称", "主线路面");
            merged.put("检测项目", "主线路面");
            merged.put("总点数", totalZds);
            merged.put("合格点数", totalHgds);
            merged.put("合格率", df.format(hgl));
            merged.put("设计值", designValue);  // 新增设计值保留
            merged.put("路面类型", "主线路面");  // 明确路面类型
            result.add(merged);
        }

        // 处理条件2：其他相同分部工程和检测项目的
        Map<String, List<Map<String, Object>>> groups = jgmap.stream()
                .filter(map -> !condition1Maps.contains(map)) // 排除条件1的数据
                .collect(Collectors.groupingBy(map ->
                        ((String) map.get("分部工程名称")) + "|" + ((String) map.get("检测项目")))
                );

        for (List<Map<String, Object>> group : groups.values()) {
            Map<String, Object> first = group.get(0);
            int totalZds = group.stream().mapToInt(m -> Integer.valueOf(m.get("总点数").toString())).sum();
            int totalHgds = group.stream().mapToInt(m -> Integer.valueOf(m.get("合格点数").toString())).sum();
            double hgl = totalZds != 0 ? (totalHgds * 100.0 / totalZds) : 0;

            Map<String, Object> merged = new HashMap<>(first);
            merged.put("总点数", totalZds);
            merged.put("合格点数", totalHgds);
            merged.put("合格率", df.format(hgl));
            result.add(merged);
        }

        return result;
    }

    // 这个函数没有左右幅相加
    private static List<Map<String, Object>> processPZDList(List<Map<String, Object>> jgmap) {
        DecimalFormat df = new DecimalFormat("#.00");
        List<Map<String, Object>> result = new ArrayList<>();

        // 处理条件1：沥青路面/匝道合并
        List<Map<String, Object>> asphaltMerge = jgmap.stream()
                .filter(map -> {
                    String roadType = (String) map.get("路面类型");
                    String testItem = (String) map.get("检测项目");
                    return (roadType != null && (roadType.equals("沥青路面") || roadType.equals("沥青匝道")))
                            && (testItem != null && !testItem.contains("连接线"));
                })
                .collect(Collectors.toList());

        // 处理沥青类合并
        if (!asphaltMerge.isEmpty()) {
            Map<String, Object> merged = mergeMaps(asphaltMerge, "沥青路面", "主线路面", df);
            result.add(merged);
        }

        // 处理条件2：混凝土收费站合并
        List<Map<String, Object>> concreteMerge = jgmap.stream()
                .filter(map -> "混凝土收费站".equals(map.get("路面类型")))
                .collect(Collectors.toList());

        // 处理混凝土收费站合并
        if (!concreteMerge.isEmpty()) {
            Map<String, Object> merged = mergeMaps(concreteMerge, null, "混凝土收费站", df);
            result.add(merged);
        }

        // 添加未处理条目（排除两种合并类型）
        jgmap.stream()
                .filter(map -> !asphaltMerge.contains(map) && !concreteMerge.contains(map))
                .forEach(result::add);

        return result;
    }

    // 通用合并方法
    private static Map<String, Object> mergeMaps(List<Map<String, Object>> maps,
                                                 String priorityType,
                                                 String newDeptName,
                                                 DecimalFormat df) {
        // 数值累加
        int totalZds = maps.stream()
                .mapToInt(m -> Integer.valueOf(m.get("总点数").toString()))
                .sum();
        int totalHgds = maps.stream()
                .mapToInt(m -> Integer.valueOf(m.get("合格点数").toString()))
                .sum();

        // 获取基准数据（优先选择指定类型）
        Map<String, Object> baseMap = maps.stream()
                .filter(m -> priorityType != null && priorityType.equals(m.get("路面类型")))
                .findFirst()
                .orElseGet(() -> maps.get(0));

        // 构建合并结果
        Map<String, Object> merged = new HashMap<>();
        merged.put("路面类型", baseMap.get("路面类型"));
        merged.put("检测项目", baseMap.get("检测项目"));
        merged.put("分部工程名称", newDeptName);
        merged.put("总点数", totalZds);
        merged.put("合格点数", totalHgds);
        merged.put("合格率", totalZds != 0 ? df.format(totalHgds * 100.0 / totalZds) : "0.00");
        merged.put("设计值", baseMap.get("设计值"));
        merged.put("Max", baseMap.get("Max"));
        merged.put("Min", baseMap.get("Min"));

        return merged;
    }

    // 渗水系数的mapList处理： 路面和匝道合并（有匝道表，而不是互通，相当于直接把互通的路面分出来了）； 其他相同分部工程的左右幅相加
    public List<Map<String, Object>> processSsxsMapList(List<Map<String, Object>> mapList) {
        // 按分部工程名称分组
        Map<String, List<Map<String, Object>>> grouped = mapList.stream()
                .collect(Collectors.groupingBy(map -> map.get("分部工程名称").toString()));

        List<Map<String, Object>> result = new ArrayList<>();

        for (Map.Entry<String, List<Map<String, Object>>> entry : grouped.entrySet()) {
            String fbgc = entry.getKey();
            List<Map<String, Object>> group = entry.getValue();

            Map<String, Object> mergedMap = new HashMap<>();

            // 合并逻辑
            if ("路面".equals(fbgc)) {
                // 特殊处理"路面"分部工程，检测项目统一为"路面"
                mergedMap.put("分部工程名称", fbgc);
                mergedMap.put("检测项目", "路面");
                mergedMap.put("规定值", group.get(0).get("规定值"));

                int totalJcds = 0;
                int totalHgds = 0;
                double maxValue = Double.MIN_VALUE;
                double minValue = Double.MAX_VALUE;

                for (Map<String, Object> item : group) {
                    totalJcds += Integer.parseInt(item.get("检测点数").toString());
                    totalHgds += Integer.parseInt(item.get("合格点数").toString());

                    double currentMax = Double.parseDouble(item.get("最大值").toString());
                    double currentMin = Double.parseDouble(item.get("最小值").toString());

                    if (currentMax > maxValue) maxValue = currentMax;
                    if (currentMin < minValue) minValue = currentMin;
                }

                mergedMap.put("检测点数", String.valueOf(totalJcds));
                mergedMap.put("合格点数", String.valueOf(totalHgds));
                mergedMap.put("合格率",
                        totalJcds == 0 ? "0.00" : String.format("%.2f", (totalHgds * 100.0 / totalJcds)));
                mergedMap.put("最大值", String.valueOf(maxValue));
                mergedMap.put("最小值", String.valueOf(minValue));
            } else {
                // 其他分部工程按原检测项目名称保留
                mergedMap.put("分部工程名称", fbgc);
                mergedMap.put("检测项目", group.get(0).get("检测项目")); // 保留第一个检测项目名称
                mergedMap.put("规定值", group.get(0).get("规定值"));

                int totalJcds = 0;
                int totalHgds = 0;
                double maxValue = Double.MIN_VALUE;
                double minValue = Double.MAX_VALUE;

                for (Map<String, Object> item : group) {
                    totalJcds += Integer.parseInt(item.get("检测点数").toString());
                    totalHgds += Integer.parseInt(item.get("合格点数").toString());

                    double currentMax = Double.parseDouble(item.get("最大值").toString());
                    double currentMin = Double.parseDouble(item.get("最小值").toString());

                    if (currentMax > maxValue) maxValue = currentMax;
                    if (currentMin < minValue) minValue = currentMin;
                }

                mergedMap.put("检测点数", String.valueOf(totalJcds));
                mergedMap.put("合格点数", String.valueOf(totalHgds));
                mergedMap.put("合格率",
                        totalJcds == 0 ? "0.00" : String.format("%.2f", (totalHgds * 100.0 / totalJcds)));
                mergedMap.put("最大值", String.valueOf(maxValue));
                mergedMap.put("最小值", String.valueOf(minValue));
            }

            result.add(mergedMap);
        }

        return result;
    }


}
