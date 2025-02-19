package glgc.jjgys.system.controller;


import glgc.jjgys.common.result.Result;
import glgc.jjgys.system.service.JjgFbgcGenerateWordService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;



/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/generate/word")
public class JjgFbgcGenerateWordController {

    @Autowired
    private JjgFbgcGenerateWordService jjgFbgcGenerateWordService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @ApiOperation("生成word")
    @PostMapping("generateword")
    public Result generateword(String proname) throws IOException, InvalidFormatException {
        boolean is_Success = jjgFbgcGenerateWordService.generateword(proname);
        return is_Success ? Result.ok().code(200).message("成功生成") : Result.fail().code(201).message("生成失败");
    }

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void download(HttpServletResponse response, @RequestParam String proname) throws IOException {
        String fileName = "高速公路交工验收质量检测报告.docx";
        String p = filespath+ File.separator+proname+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }
}

