package glgc.jjgys.model.projectvo.sdgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-05-04
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcSdgcTlmxlbgcVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 隧道名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "隧道名称" ,index = 1)
    private String sdmc;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    /**
     * 取样缝位置
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "规定值" ,index = 3)
    private String qyfwz;

    /**
     * 实测值1(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实测值1(mm)" ,index = 4)
    private String scz1;

    /**
     * 实测值2(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实测值2(mm)" ,index = 5)
    private String scz2;

    /**
     * 实测值3
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实测值3(mm)" ,index = 6)
    private String scz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值4(mm)" ,index = 7)
    private String scz4;



}
