package glgc.jjgys.model.projectvo.jagc;

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
 * @since 2023-03-01
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcJtaqssJabxfhlJgfcVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "合同段名称" ,index = 0)
    private String htd;

    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 1)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @ColumnWidth(23)
    @ExcelProperty(value = "位置及类型" ,index = 2)
    private String wzjlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度规定值(mm)" ,index = 3)
    private String jdhdgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值1(mm)" ,index = 4)
    private String jdhdscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值2(mm)" ,index = 5)
    private String jdhdscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值3(mm)" ,index = 6)
    private String jdhdscz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值4(mm)" ,index = 7)
    private String jdhdscz4;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值5(mm)" ,index = 8)
    private String jdhdscz5;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚规定值(mm)" ,index = 9)
    private String lzbhgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值1(mm)" ,index = 10)
    private String lzbhscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值2(mm)" ,index = 11)
    private String lzbhscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值3(mm)" ,index = 12)
    private String lzbhscz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值4(mm)" ,index = 13)
    private String lzbhscz4;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值5(mm)" ,index = 14)
    private String lzbhscz5;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度规定值(mm)" ,index = 15)
    private String zxgdgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度允许偏差+(mm)" ,index = 16)
    private String zxgdyxpsz;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度允许偏差-(mm)" ,index = 17)
    private String zxgdyxpsf;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值1(mm)" ,index = 18)
    private String zxgdscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值2(mm)" ,index = 19)
    private String zxgdscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值3(mm)" ,index = 20)
    private String zxgdscz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值4(mm)" ,index = 21)
    private String zxgdscz4;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值5(mm)" ,index = 22)
    private String zxgdscz5;

    @ColumnWidth(23)
    @ExcelProperty(value = "埋入深度规定值(mm)" ,index = 23)
    private String mrsdgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "埋入深度实测值(mm)" ,index = 24)
    private String mrsdscz;



}
