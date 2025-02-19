package glgc.jjgys.model.projectvo.lqs;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-12-05
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgRyinfoJgfcVo extends BaseRowModel {


    /**
     * 姓名
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "姓名" ,index = 0)
    private String name;

    /**
     * 本项目承担职务
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "本项目承担职务" ,index = 1)
    private String zw;

    /**
     * 技术职称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "技术职称" ,index = 2)
    private String jszc;

    /**
     * 执业资格证书编号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "执业资格证书编号" ,index = 3)
    private String zgzsbh;

    /**
     * 备注
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "备注" ,index = 4)
    private String bz;



}
