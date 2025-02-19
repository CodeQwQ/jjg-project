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
public class JjgYqinfoVo extends BaseRowModel {

    /**
     * 仪器名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "仪器名称" ,index = 0)
    private String name;

    /**
     * 数量
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "数量" ,index = 1)
    private String num;

    /**
     * 管理编号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "管理编号" ,index = 2)
    private String glbh;

    /**
     * 检测项目
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测项目" ,index = 3)
    private String jcxm;

    /**
     * 试验方法
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "试验方法" ,index = 4)
    private String syff;

    /**
     * 仪器状态
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "仪器状态" ,index = 5)
    private String yqzt;



}
