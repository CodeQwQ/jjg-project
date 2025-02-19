package glgc.jjgys.model.projectvo.lmgc;

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
 * @since 2023-04-18
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLmgcLmwcLcfJgfcVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "合同段名称" ,index = 0)
    private String htd;

    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 1)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 工程部位
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "工程部位" ,index = 2)
    private String gcbw;

    /**
     * 验收弯沉值(0.01mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "验收弯沉值(0.01mm)" ,index = 3)
    private String yswcz;

    /**
     * 季节影响系数
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "季节影响系数" ,index = 4)
    private String jjyxxs;

    /**
     * 结构层次
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "结构层次" ,index = 5)
    private String jgcc;

    /**
     * 结构类型
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "结构类型" ,index = 6)
    private String jglx;

    /**
     * 目标可靠指标
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "目标可靠指标" ,index = 7)
    private String mbkkzb;

    /**
     * 湿度影响系数
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "湿度影响系数" ,index = 8)
    private String sdyxxs;

    /**
     * 材料层厚度Ha
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "材料层厚度Ha" ,index = 9)
    private String clchd;

    /**
     * 路基顶面回弹模量E0
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路基顶面回弹模量E0" ,index = 10)
    private String ljdmhtml;

    /**
     * 落锤重(T)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "落锤重(T)" ,index = 11)
    private String lcz;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 12)
    private String zh;

    /**
     * 车道
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "车道" ,index = 13)
    private String cd;

    /**
     * 实测弯沉值(0.01㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实测弯沉值(0.01㎜)" ,index = 14)
    private String scwcz;

    /**
     * 路表温度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路表温度" ,index = 15)
    private String lbwd;

    /**
     * 路面类型标注
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路面类型标注" ,index = 16)
    private String lmbzlx;

    /**
     * 沥青层总厚度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "沥青层总厚度" ,index = 17)
    private String lqczhd;

    /**
     * 沥青表面温度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "沥青表面温度" ,index = 18)
    private String lqbmwd;

    /**
     * 测前5h平均温度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测前5h平均温度" ,index = 19)
    private String cqpjwd;

    /**
     * 序号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "序号" ,index = 20)
    private String xh;

    /**
     * 备注
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "备注" ,index = 21)
    private String bz;



}
