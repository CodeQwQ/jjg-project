package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableName;
import java.time.LocalDate;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableField;
import java.io.Serializable;
import java.util.Date;

import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-12-05
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_yqinfo")
public class JjgYqinfo implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 仪器名称
     */
    @TableField("name")
    private String name;

    /**
     * 数量
     */
    @TableField("num")
    private String num;

    /**
     * 管理编号
     */
    @TableField("glbh")
    private String glbh;

    /**
     * 检测项目
     */
    @TableField("jcxm")
    private String jcxm;

    /**
     * 试验方法
     */
    @TableField("syff")
    private String syff;

    /**
     * 仪器状态
     */
    @TableField("yqzt")
    private String yqzt;

    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;


}
