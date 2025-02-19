package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;
import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2024-05-02
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_ryinfo_jgfc")
public class JjgRyinfoJgfc implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 姓名
     */
    @TableField("name")
    private String name;

    /**
     * 本项目承担职务
     */
    @TableField("zw")
    private String zw;

    /**
     * 技术职称
     */
    @TableField("jszc")
    private String jszc;

    /**
     * 执业资格证书编号
     */
    @TableField("zgzsbh")
    private String zgzsbh;

    /**
     * 备注
     */
    @TableField("bz")
    private String bz;

    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;


}
