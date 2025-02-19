package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import com.fasterxml.jackson.annotation.JsonFormat;
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
 * @since 2023-10-23
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_sdgc_zdhgzsd")
public class JjgFbgcSdgcZdhgzsd implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 起点桩号
     */
    @TableField("qdzh")
    private Double qdzh;

    /**
     * 终点桩号
     */
    @TableField("zdzh")
    private Double zdzh;

    /**
     * SFC
     */
    @TableField("mtd")
    private String mtd;

    /**
     * 车道
     */
    @TableField("cd")
    private String cd;

    /**
     * 类型标识
     */
    @TableField("lxbs")
    private String lxbs;

    /**
     * 匝道标识
     */
    @TableField("zdbs")
    private String zdbs;

    @TableField("val")
    private Integer val;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("userid")
    private String userid;

    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("createTime")
    private Date createtime;


}
