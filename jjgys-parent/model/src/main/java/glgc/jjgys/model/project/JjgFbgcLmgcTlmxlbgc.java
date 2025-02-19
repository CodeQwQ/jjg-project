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
 * @since 2023-04-21
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_lmgc_tlmxlbgc")
public class JjgFbgcLmgcTlmxlbgc implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 检测时间
     */
    @TableField("jcsj")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 取样位置名称
     */
    @TableField("qywzmc")
    private String qywzmc;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 实测值1(mm)
     */
    @TableField("scz1")
    private String scz1;

    /**
     * 实测值2(mm)
     */
    @TableField("scz2")
    private String scz2;

    /**
     * 实测值3(mm)
     */
    @TableField("scz3")
    private String scz3;

    /**
     * 实测值4(mm)
     */
    @TableField("scz4")
    private String scz4;

    /**
     * 板高差规定值
     */
    @TableField("bgcgdz")
    private String bgcgdz;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("userid")
    private String userid;


    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("createTime")
    private Date createtime;


}
