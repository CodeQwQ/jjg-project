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
 * @since 2023-03-26
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_sdgc_ztkd")
public class JjgFbgcSdgcZtkd implements Serializable {

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
     * 桥梁名称
     */
    @TableField("sdmc")
    private String sdmc;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 位置
     */
    @TableField("zbk")
    private String zbk;

    /**
     * 前视读数(mm)
     */
    @TableField("ybk")
    private String ybk;

    /**
     * 后视读数(mm)
     */
    @TableField("sjz")
    private String sjz;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("userid")
    private String userid;


    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("createTime")
    private Date createtime;


}
