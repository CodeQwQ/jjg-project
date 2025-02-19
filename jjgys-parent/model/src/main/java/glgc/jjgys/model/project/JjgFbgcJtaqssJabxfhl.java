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
 * @since 2023-03-01
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_jtaqss_jabxfhl")
public class JjgFbgcJtaqssJabxfhl implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @TableField("jcsj")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @TableField("wzjlx")
    private String wzjlx;

    @TableField("jdhdgdz")
    private String jdhdgdz;

    @TableField("jdhdscz1")
    private String jdhdscz1;

    @TableField("jdhdscz2")
    private String jdhdscz2;

    @TableField("jdhdscz3")
    private String jdhdscz3;

    @TableField("jdhdscz4")
    private String jdhdscz4;

    @TableField("jdhdscz5")
    private String jdhdscz5;

    @TableField("lzbhgdz")
    private String lzbhgdz;

    @TableField("lzbhscz1")
    private String lzbhscz1;

    @TableField("lzbhscz2")
    private String lzbhscz2;

    @TableField("lzbhscz3")
    private String lzbhscz3;

    @TableField("lzbhscz4")
    private String lzbhscz4;

    @TableField("lzbhscz5")
    private String lzbhscz5;

    @TableField("zxgdgdz")
    private String zxgdgdz;

    @TableField("zxgdyxpsz")
    private String zxgdyxpsz;

    @TableField("zxgdyxpsf")
    private String zxgdyxpsf;

    @TableField("zxgdscz1")
    private String zxgdscz1;

    @TableField("zxgdscz2")
    private String zxgdscz2;

    @TableField("zxgdscz3")
    private String zxgdscz3;

    @TableField("zxgdscz4")
    private String zxgdscz4;

    @TableField("zxgdscz5")
    private String zxgdscz5;

    @TableField("mrsdgdz")
    private String mrsdgdz;

    @TableField("mrsdscz")
    private String mrsdscz;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date createtime;

    @TableField("userid")
    private String userid;


}
