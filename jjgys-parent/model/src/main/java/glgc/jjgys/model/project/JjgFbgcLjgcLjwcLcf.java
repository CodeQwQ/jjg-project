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
 * @since 2023-03-08
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_ljgc_ljwc_lcf")
public class JjgFbgcLjgcLjwcLcf implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @TableField("jcsj")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @TableField("zh")
    private String zh;

    @TableField("sjwcz")
    private String sjwcz;

    @TableField("jgcc")
    private String jgcc;

    @TableField("wdyxxs")
    private String wdyxxs;

    @TableField("jjyxxs")
    private String jjyxxs;

    @TableField("jglx")
    private String jglx;

    @TableField("mbkkzb")
    private String mbkkzb;

    @TableField("sdyxxs")
    private String sdyxxs;

    @TableField("lcz")
    private String lcz;

    @TableField("yqmc")
    private String yqmc;

    @TableField("cjzh")
    private String cjzh;

    @TableField("cd")
    private String cd;

    @TableField("scwcz")
    private String scwcz;

    @TableField("lbwd")
    private String lbwd;

    @TableField("xh")
    private String xh;

    @TableField("bz")
    private String bz;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("userid")
    private String userid;

    @TableField("htd")
    private String htd;

    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("createTime")
    private Date createtime;


}
