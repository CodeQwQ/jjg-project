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
 * @since 2024-05-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_jtaqss_jabx_jgfc")
public class JjgFbgcJtaqssJabxJgfc implements Serializable {

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
     * 标线类型
     */
    @TableField("bxlx")
    private String bxlx;

    /**
     * 位置(ZK0+001)
     */
    @TableField("wz")
    private String wz;

    /**
     * 厚度规定值(mm)
     */
    @TableField("hdgdz")
    private String hdgdz;

    /**
     * 厚度允许偏差+
     */
    @TableField("hdyxpsz")
    private String hdyxpsz;

    /**
     * 厚度允许偏差-
     */
    @TableField("hdyxpsf")
    private String hdyxpsf;

    /**
     * 厚度实测值1
     */
    @TableField("hdscz1")
    private String hdscz1;

    /**
     * 厚度实测值2
     */
    @TableField("hdscz2")
    private String hdscz2;

    /**
     * 厚度实测值3
     */
    @TableField("hdscz3")
    private String hdscz3;

    /**
     * 厚度实测值4
     */
    @TableField("hdscz4")
    private String hdscz4;

    /**
     * 厚度实测值5
     */
    @TableField("hdscz5")
    private String hdscz5;

    /**
     * 白线逆反射系数规定值(mcd*m-2*lx-1)
     */
    @TableField("bxnfsxsgdz")
    private String bxnfsxsgdz;

    /**
     * 白线实测值1
     */
    @TableField("bxscz1")
    private String bxscz1;

    /**
     * 白线实测值2
     */
    @TableField("bxscz2")
    private String bxscz2;

    /**
     * 白线实测值3
     */
    @TableField("bxscz3")
    private String bxscz3;

    /**
     * 白线实测值4
     */
    @TableField("bxscz4")
    private String bxscz4;

    /**
     * 白线实测值5
     */
    @TableField("bxscz5")
    private String bxscz5;

    /**
     * 黄线逆反射系数规定值
     */
    @TableField("hxnfsxsgdz")
    private String hxnfsxsgdz;

    /**
     * 黄线实测值1
     */
    @TableField("hxscz1")
    private String hxscz1;

    /**
     * 黄线实测值2
     */
    @TableField("hxscz2")
    private String hxscz2;

    /**
     * 黄线实测值3
     */
    @TableField("hxscz3")
    private String hxscz3;

    /**
     * 黄线实测值4
     */
    @TableField("hxscz4")
    private String hxscz4;

    /**
     * 黄线实测值5
     */
    @TableField("hxscz5")
    private String hxscz5;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("userid")
    private Long userid;

    @TableField("fbgc")
    private String fbgc;

    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("createTime")
    private Date createtime;


}
