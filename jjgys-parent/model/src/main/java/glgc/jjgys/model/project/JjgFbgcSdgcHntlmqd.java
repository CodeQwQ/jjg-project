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
 * @since 2023-05-04
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_sdgc_hntlmqd")
public class JjgFbgcSdgcHntlmqd implements Serializable {

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
     * 隧道名称
     */
    @TableField("sdmc")
    private String sdmc;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 试样平均直径(mm)
     */
    @TableField("sypjzj")
    private String sypjzj;

    /**
     * 试样平均厚度(mm)
     */
    @TableField("sypjhd")
    private String sypjhd;

    /**
     * 极限荷载(N)
     */
    @TableField("jxhz")
    private String jxhz;

    /**
     * 路面强度规定值
     */
    @TableField("lmqdgdz")
    private String lmqdgdz;

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
