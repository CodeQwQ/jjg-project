package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-08-30
 */
@Data
@TableName("jjg_planinfo")
public class JjgPlaninfo  {


    @TableId("id")
    private String id;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("zb")
    private String zb;

    @TableField("num")
    private String num;

    @TableField("create_time")
    private Date createTime;


}
