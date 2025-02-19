package glgc.jjgys.model.system;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.util.Date;


@Data
@ApiModel(description = "Project")
@TableName("jjg_projectinfo")
public class Project{

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proName;

    @TableField("xmqc")
    private String xmqc;

    @ApiModelProperty(value = "公路等级")
    @TableField("grade")
    private String grade;

    @TableField("participant")
    private String participant;

    @TableField("tze")
    private String tze;

    @TableField("lxcd")
    private String lxcd;

    @TableField(exist = false)
    private String username;

    @TableField("create_time")
    private Date createTime;

    @JsonFormat(shape = JsonFormat.Shape.STRING)
    @TableField("userid")
    private String userid;
}
