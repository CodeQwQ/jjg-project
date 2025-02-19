package glgc.jjgys.model.system;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

@Data
@ApiModel(description = "Xmtj")
@TableName("jjg_xmtj")
public class Xmtj {
    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proName;

    @TableField("gldj")
    private String gldj;

    @TableField("tze")
    private String tze;

    @TableField("lxcd")
    private String lxcd;

    @TableField("jgnf")
    private String jgnf;

    @TableField("type")
    private String type;
}
