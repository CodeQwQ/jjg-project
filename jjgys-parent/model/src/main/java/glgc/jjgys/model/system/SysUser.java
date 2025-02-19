package glgc.jjgys.model.system;

import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableName;

import com.fasterxml.jackson.annotation.JsonFormat;
import glgc.jjgys.model.base.BaseEntity;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import org.springframework.format.annotation.DateTimeFormat;

import java.time.LocalDateTime;
import java.util.List;

@Data
@ApiModel(description = "用户")
@TableName("sys_user")
public class SysUser extends BaseEntity {
	
	private static final long serialVersionUID = 1L;

	@ApiModelProperty(value = "用户名")
	@TableField("username")
	private String username;

	@ApiModelProperty(value = "密码")
	@TableField("password")
	private String password;

	@ApiModelProperty(value = "姓名")
	@TableField("name")
	private String name;

	@ApiModelProperty(value = "手机")
	@TableField("phone")
	private String phone;


	@ApiModelProperty(value = "部门id")
	@TableField("dept_id")
	private Long deptId;

	@ApiModelProperty(value = "岗位id")
	@TableField("post_id")
	private Long postId;

	@ApiModelProperty(value = "公司id")
	@TableField("company_id")
	private Long companyId;

	@ApiModelProperty(value = "用户类型 1 超级管理员，2 公司管理员，3 普通用户")
	@TableField("type")
	private String type;

	@ApiModelProperty(value = "描述")
	@TableField("description")
	private String description;

	@TableField("expire")
	@DateTimeFormat(pattern = "yyyy-MM-dd hh:mm:ss")
	@JsonFormat(pattern = "yyyy-MM-dd HH:mm:ss", timezone = "GMT+8")
	private LocalDateTime expire;


	@TableField(exist = false)
	private List<SysRole> roleList;
	//岗位
	@TableField(exist = false)
	private String postName;
	//部门
	@TableField(exist = false)
	private String deptName;
}

