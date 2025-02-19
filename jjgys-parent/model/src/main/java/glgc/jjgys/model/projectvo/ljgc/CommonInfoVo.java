package glgc.jjgys.model.projectvo.ljgc;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class CommonInfoVo implements Serializable {

    private static final long serialVersionUID = 1L;

    private String proname;
    private String htd;
    private String fbgc;
    private String sjz;
    private String sdsjz;
    private String qlsjz;
    private String fhlmsjz;
    private String username;

    //雷达厚度
    private String yfsjz;
    private String yffhlmsjz;
    //平整度
    private String zlsjz;
    private String zssjz;
    private String zqsjz;
    private String hlsjz;
    private String hssjz;
    private String hqsjz;
    private String llsjz;
    private String lssjz;
    private String lqsjz;

    @JsonFormat(shape = JsonFormat.Shape.STRING)
    private String userid;
}
