package glgc.jjgys.system.controller;

import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.system.Xmtj;
import glgc.jjgys.system.service.JjgXmtjService;
import io.swagger.annotations.Api;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.*;

import java.util.List;
@Api(tags = "项目统计接口")
@RestController
@Transactional
@RequestMapping(value = "/xmtj")
@CrossOrigin//解决跨域
public class JjgXmtjController {

    @Autowired
    private JjgXmtjService jjgXmtjService;

    @GetMapping("findAll")
    public Result findAllProject(){
        List<Xmtj> list = jjgXmtjService.list();
        return Result.ok(list);
    }
    /**
     * 分页查询
     */
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit
                                ){
        //创建page对象
        Page<Xmtj> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部

        IPage<Xmtj> pageModel = jjgXmtjService.page(pageParam,null);
        return Result.ok(pageModel);



        }

}
