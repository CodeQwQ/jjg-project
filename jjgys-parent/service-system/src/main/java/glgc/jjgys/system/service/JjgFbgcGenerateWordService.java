package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;


/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
public interface JjgFbgcGenerateWordService extends IService<Object> {


    boolean generateword(String proname) throws IOException, InvalidFormatException;
}
