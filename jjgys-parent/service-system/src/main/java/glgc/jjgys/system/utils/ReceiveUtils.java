package glgc.jjgys.system.utils;

import com.alibaba.excel.annotation.ExcelProperty;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.system.SysMenu;
import org.springframework.beans.BeanUtils;

import javax.servlet.http.HttpServletRequest;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

public class ReceiveUtils {
    // todo: receiveHandler

    public static Map<String, String> getAnnToMap(Class<?> clz) {
        Map<String, String> kv = new HashMap<>();
        for (Field field : clz.getDeclaredFields()) {
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                kv.put(field.getName(), annotation.value()[0]);
            }
        }
        return kv;
    }

    // may have efficiency problem
    // adding a new 'insertBatch' method into the mapper may be better
    public static <T, U extends BaseMapper<T>> int createMore(List<T> data, U mapper, String userID) {
        int res = 0;
        for (T record : data) {
            T t;
            try {
                Constructor<?> constructor = record.getClass().getConstructor();
                t = (T) constructor.newInstance();
                BeanUtils.copyProperties(record, t);
                t.getClass().getMethod("setCreatetime", Date.class).invoke(t, new Date());
                t.getClass().getMethod("setUserid", String.class).invoke(t, userID);
            } catch (Exception e) {
                break;
            }
            mapper.insert(t);
            res++;
        }
        return res;
    }

    public static List<SysMenu> getProjectsTree(List<SysMenu> nodes) {
        List<SysMenu> res = nodes.stream().filter(node -> "交工管理".equals(node.getName())).collect(Collectors.toList());
        List<SysMenu> projects = res.get(0).getChildren().stream().filter(node -> !"项目信息".equals(node.getName())).collect(Collectors.toList());
        List<SysMenu> projList = new ArrayList<>();
        for (SysMenu proj : projects) {
            SysMenu p = new SysMenu();
            p.setName(proj.getName());
            List<SysMenu> collect = proj.getChildren().stream().filter(pro -> "实测数据".equals(pro.getName())).collect(Collectors.toList());
            if (!collect.isEmpty()) {
                p.setChildren(collect.get(0).getChildren());
            }
            projList.add(p);
        }
        return projList;
    }

    public static String getUserIDFromRequest(HttpServletRequest request) {
        String token = request.getHeader("token");
        return JwtHelper.getUserId(token);
    }

}
