package com.changgou.oauth.interceptor;

import com.changgou.oauth.util.AdminToken;
import feign.RequestInterceptor;
import feign.RequestTemplate;
import org.springframework.context.annotation.Configuration;

import java.util.HashMap;


@Configuration
public class TokenRequestInterceptor implements RequestInterceptor {

    /**
     * feign执行之前拦截
     * @param template
     */
    @Override
    public void apply(RequestTemplate template) {
        //生成令牌
        String token = AdminToken.adminToken();
        System.out.println("========拦截器开始了");
//        HashMap<String, String> map = new HashMap<>();
//        map.put("Authorities", "bearer " + token);
        RequestTemplate authorization = template.header("Authorization", "bearer " + token);
        System.out.println(authorization.toString());
//        template.header(map);
    }
}
