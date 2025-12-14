package com.changgou.token;

import org.junit.Test;
import org.springframework.security.jwt.Jwt;
import org.springframework.security.jwt.JwtHelper;
import org.springframework.security.jwt.crypto.sign.RsaVerifier;

/*****
 * @Author: www.itheima
 * @Date: 2019/7/7 13:48
 * @Description: com.changgou.token
 *  使用公钥解密令牌数据
 ****/
public class ParseJwtTest {

    /***
     * 校验令牌
     */
    @Test
    public void testParseToken(){
        //令牌
        String token = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJuYW1lIjoiY2hhbmdnb3UiLCJpZCI6IjEiLCJhdXRob3JpdGllcyI6WyJ2aXAiLCJhZG1pbiIsInVzZXIiXX0.XbESfhOJUxE5LXIgGRgwHx6DecgZINCdezsuIkWm7zAuPww37oOuYdpDSh0Pf3x3gFWQIV70UAlgkoGwDrXwTNy3R-1ONSzH6iOZBy5SmXd179EZeAcufJ0x8TFgc4emtcsOUk_SleDHasEZm1V_D_prAgB2HfPKTA5yfeNGhTyAp0rtTzSyAdkc5vJAv7O3oWTVlBxzaWGDrCSIli8w0iqIFU0yKbF26k_iiSpWgc49DpgdplTvLsO5jfwdOGjHJtK-QNEesBip8Lce3hVHKrwJG-ii03-97ub9BvJ2-RQ6xN0fK3X26bNBAOfquNGEpTwA_-dSlnFa1e7JGIxniQ";

        //公钥
        String publickey = "-----BEGIN PUBLIC KEY-----MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAmtSrLmMI5Z8GJ2qtRrCwgOypcjDXUZIs3EqAeSkmzYsv4Ph+ooYGoKz87a3AWqq7NYlJbW0ZRSDgnjXXLi1DCDNvXsrXVB50fVdhqXWx1Vj20/flBzYaHw1f8813AwlK2rDg9H1LcdrcaCp81IS2dXfTuLnyb1deHxn6AM3jjnCzP3Q/tTCAOzlcOJnkkWMfRdkCjtz/jFZ6swMS5heIaGcgwYO9PT2t7zwH7dMuLwKuvIeMG3ZJb4hvBOMxGptiyyY0ozzasQWfNe/Xsjq2KUdZ2lSYszxMOX1Wxnseq9f6o9c0uloUEJGCaZyRfJRtn2jULb86IFaRvRFU7Gz4XQIDAQAB-----END PUBLIC KEY-----";

        //校验Jwt
        Jwt jwt = JwtHelper.decodeAndVerify(token, new RsaVerifier(publickey));

        //获取Jwt原始内容 载荷
        String claims = jwt.getClaims();
        System.out.println(claims);

//        System.out.println(encoded);
    }
}
