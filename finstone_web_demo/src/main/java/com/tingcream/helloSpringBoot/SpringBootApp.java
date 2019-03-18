package com.tingcream.helloSpringBoot;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.support.SpringBootServletInitializer;
  
@SpringBootApplication
//@SpringBootConfiguration  // the same as @Configuration
  
public class SpringBootApp extends SpringBootServletInitializer{
 
    //实现SpringBootServletInitializer接口 ，需实现此方法。  可以war包形式发布到tomcat中 
    @Override
    protected SpringApplicationBuilder configure(
            SpringApplicationBuilder builder) {
         return builder.sources(SpringBootApp.class);
    }
 
    public static void main(String[] args) {
        SpringApplication.run(SpringBootApp.class, args);
    }
 
}