#mybatisplus-maven-plugin

来源：https://github.com/zouhuashan/zouhs
#####一、简介

  * [mybatis-plus](http://git.oschina.net/baomidou/mybatis-plus) 代码生成工具的 maven 插件版本
 
#####二、使用方法

  * 在项目的pom文件中配置以下内容
  
```xml
<plugin>
    <groupId>com.baomidou</groupId>
    <artifactId>mybatisplus-maven-plugin</artifactId>
    <version>1.0</version>
    <configuration>
        <!-- 输出目录(默认java.io.tmpdir) -->
        <outputDir>e:\cache</outputDir>
        <!-- 是否覆盖同名文件(默认false) -->
        <fileOverride>true</fileOverride>
        <!-- mapper.xml 中添加二级缓存配置(默认true) -->
        <enableCache>true</enableCache>
        <!-- 开发者名称 -->
        <author>Yanghu</author>
        <!-- 是否开启 ActiveRecord 模式(默认true) -->
		<activeRecord>false</activeRecord>
        <!-- 数据源配置，( **必配** ) -->
        <dataSource>
            <driverName>com.mysql.jdbc.Driver</driverName>
            <url>jdbc:mysql://127.0.0.1:3306/demo?useUnicode=true&amp;useSSL=false</url>
            <username>root</username>
            <password>123456</password>
        </dataSource>
        <strategy>
			<!-- 字段生成策略，四种类型，从名称就能看出来含义： 
				nochange(默认), 
				underline_to_camel,(下划线转驼峰) 
				remove_prefix,(去除第一个下划线的前部分，后面保持不变) 
				remove_prefix_and_camel(去除第一个下划线的前部分，后面转驼峰) -->
			<naming>remove_prefix_and_camel</naming>
			<!-- 表前缀 -->
			<tablePrefix>bmd_</tablePrefix>
            <!--Entity中的ID生成策略（默认 id_worker）-->
            <idGenType>uuid</idGenType>
            <!--自定义超类-->
            <!--<superServiceClass>com.baomidou.base.BaseService</superServiceClass>-->
            <!-- 要包含的表 与exclude 二选一配置-->
            <!--<include>-->
                <!--<property>sec_user</property>-->
                <!--<property>table1</property>-->
            <!--</include>-->
            <!-- 要排除的表 -->
            <!--<exclude>-->
                <!--<property>schema_version</property>-->
            <!--</exclude>-->
        </strategy>
        <packageInfo>
            <!-- 父级包名称，如果不写，下面的service等就需要写全包名(默认com.baomidou) -->
            <parent>com.baomidou</parent>
            <!--service包名(默认service)-->
            <service>service</service>
            <!--serviceImpl包名(默认service.impl)-->
            <serviceImpl>service.impl</serviceImpl>
            <!--entity包名(默认entity)-->
            <entity>entity</entity>
            <!--mapper包名(默认mapper)-->
            <mapper>mapper</mapper>
            <!--xml包名(默认mapper.xml)-->
            <xml>mapper.xml</xml>
        </packageInfo>
        <template>
        	<!-- 定义controller模板的路径 -->
			<!--<controller>/template/controller1.java.vm</controller>-->
		</template>
    </configuration>
    <dependencies>
        <dependency>
            <groupId>mysql</groupId>
            <artifactId>mysql-connector-java</artifactId>
            <version>${mysql.version}</version>
        </dependency>
    </dependencies>
</plugin>
```

# 执行
  * 下载源码的同学！可使用mvn clean install将自定义的这个插件安装到本地仓库。
	执行：mvn clean install

	命令：mvn com.baomidou:mybatisplus-maven-plugin:1.0:code

  * 显然这个命令太长了，使用很不方便，可在settings.xml中配置如下：
	```
  	<pluginGroups>
		<pluginGroup>com.baomidou</pluginGroup>
	</pluginGroups>
	```
  * 然后使用简单命令：mvn mp:code

  
#####二、参数说明
  
  
  
  
