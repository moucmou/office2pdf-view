# office2pdf-view
  office文档通过openoffice或者microsoft多线程转换成pdf文档，并通过pdf.js显示
  
- 要是觉得前面的介绍复杂直接跳到最后    使用教程
- 希望大家能将改进的代码push请求，  尤其是多线程优化及使用office的com组件来转换ppt的解决方案  现在看到 can't save as 0^ to 1^ 心里就爆炸

# 文档在线显示：

### 1、实现方案：
	Java利用OpenOffice将office文档转换成PDF文档，再使用pdf.js实现在线预览  pdf.js
### 2、通过测试的环境及软件版本：
- 	 1.环境 windows 10:
 		
		1、OpenOffice：
 		Apache_OpenOffice_4.1.1_Win_x86_langpack_zh-CN.exe（中文语言包） 网上下的就这个版本
-	 2.office 2016:
		之前学校的学生正版  但office的ppt转pdf有点问题。。。
-	 3.jodconvert-core-3.0-beta-4
	      在本项目的lib目录下  将该文件添加到项目的依赖就行了
-	 4.jacob.jar
		jar包也在lib下
		但有两个.dll文件需要放到本地目录 java安装位置->jdk1.8->jre->bin 目录下 jacob-1.18-x64.dll
		放在resources目录下 方便你们使用

### 3、通过OpenOffice + pdf.js 实现文档转换
	使用pdf.js展示转换后的pdf 文件,其实chrome跟firfox都内置了pdf.js 只是部分功能被裁剪了
	

## 一、给定id+type+url 下载oss中的文件
- 1.模拟java.net 中的url连接直接请求下载

## 二、利用OpenOffice将Office文档转换为PDF
- 1.需要用的软件

    OpenOffice 下载地址http://www.openoffice.org/
	  office文档的安装目录   配置在了application.properties里面 根据自己安装自行修改
    JodConverter 以jar包的形式存在  本项目中lib下面的都是jodconvert 2.2.2jar包 应该是缺一不可了.  maven中央仓库没有该jar包
    由于jodconvert 2.2.2本身线程同步问题  jar包里面convert()源码没有实现考虑多线程 同步锁也只锁定了connection ，弃用
    
    但有想用的也可以问我要
     jdoconvert 2.2.2 readlock.lock() 让我直接锁死了，一次只能一个请求，其他线程全部被阻塞。。
     只到一个线程被执行完了 其他线程才有机会 差不多14秒转换一个挺小的ppt
    本项目采用jodconvert-core-3.0-beta-4  由于jar包本身支持并发 ，只需要处理线程协作即可 ，生产消费者模型，代码注释非常详细。。

- 2.单步调试Controller getfile中的各个方法就能看见转换的整个流程  代码注释很详细

## 三、利用jacob连接office com组件转换pdf

- 1.需要的工具 jacob jar包 及jacob-1.18-x64.dll

- 2.详细见代码 注释很详细

- 3.jacob使用office com组件 使用 参考文档

- //https://msdn.microsoft.com/ZH-CN/library/ff822963.aspx
- //https://msdn.microsoft.com/zh-cn/library/ff746640.aspx
  //如何使用基本的jacob 网上别人的介绍
- //http://blog.csdn.net/wkwanglei/article/details/8939277
    
## 四、  pdf.js

- 1.在github上搜索pdf.js 就能搜索到该项目  

    我已经编译完成  直接将编译后的真个文件夹放到了该项目的 resources\static\generic下
    这样就不用每个例子去找比如分页预览 或者 下载分页 什么的  这个是整个预览效果。  不然的话 官网也有效例子 一个组价一个组件的写。
    当然  做自己的权限跟功能得自己写 官网都有例子

- 2.pdf.js简单使用的话就只需要  访问 如本机  http://localhost:8080/generic/web/viewer.html?file=zz.pdf  file后面的为文件名。

- 3. 样式修改resources\static\generic\web\viewer.css  的样式就行。访问网页就是viewer.html

    viewer.js  1898行 会去读取url里面的file参数  里面的 10039行 设置的默认文档名称。 所以参数中file能直接打开要预览的文档、
- 4.pdf.js  http://mozilla.github.io/pdf.js/examples/index.html#interactive-examples  
  
     官网介绍了自己只用pdf.js跟pdfwork.js库自己完全		重写预览的界面 。
      我自己改的前端改的样式还不如原来好看 。。。  这样想要的功能都能自己写。这两个库提供了pdf读取跟解析的功能。
      通过权限的禁用打印跟 下载功能修改 viewer.html的 203行的几个button就行了。
- 5. test1.html模拟点击页面跳转到viewer.html预览页面。 在预览页面里ajax请求，获取回传参数pdf流文件即可显示在网页上。
- 6.其他html页面大多是测试页面  demo开头的是 pdf.js官网的例子 官方网站  哦 使用pdf.js 需要使用npm编译js 我将编译后的zip包放到resources下
     
## 五、 完整流程及代码大概总结

        1.controller就一个getfile 所有请求都在里面 注释一目了然、  建议调试启动 断点。

        2. com组件的使用  看代码就能了解

        3.openoffice底层通过类似com组件的 uno组件技术 。
 ## 六、 使用教程

- a.环境部署

- 1.安装openoffice 上面有版本号 ，自己确定安装目录 ，然后将application.properties中的 DocConverterUtil类中的openoffice变量改一下 安装目录 //哈哈 垃圾代码

- 2.使用openoffice需要把lib下的 jodconvert-core-3.0-beta-4 里面所有的lib放到本项目的依赖里面 这里面jacob.jar是为了使用office com组件的  也需要依赖

- 3.根据需要启动端口 使用简单生产消费者模型 DocConverterUtil类中的ports[]数组中有多少个端口模拟连接池  有多少个端口就最大默认多少个线程

- 4.pdf.js编译后的模块以压缩包的方式放到了本项目resources下build.zip pdf.js项目所在网站 https://github.com/mozilla/pdf.js 里面有案例教程 如http://mozilla.github.io/pdf.js/examples/index.html#interactive-examples
- 5.jacob jar包依赖步骤2已解决，resource下的jacob.1.18-**.dll 根据服务器版本放置到java安装位置->jdk1.8->jre->bin 目录下

- b.运行
- 1.将本工程作为maven项目导入开发工具，编译pom.xml文件
- 2.本项目采用springboot 直接右键运行 Office2h5Application文件即可
- 3.因为本项目所有的资源都放在了静态资源statices下 直接访问localhost:8080/要访问的html
- 4.采用最简单pdf.js 访问pdf  localhost:8080/viewer.html?file=文件相对目录
- 5.主要是先访问test1.html 点连接访问到viewer.html触发ajax请求 读取文件流显示 。解决跨域及文件不是网络资源

- 建议运行流程：
    //下载源文件
    1 http://localhost:8080/downloaddocx.html
    
    2.http://localhost:8080/downloadpptx.html
    
    3.http://localhost:8080/downloadxlsx.html
    //转换   用 com组件 ppt无法转换,
    
    4.http://localhost:8080/com2word.html
    
    5.http://localhost:8080/com2xlsx.html
    //用openoffice转换
    
    6.http://localhost:8080/openoffice2ppt.html
    
    //访问
    7.http://localhost:8080/test1.html
    全部测试模块大抵效果预览完毕。

- c.项目结构
- 1.所有的请求获取都在controller里面  看代码
- 2.service层服务接口
- 3.ServiceImpl 处理逻辑
- 4.utils中 DocConvertutil.java  Jodconvert调用openoffice转换文档核心 JacobCOnvert 为jacob调用com组件核心
- 5.在线预览 打开连接 localhost:8080/test1.html  核心业务。   使用com系列html的时候 如果原有pdf记得删除，不然会报文件不可读。

## 七、 问题总结

我使用 com组件ppt没能实现转换， 就是报0不能保存为1 异常， 生产消费者模型多线程协作提升效率并不明显，
起4个端口，转化速度是快一点，
但一次转换4个 要的时间太长了 ，比如一个线程转换一个文件要13秒，生产消费者转换四个文档要32秒转换池中一次转换完4个，意味着第一个和第一个一块转换完毕。
可以改为读锁 一个转换一个文件。
