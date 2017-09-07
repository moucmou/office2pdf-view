package cn.inspur.cn.inspur.utils;

/**
 * Created by moucmou on 2017/9/6.
 */

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import javax.annotation.PostConstruct;
import java.io.File;
import java.util.concurrent.BlockingQueue;

import java.util.concurrent.LinkedBlockingQueue;


//注册为bean 主要是为了使用@PostConstruct注解 让在初始化ioc容器后就启动openoffice
@Component
public class DocConverterUtil {
	//线程安全 生产消费者线程协作
	public static BlockingQueue<Integer> queue = new LinkedBlockingQueue<>();
	private static Logger logger = LoggerFactory.getLogger(DocConverterUtil.class);


//	@Value("${openoffice}")
//	String openoffice;
	//openOffice安装路径
	private static String openoffice="C://Program Files (x86)//OpenOffice_4" ;
	//服务端口
	private static int ports[] = {8100,8101,8102,8103};
	@PostConstruct
	private void init()
	{	//初始化队列
		for(int port:ports)
			queue.offer(port);
	}

	//构造函数 启动程序后就把进程启动。
	public static OfficeManager startService(int port){
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		OfficeManager officeManager=null;
		try {

			//设置安装目录
			configuration.setOfficeHome(openoffice);
			//从线程安全的队列中取一个端口号
			configuration.setPortNumbers(port); //设置端口
			//任务超时 5分钟
			configuration.setTaskExecutionTimeout(1000 * 60 * 5L);
			//设置任务队列超时为24小时
			configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);
			//最大进程数
//			configuration.setMaxTasksPerProcess(1);
			//创建openoffice管理
			officeManager = configuration.buildOfficeManager();
			//启动服务
			officeManager.start();

		} catch (Exception e) {
			System.out.println("office转换服务启动失败!详细信息:" +e);
		}
		return officeManager;
	}

	public static Result officeToPdf(String source, String target){

		logger.info(">>>>>>office源文件路径："+source);
		File sourceFile = new File(source);
		//判断文件是否存在
		if(!sourceFile.exists()) {
			logger.error("源文件不存在："+source);
			return Result.FileIsNotExist;
		}

		// 如果源文件存在且目标文件不存在 尽量让要转换的文档都是需要转换的，就比如本来就是pdf或者已经有该文件的pdf了就不要触发转换文档模块
		//默认没有pdf文件存在
		if(target == null || "".equals(target.trim())) {
			//根据原文件目录拼出目标文件目录
			target = source.substring(0, source.lastIndexOf(".") + 1)+"pdf";
		}
		logger.info(">>>>>>pdf目标文件路径："+target);
		File destFile = new File(target);
		//判断目标文件在文件夹是否存在，不存在说明目标文件也不存在，则需要创建目标文件所在的目录
		if(!destFile.getParentFile().exists()) {
			//创建文件夹
			destFile.getParentFile().mkdirs();
		}

		// 获取源文件后缀名，便于后面转换时根据不同类型的文件进行适配转换格式化
		String fileExt = "";
		String fileName = sourceFile.getName();
		int i = fileName.indexOf(".");
		if (i != -1) {
			 fileExt = fileName.substring(i + 1);
	    }
		int port=0;
		OfficeManager officeManager=null;
		try {

			//当队列中没有端口号时 线程阻塞 直到某个线程转换完毕
			port=queue.take();

			//获取当前时间
			long startTime = System.currentTimeMillis();
			logger.info(">>>>>>开始转换...");
			officeManager=DocConverterUtil.startService(port);
			//不同版本的office文档处理
			OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);

			//显然也是可以直接由输入流来转换、我的需求是先把文件下载到本地，如果后期你们改  可以参照这个方法来
			//我测试的office文档都是走这条分支、自动获取office后缀名

			converter.convert(sourceFile, destFile);
			long old = System.currentTimeMillis();
	    	logger.info(">>>>>>转换成功！"+"本次转换耗时："+(old-startTime));
		} catch (Exception e) {
			logger.error(">>>>>>>文档"+source+"转换失败，原因："+e);
			//返回失败
			return Result.CONVERTFAILER;
		}finally {
			officeManager.stop();
			queue.offer(port);
		}
		//返回成功
		return Result.CONVERTSUCCESS;
	}

}
