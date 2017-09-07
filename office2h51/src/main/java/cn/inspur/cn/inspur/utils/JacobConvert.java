package cn.inspur.cn.inspur.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.Enumeration;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

/**
 * Created by moucmou on 2017/9/1.
 */


public class JacobConvert {
    private static Lock  lock=new ReentrantLock();
    private static final int wdExportFormatPDF = 17;    //msdn里面的枚举值
    private static final int wdExportFormatXPS=  18;
    private static final int xlTypePDF = 0;
    private static final int pptSaveAsPDF = 32;
    private static Logger logger= LoggerFactory.getLogger(JacobConvert.class);
//    private static ExecutorService taskManager = Executors.newCachedThreadPool();
    // 直接调用这个方法即可  自己去起线程的时候
//    private  String old;
//    private  String want;
//    public JacobConvert(String old,String want)
//    {
//        this.old=old;
//        this.want=want;
//        convert2PDF(old,want);
//    }

    /**
     * 判断格式转换的
     * @param inputFile
     * @param pdfFile
     * @return
     */
    public static boolean  convert2PDF(String inputFile, String pdfFile) {
        //判断文件类型
        int splitIndex = inputFile.lastIndexOf(".");
        String suffix=inputFile.substring(splitIndex + 1);
        if (suffix.equals("pdf")) {
            System.out.println("PDF not need to convert!");
            return false;
        }
        if (suffix.equals("doc") || suffix.equals("docx")
                || suffix.equals("txt")) {
            return word2PDF(inputFile, pdfFile);
        } else if (suffix.equals("ppt") || suffix.equals("pptx")) {
            return ppt2PDF(inputFile, pdfFile);
        } else if (suffix.equals("xls") || suffix.equals("xlsx")) {
            return excel2PDF(inputFile, pdfFile);
        } else {
            System.out.println("文件格式不支持转换!");
            return false;
        }
    }

    /**
     * word转换没有问题
     * @param inputFile
     * @param pdfFile
     * @return
     */
    public static boolean word2PDF(String inputFile, String pdfFile) {
        //https://msdn.microsoft.com/ZH-CN/library/ff822963.aspx  Document对象 即call里面能调用的参数
        ActiveXComponent app=null;
        Dispatch doc=null;
        ComThread.InitSTA();
            try {
                Long old=System.currentTimeMillis();
                //起线程方便后面结束

                // 打开word应用程序  文档类型   http://men4661273.iteye.com/blog/2097871  基本demo 不懂jacob大概要看看
                 app = new ActiveXComponent("Word.Application");
                // 设置word不可见
                app.setProperty("Visible", false);
                // 获得word中所有打开的文档,返回Documents对象
                Dispatch docs = app.getProperty("Documents").toDispatch();
                // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
                //call参数 ：对象，方法，参数 参考 http://blog.csdn.net/wkwanglei/article/details/8939277  可以看看
                 doc = Dispatch.call(docs, "Open", inputFile, false, true)
                        .toDispatch();
                logger.info("debug>>>>>>>>打开文件成功");
                // 调用Document对象的SaveAs方法，将文档保存为pdf格式
			/*
			 * Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF
			 * //word保存为pdf格式宏，值为17 );
			 */
                //方法在 https://msdn.microsoft.com/zh-cn/library/microsoft.office.tools.word.document  上面一个个找  看上哪个选哪个  上面的方法
                Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, wdExportFormatPDF // 保存后的枚举值 wdExportFormatPDF 17，wdExportFormatPDF 18
                );
                Long current=System.currentTimeMillis();
                // 关闭文档

                logger.info("debug>>>>>>>>转换成功,耗时"+(current-old));
                return true;
            } catch (Exception e) {
                e.printStackTrace();
                return false;
            } finally {

                Dispatch.call(doc, "Close", false);
                // 关闭word应用程序
                app.invoke("Quit", 0);
                ComThread.Release();
                logger.info("debug>>>>>>>>关闭线程成功");
            }
    }

    /**
     * excel转换没有问题
     * @param inputFile
     * @param pdfFile
     * @return
     */
    public static  boolean excel2PDF(String inputFile, String pdfFile) {
        System.out.println("当前进程名称"+Thread.currentThread().getName());
        ActiveXComponent app=null;
        Dispatch excel=null;
        try {
            Long old=System.currentTimeMillis();
                ComThread.InitSTA();
                 app= new ActiveXComponent("Excel.Application");
                app.setProperty("DisplayAlerts", false);
                app.setProperty("Visible", false);
                Dispatch excels = app.getProperty("Workbooks").toDispatch();
            System.out.println("获取Workbooks成功");
                 excel = Dispatch.call(excels, "Open", inputFile, false,
                        true).toDispatch();
            System.out.println("打开文件成功");
//                Dispatch.call(excel, "ExportAsFixedFormat", 0, pdfFile);
                Dispatch.call(excel, "SaveAs", pdfFile, new Variant(57));
            logger.info("文件转换成功");
                Dispatch.call(excel, "Close", false);
            Long current=System.currentTimeMillis();
            logger.info("debug>>>>>>>>转换成功,耗时"+(current-old));
                return true;
            } catch (Exception e) {
                e.printStackTrace();
                return false;
            } finally {
            app.invoke("Quit");
                ComThread.Release();
            System.out.println("关闭成功");
            }
    }

    /**
     * ppt转换      自求多福
     * @param inputFile
     * @param pdfFile
     * @return
     */
    public  static boolean ppt2PDF(String inputFile, String pdfFile) {

        // TODO: 2017/9/6 office2016转不了  查了各种文档
        //https://msdn.microsoft.com/zh-cn/library/ff746640.aspx
        // 这是函数库。 试过 saveas savecopyas ExportAsFixedFormat 都没用
        //ExportAsFixedFormat 可能希望大一点 但我官网的只需要两个必要参数，我传过去了还是错误。。
        //  如果你单步调试出现了 Presentation.SaveAs: PowerPoint can not save ^ 0 to ^ 1. 恭喜你，我挖的坑得你来填了。。
        //  上面的文档就是函数参考 即call的第一个参数（函数名字）， 百度 谷歌搜。 然后你大抵就能明白我的绝望。。

        ActiveXComponent app =null;
        Dispatch ppt=null;
        ComThread.InitSTA();
        //真的，这个问题我解决不了了，本来这个jacob调用com组件就看不到组件代码，调试看不到，都不知道在哪出的错。方法里面有好多文档，说不定你们就解决了。
        logger.info("真的，  PowerPoint 无法将 ^0 保存到 ^1。  这个问题我解决不了了，本来这个jacob调用com组件就看不到组件代码，调试看不到，都不知道在哪出的错。方法里面有好多文档，说不定你们就解决了。");
        try {
            //上不上锁都没什么区别。。
            lock.lock();
            app = new ActiveXComponent(
                    "PowerPoint.Application");
            //ppt不能被隐藏窗口  。。。。
//            app.setProperty("Visible", new Variant(false));
//            app.setProperty("Visible", false);
            logger.info("debug>>>>>>>>成功打开ppt");
            Dispatch  ppts= app.getProperty("Presentations").toDispatch();
//            ActiveXComponent presentations = app.getPropertyAsComponent("Presentations");

             ppt = Dispatch.call(ppts, "Open", inputFile, true,// ReadOnly
                    true,// Untitled指定文件是否有标题
                    false// WithWindow指定文件是否可见

            ).toDispatch();
             logger.info("debug>>>>>>>>>读取文件成功");
//
//            Dispatch.call(ppt, "SaveAs", pdfFile, pptSaveAsPDF);
//
            Dispatch.invoke(ppt, "SaveAs", Dispatch.Method, new Object[] {pdfFile , new Variant(32)},new int[1]);
            Dispatch.call(ppt,"SaveCopyAs",pdfFile,new Variant(32));
            //下面这个函数在上面的msdn网址上都能找到 这个被介绍为
           // ExportAsFixedFormat方法等效于在 PowerPoint 用户界面Office菜单上的另存为 PDF 或 XPS命令。该方法将创建一个包含当前演示文稿的静态视图文件。
            //FixedFormatType参数值可以是其中一个PpFixedFormatType常量。
            //没用 好像是我参数拼错了，  但我觉得就两个参数必须的 参数也没问题啊、、  大佬们解决了告诉我下  71610506@qq.com  一定要记得。。。
//            Dispatch.call(ppt,"ExportAsFixedFormat",pdfFile,new Variant(2));

            logger.info("debug>>>>>>>>>转换成功");
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            lock.unlock();
            Dispatch.call(ppt, "Close");
            app.invoke("Quit");
            ComThread.Release();
        }

    }


}