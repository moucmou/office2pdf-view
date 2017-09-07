package cn.inspur.serviceImpl;

import cn.inspur.cn.inspur.utils.DocConverterUtil;
import cn.inspur.cn.inspur.utils.JacobConvert;
import cn.inspur.cn.inspur.utils.Result;
import cn.inspur.service.ConvertToPdf;
import org.springframework.stereotype.Service;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Created by moucmou on 2017/8/25.
 */
@Service
//@Scope("prototype")
public class ConvertToPdfImpl implements ConvertToPdf {
    private static ExecutorService taskManager = Executors.newCachedThreadPool();


    @Override
    public Result convet2pdf(String path) {
            //拼出目标文件路径
        String target=path.split("\\.")[0]+".pdf";

        return   DocConverterUtil.officeToPdf(path,target);


    }
    public  boolean convet2pdfCom(String path)
    {
        //拼目标pdf目录
        String target=path.split("\\.")[0]+".pdf";

        // springmvc默认起6个线程  不爽的话就启用下面的改改、
        //        taskManager.execute(new ConvertBycom(sourceFilePath,pdfFilePath));
        //线程内部类变量  线程安全
//        JacobConvert jacobConvert=new JacobConvert(sourceFilePath,pdfFilePath);
//        return jacobConvert.convert2PDF(sourceFilePath,pdfFilePath);
            return JacobConvert.convert2PDF(path,target);
    }
        //自己起线程  但springmvc不需要 默认单例多线程
//    private class ConvertBycom implements Runnable{
//        String old=null;
//        String want=null;
//        public ConvertBycom(String old,String want)
//        {
//            this.old=old;
//            this.want=want;
//        }
//        @Override
//        public void run() {
//            String old=this.old;
//            String want=this.want;
//            JacobConvert jacobConvert=new JacobConvert(old,want);
//        }
//    }



}
