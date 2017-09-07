package cn.inspur.serviceImpl;


import cn.inspur.service.IsinLocal;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


/**
 * Created by moucmou on 2017/9/4.
 */
@Service
public class IsinLocalImpl implements IsinLocal{
    private static Logger logger = LoggerFactory.getLogger(IsinLocalImpl.class);
    @Value("${download-path}")
    private  String path;

    /**
     * 判断文件是否在本地  如果不在本地就下载
     * @param id
     * @param type
     * @param url
     * @return  -1表示给的url不能下载  其他字符串表示文件完整路径  这是段硬代码 哈哈
     */
    @Override
    public String isIn(String id, String type, String url) {
        //判断本地是否有该文件 没有的话就下载
        String signal=find(id,type,url);
        if(signal==null||signal.equals(""))
        {
            try {
                DowmFileImpl.downLoadFromUrl(path,id,type,url);
            } catch (IOException e) {
//                e.printStackTrace();
                return "-1";
            }
        }
        return signal;
    }

    /**
     *

     * @param id
     * @param type
     * @param url
     * @return 文件路径
     */
    public  String find(String id,String type,String url){
        //文件完整路径
        String filePath =path +id+"."+type;
        //获取pathName的File对象
        File dirFile = new File(filePath);
        //判断该文件或目录是否存在，不存在时在控制台输出提醒
        if (!dirFile.exists()) {
            logger.info(">>>>>>>文件不存在");
            return null;
        }

        return filePath;
    }

    /**
     * 判断id的pdf文件是否存在
     * @param path
     * @param id
     * @param type
     * @param url
     * @return
     */
    public static boolean findPdf(String path,String id,String type,String url){

        String filePath =path  +id+"."+"pdf";
        //获取pathName的File对象
        File dirFile = new File(filePath);
        //判断该文件或目录是否存在，不存在时在控制台输出提醒
        if (!dirFile.exists()) {
            logger.info(">>>>>>>文件不存在");
            return false;
        }
        return true;
    }
}
