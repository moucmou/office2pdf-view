package cn.inspur.serviceImpl;

import cn.inspur.cn.inspur.utils.DocConverterUtil;
import cn.inspur.service.DownFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;

/**
 * Created by moucmou on 2017/8/25.
 */
@Service
public class DowmFileImpl implements DownFile {
    private static Logger logger = LoggerFactory.getLogger(DowmFileImpl.class);
    @Value("${download-path}")
    String path;
    @Override
    public  boolean getFile(String id, String type,String url) {

//        return downLoadFromUrl(urlStr, key);
        try {
            downLoadFromUrl(path,id, type, url);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;

    }

    public static boolean downLoadFromUrl(String path,String id, String type,String urlStr) throws IOException {
        InputStream inputStream = null;
        FileOutputStream fos = null;
        boolean stat = false;

        URL url = new URL(urlStr);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        //设置超时间为3秒
        conn.setConnectTimeout(3 * 1000);
        //防止屏蔽程序抓取而返回403错误
        conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");
        //得到输入流
        inputStream = conn.getInputStream();
        //获取自己数组
        byte[] getData = readInputStream(inputStream);


        //  自己根据需要修改  测试就直接写到本项目的doc目录下



        File saveDir = new File(path);
        if (!saveDir.exists()) {
            saveDir.mkdir();
        }
        File file = new File(saveDir + File.separator + id+"."+type);
        fos = new FileOutputStream(file);
        fos.write(getData);
        stat = true;
        logger.info("下载是否成功>>>>>>>>>"+stat);

        if (fos != null) {
            fos.close();
        }
        if (inputStream != null) {
            inputStream.close();
        }

        return stat;

    }

    /**
     * 从输入流中获取字节数组
     *
     * @param inputStream
     * @return
     * @throws IOException
     */
    public static byte[] readInputStream(InputStream inputStream) throws IOException {
        byte[] buffer = new byte[1024];
        int len = 0;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((len = inputStream.read(buffer)) != -1) {
            bos.write(buffer, 0, len);
        }
        bos.close();
        return bos.toByteArray();
    }


}
