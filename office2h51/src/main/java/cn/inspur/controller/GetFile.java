package cn.inspur.controller;

import cn.inspur.service.ConvertToPdf;
import cn.inspur.service.DownFile;
import cn.inspur.serviceImpl.IsinLocalImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.InetAddress;
import java.net.UnknownHostException;

/**
 * Created by moucmou on 2017/8/25.
 */
@RestController
public class GetFile {
    @Autowired
    DownFile downFile;
    @Autowired
    ConvertToPdf convertToPdf;
    @Autowired
    IsinLocalImpl isinLocal;
    //从配置文件中读取文件要保存的目录
    @Value("${download-path}")
    private String path;
    /**
     * 之后勇哥给的标准 给的参数是id  type url
     * @param id
     * @param type
     * @param url
     * @return
     */
    @RequestMapping(value="/down",method= RequestMethod.POST)
//    @ResponseBody
    public boolean downFile(@RequestParam("id")String id, @RequestParam("type")String type,@RequestParam("url")String url )
    {
        boolean sad=true;
        boolean zz= downFile.getFile(id,type,url);
//
        return zz;
    }


    /**
     *
     * @param key  这个接口留着  主要是使用下面的调用回传参数的接口就行了 key就是名字加后缀 如 1.ppt
     * @return
     */
    @RequestMapping(value="/convert")
    public String convet2pdf(@RequestParam("key") String key)
   {
       String paths=path+key;
       //工具类里定义了该枚举类型   openoffice转换不是很快 即便多线程 7s转换一个ppt把。。
       return convertToPdf.convet2pdf(paths).getCause();
   }

    @RequestMapping(value="/convertByCom")
//    @ResponseBody
    public boolean convet2pdfByCom(@RequestParam("key") String key)
    {
        String paths=path+key;
        return convertToPdf.convet2pdfCom(paths);
    }
    /**
     * pdf.js 简单实用的话就是  viewer.html?file=path   拼个字符串就行了  实际作为服务的时候意味着 目标.pdf最好跟目录viewer.js放得不远
     * @param key key为文件名 带后缀的。
     * @return
     * @throws UnknownHostException
     */
   @RequestMapping("/view")
    public String viewOnloine(@RequestParam("key")String key) throws UnknownHostException {
        System.out.println(InetAddress.getLocalHost().getHostAddress());
        //pdf.js 最简单的用法  不要注释掉viewer.js中的10038行默认url声明  这样就可以直接拼个网络上的文件就可以访问了  注意是网络上的。
        return "localhost:8080/viewer.html?file="+"doc/"+key;
    }

    /**
     * 这是通过浏览 viewer.html 来请求服务器的文件流填充显示pdf页面 但好像没大这么用 这样可以直接读远程文件流来解决跨域问题  可以直接读取远程.pdf文件流来显示
     * 但这样就必须要直接访问  viewer.html页面 来更改显示文件的参数
     * @param key
     * @param request
     * @param response
     */
    @RequestMapping(value="/show")
    public void  pdfView(@RequestParam("key")String key, HttpServletRequest request, HttpServletResponse response) {
        String basePath = System.getProperty("user.dir").replace("\\", "/") ;
//        System.out.println(basePath);
        //文件目录  自己拼  因为我之前把下载的文件放在了 /doc/下了
//        String sourceFilePath = basePath +"\\src\\main\\resources\\static\\doc\\"+key;
       String filePath =path+key.split("\\.")[0]+".pdf";
        File file = new File(filePath);
        byte[] data = null;
        FileInputStream input=null;
        response.setStatus(HttpServletResponse.SC_OK);
        response.setContentType("application/pdf;charset=UTF-8");
        try {
             input = new FileInputStream(file);
            data = new byte[input.available()];
            input.read(data);

            response.getOutputStream().write(data);
//            input.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    /**
     * 文件在线预览  viewer ajax请求的地址 static目录下test1.html 模拟点击请求预览文件 然后跳到viewer.html预览结果
     * 这样读取流文件就即可以解决psf.js跨域问题 也能解决pdf.js不能访问本地文件的弊端。。
     * @param id 文件唯一标识
     * @param type 文件后缀 没有点 "."
     * @param url //下载文件的连接
     * @param response
     */
    @RequestMapping("/viewOnline")
    public void viewOnLine(@RequestParam("id")String id,@RequestParam("type")String type,@RequestParam("url")String url,
                           HttpServletResponse response) {
        //文件编码 防止乱码
        response.setContentType("application/pdf;charset=UTF-8");
        byte[] data = null;
        //判断文件本地是否存在
        String str = isinLocal.isIn(id, type, url);
        //判断该文件的id的pdf文件是否存在
        Boolean hasPdf= IsinLocalImpl.findPdf(path,id,type,url);
        //硬代码逻辑 哈哈  得改。。
        //如果类型不为pdf 且本地有源文件（id+type）了
        if (type != "pdf" && !str.equals("-1"))
        {
            //如果本地没有对应的pdf文F件 就启动转换
            if(!hasPdf){
                //转换失败
                if( convertToPdf.convet2pdf(id + "." + type).getIndex()!=3)
                {
                    System.out.println("转换失败");
                    return;   //直接返回 本次请求失败
                }
            }
            //转换完成后本地就会生成pdf文件 所以直接读本地的就行
            String fileStr = str.split("\\.")[0] + ".pdf";
            File file = new File(fileStr);
            FileInputStream fileInputStream = null;
            try {
                //读取文件
                fileInputStream = new FileInputStream(file);
                //创建数组
                data = new byte[fileInputStream.available()];
                //将文件流写到字节数组中
                fileInputStream.read(data);
                //写入response中，前端ajax成功获取到回传参数。
                response.getOutputStream().write(data);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


        }
    }
}
