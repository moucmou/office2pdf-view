package cn.inspur.cn.inspur.utils;

/**
 * Created by moucmou on 2017/9/6.
 */
public enum Result {
    FileIsNotExist("源文件不存在",1),CONVERTFAILER("转换失败",2),CONVERTSUCCESS("转换成功",3);
    private String cause;
    private int index;
    private Result(String cause,int index)
    {
            this.cause=cause;
            this.index=index;
    }

    public String getCause() {
        return cause;
    }

    public void setCause(String cause) {
        this.cause = cause;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }
}
