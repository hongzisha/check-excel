package com.example.bonc.entity;

/**
 * @author TanHao
 * @date 2022/10/19 0019
 */
public class ResultObject<T> {
    /**
     * 返回编码
     **/
    private String code = "0000";
    /**
     * 返回信息
     **/
    private String msg;
    /**
     * 数据
     */
    private T data;

    public ResultObject(String code, String msg, T data) {
        this.code = code;
        this.msg = msg;
        this.data = data;
    }

    public ResultObject() {
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }
}
