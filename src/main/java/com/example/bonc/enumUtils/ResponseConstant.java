package com.example.bonc.enumUtils;

/**
 * @author TanHao
 * @date 2022/10/19 0019
 */
public enum ResponseConstant {
    /**
     * 成功
     */
    SUCCESS("0000", "成功"),

    /**
     * 失败
     */
    FAILURE("0001","失败");

    /**
     * 错误编码
     */
    private final String code;
    /**
     * 错误信息
     */
    private final String msg;

    public String getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }

    ResponseConstant(String code, String msg) {
        this.code = code;
        this.msg = msg;
    }
}
