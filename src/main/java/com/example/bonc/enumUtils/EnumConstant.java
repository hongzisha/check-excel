package com.example.bonc.enumUtils;


/**
 * @author TanHao
 * @date 2022/10/22 0022
 */

public enum EnumConstant {

    /**
     * 考勤数据
     */
    ATTENDANCE("1"),
    /**
     * 周报数据
     */
    WEEK("2"),
    /**
     * 2003年excel格式
     */
    FORMAT1("xls"),
    /**
     * 2007年excel格式
     */
    FORMAT2("xlsx");


    /**
     * 编码
     */
    private final String code;

    public String getCode() {
        return code;
    }


    EnumConstant(String code) {
        this.code = code;
    }
}
