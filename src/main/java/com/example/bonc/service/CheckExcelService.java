package com.example.bonc.service;

import com.example.bonc.entity.Check;
import com.example.bonc.entity.ResultObject;
import com.example.bonc.entity.Split;

import java.io.IOException;


/**
 * @author TanHao
 * @date 2022/10/19 0019
 */
public interface CheckExcelService {

    /**
     * 校验考勤表中的人员是否在周报表中
     * @param check 文件名和路径对象
     * @return 返回接口调用结果
     */
    ResultObject checkout(Check check);

    /**
     * 拆分服务云文件
     * @param split 拆分信息
     * @return 拆分结果
     */
    ResultObject split(Split split);
}
