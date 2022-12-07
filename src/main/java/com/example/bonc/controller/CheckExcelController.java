package com.example.bonc.controller;

import com.example.bonc.entity.Check;
import com.example.bonc.entity.ResultObject;
import com.example.bonc.entity.Split;
import com.example.bonc.service.CheckExcelService;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;

/**
 * @author TanHao
 * @date 2022/10/19 0019
 */
@RestController
@RequestMapping("/bonc")
public class CheckExcelController {

    @Resource
    private CheckExcelService checkExcelService;

    @PostMapping("/check")
    public ResultObject check (@RequestBody Check check) {
        return checkExcelService.checkout(check);
    }
    @PostMapping("/shell")
    public ResultObject shell (@RequestBody Split split){
        return checkExcelService.shell(split);
    }
}
