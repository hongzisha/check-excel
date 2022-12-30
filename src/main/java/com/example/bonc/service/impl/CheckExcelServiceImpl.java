package com.example.bonc.service.impl;

import com.example.bonc.entity.Check;
import com.example.bonc.entity.ResultObject;
import com.example.bonc.entity.Split;
import com.example.bonc.enumUtils.EnumConstant;
import com.example.bonc.enumUtils.ResponseConstant;
import com.example.bonc.service.CheckExcelService;
import com.example.bonc.utils.OfficeUtils;
import com.sun.org.apache.xpath.internal.operations.Variable;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.DateUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author TanHao
 * @date 2022/10/19 0019
 */
@Slf4j
@Service
public class CheckExcelServiceImpl implements CheckExcelService {

    @Value("${filePath.path}")
    private String path;

    /**
     * 读取并处理excel数据
     *
     * @param check 文件名和路径类
     * @return 接口调用结果
     */
    @Override
    public ResultObject checkout(Check check) {
        try {
            List<Map<Object, Map<Integer, Object>>> attendanceList = OfficeUtils.creatWorkBook(check.getPath() + File.separator + check.getAttendanceSheetFilename(), Integer.parseInt(EnumConstant.ATTENDANCE.getCode()));
            List<Map<Object, Map<Integer, Object>>> weekList = OfficeUtils.creatWorkBook(check.getPath() + File.separator + check.getWeekFilename(), Integer.parseInt(EnumConstant.WEEK.getCode()));
            // 判断两个excel是否有数据
            if (attendanceList != null && weekList != null) {
                // 清洗数据
                Map<String, List<String>> weekMap = weekDataCleaning(weekList);
                Map<String, List<String>> attMap = attDataCleaning(attendanceList, weekMap);
                // 比较两者的异处
                Map<String, List<String>> checkResult = check(weekMap, attMap);
                Map<String, Map<String, List<String>>> map = new HashMap<>(8);
                map.put("以下是没有周报或者考勤数据的员工姓名", checkResult);
                log.info("以下是没有周报或者考勤数据的员工姓名：{}", checkResult);
                return new ResultObject<>(ResponseConstant.SUCCESS.getCode(), ResponseConstant.SUCCESS.getMsg(), map);
            } else {
                // 两个excel无数据则返回失败
                log.info("{}", "比较失败，请检查文件是否有内容格式错误，路径、文件名不正确等问题");
                return new ResultObject<>(ResponseConstant.FAILURE.getCode(), ResponseConstant.FAILURE.getMsg(), "比较失败，请检查文件是否有内容格式错误，路径、文件名不正确等问题");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return new ResultObject<>(ResponseConstant.FAILURE.getCode(), ResponseConstant.FAILURE.getMsg(), "比较文件失败");
    }

    /**
     * 清洗周报数据
     *
     * @param weekList 周报数据集合
     * @return 周报数据
     */
    private Map<String, List<String>> weekDataCleaning(List<Map<Object, Map<Integer, Object>>> weekList) {
        // 清洗周报数据
        List<String> weeks = new ArrayList<>();
        Map<String, List<String>> map = new HashMap<>(8);
        if (weekList != null) {
            weekList.get(0).forEach((key, value) -> weeks.add(String.valueOf(key)));
            for (String week : weeks) {
                List<String> finalList = new ArrayList<>();
                for (Map<Object, Map<Integer, Object>> objectMapMap : weekList) {
                    objectMapMap.forEach((key1, value1) -> {
                        if (week == key1) {
                            value1.forEach((key, value) -> {
                                if (value != null) {
                                    finalList.add(String.valueOf(value));
                                }
                            });
                        }
                    });
                }
                // 去除list中的重复值
                ArrayList<String> collect = finalList.stream().collect(Collectors.collectingAndThen(
                        Collectors.toCollection(() -> new TreeSet<>(
                                Comparator.comparing(
                                        String::valueOf))), ArrayList::new));
                map.put(week, collect);
            }
        }
        return map;
    }


    /**
     * 清洗考勤数据
     *
     * @param attendanceList 考勤数据集合
     * @param weekMap        周报数据对象
     * @return 考勤数据
     */
    private Map<String, List<String>> attDataCleaning(List<Map<Object, Map<Integer, Object>>> attendanceList, Map<String, List<String>> weekMap) {
        Calendar calendar = new GregorianCalendar(1900, Calendar.JANUARY, -1);
        // 清洗考勤数据
        Map<String, List<Object>> map = new HashMap<>(8);
        if (attendanceList != null && weekMap != null) {
            attendanceList.get(0).forEach((key, value) -> {
                List<Object> time = new ArrayList<>();
                value.forEach((key1, value1) -> {
                    if (value1 != null) {
                        long date = DateUtils.addDays(calendar.getTime(), Integer.parseInt(String.valueOf(value1))).getTime();
                        time.add(date);
                    }
                });
                map.put(String.valueOf(key), time);
            });

            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy/MM/dd");

            // 清洗考勤数据
            Map<String, List<String>> att = new HashMap<>(8);
            weekMap.forEach((k, v) -> {
                try {
                    long startTime = simpleDateFormat.parse(String.valueOf(k).split("-")[0]).getTime();
                    long endTime = simpleDateFormat.parse(String.valueOf(k).split("-")[1]).getTime();
                    List<String> attName = new ArrayList<>();
                    map.forEach((k1, v1) -> {
                        if (k1 != null) {
                            v1.forEach(z -> {
                                long time = (long) z;
                                if (startTime <= time && endTime >= time) {
                                    attName.add(k1);
                                }
                            });
                        }
                    });
                    ArrayList<String> collect = attName.stream().collect(Collectors.collectingAndThen(
                            Collectors.toCollection(() -> new TreeSet<>(
                                    Comparator.comparing(
                                            String::valueOf))), ArrayList::new));
                    att.put(k, collect);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
            return att;
        }
        return null;
    }

    /**
     * 比较两个excel的异处
     *
     * @param weekMap 周报数据
     * @param attMap  考勤数据
     * @return 没有写周报的员工
     */
    private Map<String, List<String>> check(Map<String, List<String>> weekMap, Map<String, List<String>> attMap) {
        if (weekMap != null && attMap != null) {
            Map<String, List<String>> map = new HashMap<>(8);
            weekMap.forEach((k, v) -> {
                List<String> list2 = attMap.get(k);
                List<String> list = checkList(v, list2);
                map.put("\n" + k, list);
            });

            return map;
        }
        return null;
    }

    /**
     * 比较两个list集合的差异
     * 1、先区分两个list的大小
     * 2、给最大的list打上标号1
     * 3、再将最大的list的值放入最小的list中去获取得到的数据标号2
     * 4、此时map中的数据被分为了两类，之后只需要找出为1的即为不同的数据
     *
     * @param list1 集合A
     * @param list2 集合B
     * @return 两个集合的不同数据
     */
    private List<String> checkList(List<String> list1, List<String> list2) {


        List<String> diff = new ArrayList<>();
        List<String> maxList = list1;
        List<String> minList = list2;
        if (list2.size() > list1.size()) {
            maxList = list2;
            minList = list1;
        }
        Map<String, Integer> map = new HashMap<>(maxList.size());
        for (String string : maxList) {
            map.put(string, 1);
        }
        for (String string : minList) {
            if (map.get(string) != null) {
                map.put(string, 2);
                continue;
            }
            diff.add(string);
        }
        for (Map.Entry<String, Integer> entry : map.entrySet()) {
            if (entry.getValue() == 1) {
                diff.add(entry.getKey());
            }
        }
        return diff;
    }


    /**
     * 按行分割文件
     *
     * @param split 文件名和个数
     * @return 返回拆分后的文件名
     */
    @Override
    public ResultObject split(Split split) {
        String sourcePath = path + File.separator + split.getFileName();
        // 判断分割文件个数是否超过限制
        if (split.getFileNumber() > 5) {
            return new ResultObject<>("0001", "失败", "拆分文件失败，请确认拆分个数是否在2-5个内");
        }
        // 源文件名
        String sourceFileName = sourcePath.substring(sourcePath.lastIndexOf(File.separator) + 1, sourcePath.lastIndexOf("."));
        // 切割后的文件名
        String splitFileName = path + File.separator + sourceFileName + "_%s.csv";
        if (!new File(sourcePath).exists()) {
            return new ResultObject<>("0001", "失败", "拆分文件失败，请检查该文件是否在/nfs_shared/fine-operation/upload/crowd/目录下存在");
        }
        File targetDirectory = new File(path);
        // 判断文件目录是否存在
        if (!targetDirectory.exists()) {
            targetDirectory.mkdirs();
        }

        List<Map<String, Integer>> resultFileNames = new ArrayList<>();
        // 字符输出流
        PrintWriter pw = null;
        String tempLine;
        // 本次行数累计,达到rows开辟新文件
        int lineNum = 0;
        // 当前文件索引
        int splitFileIndex = 1;
        // 文件行数
        long lines;
        // 文件个数标识
        int flag = 1;

        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(sourcePath)))) {
            // 文件总行数
            lines = Files.lines(Paths.get(sourcePath)).count();
            // 新文件行数
            int rows = (int) Math.ceil(1.0 * lines / split.getFileNumber());
            // 尾文件行数与新文件行数之差
            long num = (long) split.getFileNumber() * rows - lines;
            log.info("原文件行数:{}, 新文件行数:{}", lines, rows);

            // 第一个文件名
            String parsedFileName = String.format(splitFileName, splitFileIndex);
            pw = new PrintWriter(parsedFileName);
            // 记录文件名和行数
            Map<String, Integer> map = new HashMap<>(8);
            map.put(parsedFileName, rows);
            resultFileNames.add(map);
            while ((tempLine = br.readLine()) != null) {
                // 需要换新文件
                if (lineNum > 0 && lineNum % rows == 0) {
                    pw.flush();
                    pw.close();
                    // 文件名
                    parsedFileName = String.format(splitFileName, ++splitFileIndex);
                    pw = new PrintWriter(parsedFileName);
                    // 记录文件名和行数
                    map = new HashMap<>(8);
                    // 判断文件个数是否为最后一个
                    map.put(parsedFileName, (int) (++flag == split.getFileNumber() ? num == 0 ? rows : rows - num : rows));
                    resultFileNames.add(map);
                }
                pw.write(tempLine + "\n");
                lineNum++;
            }
            log.info("返回文件列表：{}", resultFileNames);
        } catch (Exception e) {
            log.error("拆分文件异常: {}", e.toString());
        } finally {
            if (null != pw) {
                pw.flush();
                pw.close();
            }
        }
        if (resultFileNames.size() > 0) {
            return new ResultObject<>("0000", "成功", String.valueOf(resultFileNames));
        }
        return new ResultObject<>("0001", "失败", "拆分文件异常");
    }
}
