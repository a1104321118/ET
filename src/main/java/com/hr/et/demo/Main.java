package com.hr.et.demo;


import com.hr.et.tool.ExcelTool;

import java.util.List;

/**
 * Created by hr on 2017/08/03.
 */
public class Main {

    public static void main(String[] args) {
        //文件输入输出路径
        String filePath = "D:\\idea\\MyProjects\\ET\\src\\main\\java\\resources\\test.xls";
        String outFilePath = "D:\\idea\\MyProjects\\ET\\src\\main\\java\\resources\\test-result.xls";

        ExcelTool excelTool = new ExcelTool(filePath);

        //读数据
        List<Person> peopleList = excelTool.readData(Person.class);

        //修改数据
        for (Person person : peopleList){
            System.out.println(person);
            person.setState("SUCCESS");
        }

        //写数据
        excelTool.writeData(peopleList, outFilePath);

    }
}
