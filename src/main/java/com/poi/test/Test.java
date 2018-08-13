package com.poi.test;

import com.poi.model.Student;
import com.poi.utils.ExcelUtil;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Created with IntelliJ IDEA.
 * Description:
 * User: weicaijia
 * Date: 2018/8/13 11:25
 * Time: 14:15
 */
public class Test {



    public static void main(String[] args) {
        String sheetName = "学生";
        String[] titles = new String[]{"Id","Name","Sex","Birth","MathGrade","Visible"};
        String[] fieldNames = new String[]{"id","name","sex","birth","mathGrade","visible"};
        List<Student> studentList = new ArrayList<>();
        Student student1 = new Student();
        student1.setId(1111111111111111111L);
        student1.setName("用了金坷垃");
        student1.setSex((byte) 1);
        student1.setBirth(new Date());
        student1.setMathGrade(88.5);
        student1.setVisible(true);

        Student student2 = new Student();
        student2.setId(222222222222222222L);
        student2.setName("亩产一千八");
        student2.setSex((byte) 3);
        student2.setBirth(new Date());
        student2.setMathGrade(10.0);
        student2.setVisible(false);

        studentList.add(student1);
        studentList.add(student2);

        String fileNamePath = "D:/test.xls";

        try {

            //导出excel文件
            ExcelUtil.defaultExport(fileNamePath,sheetName,studentList,titles,fieldNames);

            //把Excel导入到程序中
            List<Map<String, Object>> impList;
            try {
                impList = ExcelUtil.defaultImport(fileNamePath, fieldNames);

                for (Map<String, Object> map : impList) {
                    System.out.println(map.get("id"));
                    System.out.println(map.get("name"));
                    System.out.println(map.get("birth"));
                    System.out.println(map.get("sex"));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }


            //插入一条数据
            //ExcelUtil.insertRows(fileNamePath,0,sheetName,"第一行，测试一下",fieldNames.length-1);

            //Sheet sheet = ExcelUtil.getSheet(fileNamePath,sheetName);
            short a =  4*256;
            ExcelUtil.insertRows(fileNamePath,0,sheetName,"第一行，不知道应该说点啥\r\n12345444\r\n少时诵诗书",fieldNames.length-1,a);
            // ExcelUtil.insertRows(fileNamePath,0,sheetName,"第二行，哈哈哈哈，这是一个背叛",fieldNames.length-1);
            // ExcelUtil.insertRows(fileNamePath,0,sheetName,"第三行，灰发肥会挥发",fieldNames.length-1);
            // ExcelUtil.insertRows(fileNamePath,0,sheetName,"第四行，黑化肥会发灰",fieldNames.length-1);

        } catch (Exception e) {
            e.printStackTrace();
        }


    }


}
