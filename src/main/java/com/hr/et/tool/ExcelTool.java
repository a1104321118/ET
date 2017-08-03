package com.hr.et.tool;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by hr on 2017/08/03.
 *
 */
public class ExcelTool {

    //读取文件
    private Workbook workbook = null;
    //写文件
    private WritableWorkbook writableWorkbook = null;


    /**
     * 初始化工具，需要传入文件路径
     * @param filePath
     * @throws IOException
     * @throws BiffException
     */
    public ExcelTool(String filePath){
        File file = new File(filePath);
        if(!file.isFile() || !file.exists()){
            try {
                throw new FileNotFoundException("文件不存在");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }
        try {
            this.workbook = Workbook.getWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取全部数据
     * @param clazz
     * @param <T>
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public <T> List<T> readData(Class<T> clazz){
        try {
            return readData(clazz, 1, 0);
        } catch (IllegalAccessException | InstantiationException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 读取数据，从 begin 开始，数量为num
     * @param clazz
     * @param begin
     * @param num
     * @param <T>
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public <T> List<T> readData(Class<T> clazz, int begin, int num) throws IllegalAccessException, InstantiationException {
        List<T> result = new ArrayList<>();

        Sheet sheet = this.workbook.getSheet(0);

        //获得所有列名
        List<String> cloumnNames = new ArrayList<>();
        int cloumns = sheet.getColumns();
        for (int i=0; i<cloumns; i++){
            cloumnNames.add(sheet.getCell(i, 0).getContents());
        }

        //获得该类的属性
        Field[] fields = clazz.getDeclaredFields();

        int rows = sheet.getRows();
        begin = begin <= 0 ? begin+1 : begin;
        int end = (begin + num) > rows ? rows : (begin + num);
        if(num == 0){
            end = rows;
        }

        for (int i=begin; i<end; i++){//对每一行数据进行处理
            Cell[] row = sheet.getRow(i);//获得这一行的数据
            T t = clazz.newInstance(); //通过反射实例化泛型对象
            for (int j=0; j<row.length; j++){//对这一行的数据进行一个一个处理
                String thisCloumnName = cloumnNames.get(j);//先获得列名
                for (Field field : fields){
                    if(thisCloumnName.equals(field.getName())){//当列名和这个属性的名称一样时，认为他们相互映射
                        field.setAccessible(true);
                        field.set(t, row[j].getContents());//设置属性值
                    }
                }
            }
            result.add(t);
        }

        workbook.close();

        return result;
    }

    /**
     * 写数据
     * @param data
     * @param <T>
     */
    public <T> void writeData(List<T> data, String outFilePath) {

        //判断数据有效性
        if(null == data || data.size() == 0){
            return;
        }

        if(null == outFilePath || outFilePath.equals("")){
            return;
        }

        WritableSheet sheet = null;
        File file = new File(outFilePath);
        try {
            this.writableWorkbook = Workbook.createWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        sheet = this.writableWorkbook.createSheet("result", 0);

        //写入标题,顺序是fields的顺序
        T t = data.get(0);
        Field[] fields = t.getClass().getDeclaredFields();
        Label[] head = new Label[fields.length];
        for (int i=0; i<fields.length; i++){
            head[i] = new Label(i, 0, fields[i].getName());
        }
        writeLine(sheet, head);//先把Person的属性都写入excel的第一行

        //写入数据
        try {
            for (int i = 0; i < data.size(); i++) {//开始写list里面的数据
                T result = data.get(i);
                for (int j = 0; j < fields.length; j++) {//逐个写属性

                    Field field = result.getClass().getDeclaredField(fields[j].getName());
                    field.setAccessible(true);
                    String s = (String)field.get(result);
                    sheet.addCell(new Label(j, i+1, s));

                }
            }
        }catch (NoSuchFieldException | WriteException | IllegalAccessException e) {
            e.printStackTrace();
        }

        try {
            this.writableWorkbook.write();
            this.writableWorkbook.close();
        } catch (WriteException | IOException e) {
            e.printStackTrace();
        }


    }

    private void writeLine(WritableSheet sheet, Label[] data){

        try {
            for (int i = 0; i < data.length; i++) {
                sheet.addCell(data[i]);
            }
        } catch (WriteException e) {
            e.printStackTrace();
        }

    }

}
