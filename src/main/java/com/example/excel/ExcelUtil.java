package com.example.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.util.StringUtils;

import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtil {

    public static void main(String[] args) throws Exception{
        String[] titleArray = {"姓名","性别","年龄"};
        List<People> list = new ArrayList<>();
        list.add(new People("测试1","男","18"));
        //添加导出数据 并获取poi对象
        HSSFWorkbook export = export(list, titleArray);
        //继续添加导出数据
        List<People> appendList = new ArrayList<>();
        appendList.add(new People("测试2","女","17"));
        appendExportData(workbook,list);
        //输出文件 //可根据业务需求调整
        //1.直接输出文件
        export.write(new File("测试.xlsx"));
        //2.直接设置在HttpResponse中读取 请求接口自动下载
        //outputSetting()
        //3.如果有oss可以直接将文件传到oss 返回文件url
    }


    private static final HSSFCellStyle style;
    private static final HSSFCellStyle dataStyle;
    private static final HSSFWorkbook workbook = new HSSFWorkbook();
    static {
        style = workbook.createCellStyle();//设置表头的类型
        style.setAlignment(HorizontalAlignment.CENTER);

        dataStyle = workbook.createCellStyle(); //设置数据类型
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    /**
     * 添加标题并输出其他行
     * @param dataList 任意类型集合
     * @param titleArray 标题数组
     * @param <T>
     * @return
     */
    public static<T> HSSFWorkbook export(List<T> dataList, String[] titleArray){
        HSSFFont font = workbook.createFont(); //设置字体
        HSSFSheet sheet = workbook.createSheet("sheet"); //创建一个sheet
        HSSFRow row;
        HSSFCell cell;
        try {
            //根据是否取出数据，设置header信息
            if(dataList.size() < 1){
                return null;
            }else{
                if(titleArray != null){
                    int lastRowNum = sheet.getLastRowNum() + 1;
                    row = sheet.createRow(lastRowNum);
                    row.setHeight((short)400);
                    for(int k = 0;k < titleArray.length; k++){
                        cell = row.createCell((short) k);//创建第0行第k列
                        cell.setCellValue(titleArray[k]);//设置第0行第k列的值
                        sheet.setColumnWidth((short)k,(short)8000);//设置列的宽度
                        font.setColor(HSSFFont.COLOR_NORMAL); // 设置单元格字体的颜色.
                        font.setFontHeight((short)350); //设置单元字体高度
                        dataStyle.setFont(font);//设置字体风格
                        cell.setCellStyle(dataStyle);
                    }
                }
                appendExportData(workbook,dataList);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbook;
    }

    /**
     * 输出其他行 无标题
     * @param workbook
     * @param dataList
     * @param <T>
     * @return
     */
    public static<T> HSSFWorkbook appendExportData(HSSFWorkbook workbook, List<T> dataList){
        try {
            HSSFCellStyle cellStyle = workbook.getCellStyleAt(21);
            //根据是否取出数据，设置header信息
            if(dataList.size() < 1){
                return workbook;
            }else{
                HSSFSheet sheet = workbook.getSheet("sheet");
                HSSFRow row;
                HSSFCell cell;
                int lastRowNum = sheet.getLastRowNum();
                // 给Excel填充数据
                Map<String, Integer> sortMap = new HashMap<>();
                Field[] listField = dataList.get(0).getClass().getDeclaredFields();
                for (int i = 0; i < listField.length; i++) {
                    sortMap.put("get" + StringUtils.capitalize(listField[i].getName()),i);
                }
                for(int i = 0; i < dataList.size(); i++){
                    Object obj = dataList.get(i);
                    Class<?> clazz  = obj.getClass();
                    row = sheet.createRow((short) (lastRowNum + i + 1));//创建第i+1行
                    row.setHeight((short)400);//设置行高

                    Method[] listMethod = clazz.getDeclaredMethods();
                    for (Method method : listMethod) {
                        String name = method.getName();
                        Integer column = sortMap.get(name);
                        if(null != column){
                            Class<?> returnType = method.getReturnType();
                            String returnTypeName = returnType.getName();
                            Object invoke = method.invoke(obj);
                            cell = row.createCell(column);//创建第i+1行第0列
                            if(returnTypeName.equals("java.util.Date")){
                                Date date = (Date) invoke;
                                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                                String str = sdf.format(date);
                                cell.setCellValue(str);
                            }else {
                                cell.setCellValue((String) invoke);//设置第i+1行第0列的值
                            }
                            cell.setCellStyle(cellStyle);//设置风格
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbook;
    }


    /**
     * 简单设置表头格式 可根据需求更新
     * @param alignment
     */
    public static void setStyle(HorizontalAlignment alignment){
        style.setAlignment(alignment);
    }

    /**
     * 简单设置数据格式  可根据需求更新
     * @param alignment
     */
    public static void setDataStyle(HorizontalAlignment alignment){
        dataStyle.setAlignment(alignment);
    }


//    public void outputSetting(String fileName, HttpServletResponse response, HSSFWorkbook workbook) {
//        OutputStream out = null;//创建一个输出流对象
//        try {
//            out = response.getOutputStream();// 得到输出流
//            response.setHeader("Content-disposition","attachment; filename=" + new String(fileName.getBytes(),"ISO-8859-1"));//filename是下载的xls的名
//            response.setContentType("application/msexcel;charset=UTF-8");//设置类型
//            response.setHeader("Pragma","No-cache");//设置头
//            response.setHeader("Cache-Control","no-cache");//设置头
//            response.setDateHeader("Expires", 0);//设置日期头
//            workbook.write(out);
//            out.flush();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }finally{
//            try{
//                if(out!=null){
//                    out.close();
//                }
//            }catch(IOException e){
//                e.printStackTrace();
//            }
//        }
//    }
}
