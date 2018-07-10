package Utils;

import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class ListToExcel {

    //每个excel中sheet的容量
    private static final int  sheetContent=10000;
    //每个excel中的容量
    private static final int  excelBig=50000;

    public static void main(String[] args) throws Exception {
        String[] enFields={"name","age"};
        String[] cnFields={"名字","年龄"};
        Student student1=new Student();
        student1.setName("李红");
        student1.setAge(12);
        Student student2=new Student();
        student2.setName("张三");
        student2.setAge(21);
        Student student3=new Student();
        student3.setName("张三");
        student3.setAge(21);
        List<Student> list=new ArrayList<Student>();
        list.add(student1);
        list.add(student2);
        list.add(student3);
        list.add(student1);
        list.add(student2);
        list.add(student3);
        list.add(student1);
        list.add(student2);
        list.add(student3);
        list.add(student1);
        list.add(student2);
        list.add(student3);
        String excelname="test";
        exportExcel(excelname,list,enFields,cnFields);
    }
    /**
     * 导出Excel
     * @param excelName   要导出的excel名称
     * @param list   要导出的数据集合
     * @param enFields 英文字段
     * @param cnFields 中文字段,即要导出的excel表头
     * @return
     */
    public static <T> void exportExcel(String excelName, List<T> list, String[] enFields, String[] cnFields) throws Exception {
        // 设置默认文件名为当前时间：年月日时分秒
        if (excelName==null || excelName=="") {
            excelName = new SimpleDateFormat("yyyyMMddhhmmss").format(
                    new Date()).toString();
        }

        int total=list.size();  //总记录数

        //判断需要分成多少个excel
        int excelCount=(total%excelBig==0)?(total/excelBig):(total/excelBig+1);

        //定义一个excel集合
        List<HSSFWorkbook>  excelList=new ArrayList<HSSFWorkbook>();

        HSSFWorkbook wb=null;  //excel工作表
        int begin = 0;   //数据源的开始位置
        int end = 0;      //数据源的结束位置

        //循环得到excel集合
        for(int i=0;i<excelCount;i++){
            //第一次begin从零开始,否则从结束位置加1开始
            if(end!=0){
                begin=end;
            }
            end=begin+excelBig;
            //判断如果最后的位置大于总记录数，则修改为总记录数
            if(end>list.size()){
                end=list.size();
            }
            //生成excel
            wb=createExcel(excelName+(i+1), list,begin,end, enFields, cnFields);
            excelList.add(wb);
        }
        for(HSSFWorkbook workbook:excelList){
            FileOutputStream fos=new FileOutputStream(new File("/Users/yunjian/Downloads/test.xls"));
            workbook.write(fos);
        }


        /*OutputStream output = null;
        try {
            String setHeader = setHeader(request, excelName);
            // 输出Excel文件
            output = response.getOutputStream();
            response.reset();
            response.setHeader("Content-disposition", setHeader + ".xls");
            response.setContentType("application/msexcel");

            for(HSSFWorkbook workbook:excelList){
                workbook.write(output);
            }


        } catch (Exception e) {
            logger.error(e.getMessage(), e);
        } finally {
            if(output != null){
                try {
                    output.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(), e);
                }
            }
        }*/
    }
        /**
         * 生成Excel
         * @param excelName   要导出的excel名称
         * @param list       要导出的数据集合
         * @param begin   数据集合的开始位置
         * @param end      数据集合的结束位置
         * @param enFields 英文字段对应数组
         * @param cnFields 中文字段对应数组,即要导出的excel表头
         * @return
         */
        private static <T> HSSFWorkbook createExcel(String excelName,List<T> list,Integer begin,Integer end,
                String[] enFields,String[] cnFields){

            //计算sheet的数量
            int sheetCount=0;
            if ((end-begin)%sheetContent==0){
                sheetCount=(end-begin)/sheetContent;
            }else{
                sheetCount=(end-begin)/sheetContent+1;
            }

            //创建一个WorkBook,对应一个Excel文件
            HSSFWorkbook wb=new HSSFWorkbook();
            //创建单元格，并设置值表头 设置表头居中
            HSSFCellStyle style=wb.createCellStyle();
            //创建一个居中格式
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            //定义sheet
            HSSFSheet sheet=null;
            //定义sheet的名称
            String sheetName=excelName+"-";
            //sheet中的数据开始位置
            int beginSheet=0;
            //sheet中的数据结束位置
            int endSheet=0;

            //循环创建sheet
            for (int i=0;i<sheetCount;i++){
                //在Workbook中，创建一个sheet，对应Excel中的工作薄（sheet）
                sheet=wb.createSheet(sheetName+(i+1));

                //第一次beginSheet从begin开始,否则从结束位置加1开始
                if(endSheet==0){
                    beginSheet=begin;
                }else{
                    beginSheet=endSheet;
                }
                //sheet结束位置=sheet开始位置+每个sheet的容量
                endSheet=beginSheet+sheetContent;
                //判断如果最后的位置大于总记录数，则修改为总记录数
                if(endSheet>end){
                    endSheet=end;
                }
                //获取样式
                HSSFCellStyle bodyStyle = getTbodyStyle(wb);
                HSSFCellStyle tableNameSty = getTableNameStyle(wb);

//			sheet = wb.getSheet(sheetName+"1");
//			insertRow(wb,sheet,0,1,"预付卡消费日报",enFields.length);//插入行

                try {
                    // 填充工作表
                    fillSheet(sheet,list,beginSheet,endSheet,enFields,cnFields,bodyStyle,tableNameSty);


                } catch (Exception e) {

                }
            }

            return wb;
        }
    /**
     * 获得数据的样式
     * @return
     */
    public static HSSFCellStyle getTbodyStyle(HSSFWorkbook wb) {
        // 设置样式
        HSSFCellStyle style = wb.createCellStyle();
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        // 设置字体
        HSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short)10);
        font.setFontName("宋体");
        // 把字体应用到当前的样式
        style.setFont(font);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        return style;
    }
    /**
     * 设置表名的样式
     * @param wb
     * @return
     */
    public static HSSFCellStyle getTableNameStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 设置字体
        HSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("宋体");
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        style.setFont(font);
        return style;
    }
    /**
     * 向工作表中填充数据(导出大数据量表格时,建议传入的数据源的List 的泛型为Map,效率快)
     *
     * @param sheet
     *            excel的工作表名称
     * @param list
     *            数据源
     *@param beginSheet
     *            数据源开始位置
     *@param endSheet
     *            数据源结束位置
     * @param enFields
     *            英文字段名
     * @param cnFields
     * 			  中文字段名(表头)
     * @param bodyStyle
     * 			  表格中的格式
     * @throws Exception
     *             异常
     */
    public static <T> void fillSheet(HSSFSheet sheet, List<T> list,Integer beginSheet,Integer endSheet,
                                     String[] enFields,String[] cnFields,HSSFCellStyle bodyStyle, HSSFCellStyle tableNameStyle) throws Exception {
        //logger.info("向工作表中填充数据:fillSheet()");

        //在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        HSSFRow row=sheet.createRow((int)0);

        // 填充表头
        int i =0;
        HSSFCell cell=null;
        for (String field :  cnFields) {
            cell=row.createCell(i);
            cell.setCellValue(field);
            cell.setCellStyle(tableNameStyle);
            //sheet.autoSizeColumn(i);
            i++;
        }

        // 填充内容
        int rowInt=1;
        Object fieldValue=null;
        int place=0;
        T item=null;
        for (place=beginSheet;place<endSheet;place++) {
            row = sheet.createRow(rowInt);
            int j=0;
            //得到数据的一条记录
            item=list.get(place);
            if(item instanceof Map){
                //结果集为 Map 大数据量导出表格时 建议使用此方法  效率高
                Map<String,Object> values = (Map<String, Object>)item;
                for (String field : enFields) {
                    fieldValue = values.get(field);
                    HSSFCell bodyCell = row.createCell(j);
                    bodyCell.setCellStyle(bodyStyle);
                    if(fieldValue != null && !"".equals(fieldValue)){
                        bodyCell.setCellValue(fieldValue.toString());
                        j++;
                    } else {
                        bodyCell.setCellValue("");
                        j++;
                    }
                }
            } else {
                //结果集为一个 实体 Bean 通过反射未每个列赋值  大数据量时速度较慢  不建议使用
                for (String field : enFields) {
                    fieldValue = getFieldValueByNameSequence(field, item);
                    HSSFCell bodyCell = row.createCell(j);
                    bodyCell.setCellStyle(bodyStyle);
                    if(fieldValue != null && !"".equals(fieldValue)){
                        bodyCell.setCellValue(fieldValue.toString());
                        j++;
                    } else {
                        bodyCell.setCellValue("");
                        j++;
                    }
                }
            }
            rowInt++;
        }

        for (int index = 0 ; index < cnFields.length; index++) {
            String field = cnFields[index];
//			sheet.setColumnWidth(index,  field.length()*4*256);//设置列宽(自适应列头名字)
            sheet.setColumnWidth(index,  30*256);//设置列宽(自定义列宽)
        }
    }
    /**
     * 根据带路径或不带路径的属性名获取属性值,即接受简单属性名，
     * 如userName等，又接受带路径的属性名，如student.department.name等
     *
     * @param fieldNameSequence 带路径的属性名或简单属性名
     * @param o                 对象
     * @return                  属性值
     * @throws Exception        异常
     *
     */
    public static Object getFieldValueByNameSequence(String fieldNameSequence,
                                                     Object o) throws Exception {
        //logger.info("根据带路径或不带路径的属性名获取属性值,即接受简单属性名:getFieldValueByNameSequence()");
        Object value = null;

        // 将fieldNameSequence进行拆分
        String[] attributes = fieldNameSequence.split("\\.");
        if (attributes.length == 1) {
            value = getFieldValueByName(fieldNameSequence, o);
        } else {
            // 根据数组中第一个连接属性名获取连接属性对象，如student.department.name
            Object fieldObj = getFieldValueByName(attributes[0], o);
            //截取除第一个属性名之后的路径
            String subFieldNameSequence = fieldNameSequence
                    .substring(fieldNameSequence.indexOf(".") + 1);
            //递归得到最终的属性对象的值
            value = getFieldValueByNameSequence(subFieldNameSequence, fieldObj);
        }
        return value;
    }
    /**
     * 根据字段名获取字段值
     *
     * @param fieldName  字段名
     * @param o          对象
     * @return           字段值
     * @throws Exception 异常
     *
     */
    public static Object getFieldValueByName(String fieldName, Object o)
            throws Exception {

        //logger.info("根据字段名获取字段值:getFieldValueByName()");
        Object value = null;
        //根据字段名得到字段对象
        Field field = getFieldByName(fieldName, o.getClass());

        //如果该字段存在，则取出该字段的值
        if (field != null) {
            field.setAccessible(true);//类中的成员变量为private,在类外边使用属性值，故必须进行此操作
            value = field.get(o);//获取当前对象中当前Field的value
        } else {
            throw new Exception(o.getClass().getSimpleName() + "类不存在字段名 "
                    + fieldName);
        }
        return value;
    }
    /**
     * 根据字段名获取字段对象
     *
     * @param fieldName
     *            字段名
     * @param clazz
     *            包含该字段的类
     * @return 字段
     */
    public static Field getFieldByName(String fieldName, Class<?> clazz) {
        //logger.info("根据字段名获取字段对象:getFieldByName()");
        // 拿到本类的所有字段
        Field[] selfFields = clazz.getDeclaredFields();

        // 如果本类中存在该字段，则返回
        for (Field field : selfFields) {
            //如果本类中存在该字段，则返回
            if (field.getName().equals(fieldName)) {
                return field;
            }
        }

        // 否则，查看父类中是否存在此字段，如果有则返回
        Class<?> superClazz = clazz.getSuperclass();
        if (superClazz != null && superClazz != Object.class) {
            //递归
            return getFieldByName(fieldName, superClazz);
        }

        // 如果本类和父类都没有，则返回空
        return null;
    }
}
class Student{
    private String name;
    private int age;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }
}
