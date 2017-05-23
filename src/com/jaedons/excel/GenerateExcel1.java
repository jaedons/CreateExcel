package com.jaedons.excel;

import java.io.File;
import java.io.IOException;

import jxl.CellView;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class GenerateExcel1 {
	/**利用第三方jar包:jxl.jar 创建Excel*/
	public static void createExcel(){  
        try{  
            //打开文件  
            WritableWorkbook workbook = Workbook.createWorkbook(new File("D:\\test.xls"));  
            //生成名为“第一页”的工作表，参数0表示这是第一页   
            WritableSheet sheet = workbook.createSheet("第一页", 0);  
            //在Label对象的构造子中指名单元格位置是第一列第一行(0,0)   
            //以及单元格内容为test   
            Label label = new Label(0,0,"test");  
            //将定义好的单元格添加到工作表中   
            sheet.addCell(label);  
            /*生成一个保存数字的单元格 　　 
             * 必须使用Number的完整包路径，否则有语法歧义 　　 
             * 单元格位置是第二列，第一行，值为789.123*/   
            jxl.write.Number number = new jxl.write.Number(1,0,756);  
              
            sheet.addCell(number);  
              
            
//            sheet.insertColumn(1);//  在第一列插入空列  
              
            workbook.copySheet(0, "第二页", 1);  
             // 复制第一页的数据到第二页
            WritableSheet sheet2 = workbook.getSheet(1);  
            Range range = sheet2.mergeCells(0, 0, 0, 8);  
            sheet2.unmergeCells(range);  
              
            sheet2.addImage(new WritableImage(5, 5, 10, 20, new File("D:\\qr.png")));  
              
              
            CellView cv = new CellView();  
              
            WritableCellFormat cf = new WritableCellFormat();  
            cf.setBackground(Colour.BLUE);  
              
            cv.setFormat(cf);  
            cv.setSize(6000);  
            cv.setDimension(10);  
              
            sheet2.setColumnView(2, cv);  
              
            workbook.write();  
            workbook.close();  
              
        }catch(Exception e){}  
    }
	/**读取Excel*/
	public static void displayExcel(){  
        try {  
            Workbook wb = Workbook.getWorkbook(new File("D:\\test.xls"));  
            Sheet s = wb.getSheet(0);  
            System.out.println(s.getCell(0, 0).getContents());  
        } catch (BiffException e) {  
            // TODO Auto-generated catch block  
            e.printStackTrace();  
        } catch (IOException e) {  
            // TODO Auto-generated catch block  
            e.printStackTrace();  
        }  
          
    }  
}
