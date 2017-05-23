package com.jaedons.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

public class GenerateExcel2 {
	public static void exportExcel(){  
        
        HSSFWorkbook wb = new HSSFWorkbook();//����������  
          
        HSSFFont font = wb.createFont();  
        font.setFontHeightInPoints((short)24);  
        font.setFontName("����");  
        font.setColor(HSSFColor.BLACK.index);  
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
          
        HSSFCellStyle style = wb.createCellStyle();  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);  
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style.setBorderBottom(HSSFCellStyle.BORDER_THICK);  
        style.setFont(font);  
          
        HSSFSheet sheet = wb.createSheet("test");//��������������Ϊtest  
          
        int iRow = 0;//�к�  
        int iMaxCol = 17;//�������  
        HSSFRow row = sheet.createRow(iRow);  
        HSSFCell cell = row.createCell((short)0);  
        cell.setCellValue(new HSSFRichTextString("����excel"));  
        cell.setCellStyle(style);  
        sheet.addMergedRegion(new Region(iRow,(short)0,iRow,(short)(iMaxCol-1)));  
          
        ByteArrayOutputStream os = new ByteArrayOutputStream();  
          
        try{  
            wb.write(os);  
        }catch(IOException e){  
            e.printStackTrace();  
            //return null;  
        }  
          
        byte[] xls = os.toByteArray();  
          
        File file = new File("D:\\test01.xls");  
        OutputStream out = null;  
        try {  
             out = new FileOutputStream(file);  
             try {  
                out.write(xls);  
            } catch (IOException e) {  
                // TODO Auto-generated catch block  
                e.printStackTrace();  
            }  
        } catch (FileNotFoundException e1) {  
            // TODO Auto-generated catch block  
            e1.printStackTrace();  
        }  
          
    }  

}
