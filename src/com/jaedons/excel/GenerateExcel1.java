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
	/**���õ�����jar��:jxl.jar ����Excel*/
	public static void createExcel(){  
        try{  
            //���ļ�  
            WritableWorkbook workbook = Workbook.createWorkbook(new File("D:\\test.xls"));  
            //������Ϊ����һҳ���Ĺ���������0��ʾ���ǵ�һҳ   
            WritableSheet sheet = workbook.createSheet("��һҳ", 0);  
            //��Label����Ĺ�������ָ����Ԫ��λ���ǵ�һ�е�һ��(0,0)   
            //�Լ���Ԫ������Ϊtest   
            Label label = new Label(0,0,"test");  
            //������õĵ�Ԫ����ӵ���������   
            sheet.addCell(label);  
            /*����һ���������ֵĵ�Ԫ�� ���� 
             * ����ʹ��Number��������·�����������﷨���� ���� 
             * ��Ԫ��λ���ǵڶ��У���һ�У�ֵΪ789.123*/   
            jxl.write.Number number = new jxl.write.Number(1,0,756);  
              
            sheet.addCell(number);  
              
            
//            sheet.insertColumn(1);//  �ڵ�һ�в������  
              
            workbook.copySheet(0, "�ڶ�ҳ", 1);  
             // ���Ƶ�һҳ�����ݵ��ڶ�ҳ
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
	/**��ȡExcel*/
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
