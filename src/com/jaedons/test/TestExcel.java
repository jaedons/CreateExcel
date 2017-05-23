package com.jaedons.test;

import org.junit.Test;

import com.jaedons.excel.GenerateExcel1;
import com.jaedons.excel.GenerateExcel2;

public class TestExcel {
	@Test
	public void test1(){
		GenerateExcel1.createExcel();
		GenerateExcel1.displayExcel();
	}
	@Test
	public void test2(){
		GenerateExcel2.exportExcel();
	}
}
