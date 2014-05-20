package br.com.starcode.poisecurity;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import br.com.starcode.poisecurity.decryptor.WorksheetDecryptor;
import br.com.starcode.poisecurity.encryptor.WorksheetEncryptor;

public class EncryptXlsTests {

    /*@Test(expected=UnsupportedOperationException.class)
    public void testXLS() throws IOException {
	HSSFWorkbook workbook = (HSSFWorkbook) WorksheetDecryptor.decrypt(
		getClass().getResourceAsStream("Book1.xls"),
		DocumentType.LEGACY, 
		"12345");
	
	workbook.getSheetAt(0).getRow(0).getCell(0).setCellValue("Alterado!");
	
	//workbook.writeProtectWorkbook("54321", "54321");
	ByteArrayOutputStream bos = new ByteArrayOutputStream();
	workbook.write(bos);
	ByteArrayInputStream bis = new ByteArrayInputStream(bos.toByteArray());
	WorksheetEncryptor.encrypt(bis, new FileOutputStream("target/BookOut.xls"), DocumentType.LEGACY, "54321");
	
    }*/
    
    @Test
    public void createXLSModifyProtected() throws IOException {
	System.out.println("createXLSModifyProtected");
	
	//creates sheet
	HSSFWorkbook workbook = new HSSFWorkbook();
	HSSFSheet sheet = workbook.createSheet();
	HSSFRow row = sheet.createRow(0);
	HSSFCell cell = row.createCell(0);
	cell.setCellValue("Gravado");
	//workbook.writeProtectWorkbook("54321", "user");
	sheet.protectSheet("54321");
	
	//saves sheet
	new File("target/.output-file/xls").mkdirs();
	workbook.write(new FileOutputStream("target/.output-file/xls/ModifyProtected.xls"));
	
	//read again and check
	HSSFWorkbook workbook2 = new HSSFWorkbook(new FileInputStream("target/.output-file/xls/ModifyProtected.xls"));
	Assert.assertEquals("Gravado", workbook2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
    }
    
    @Test(expected = UnsupportedOperationException.class)
    public void createXLSOpenProtected() throws IOException {
	System.out.println("createXLSOpenProtected");
	
	//creates sheet
	HSSFWorkbook workbook = new HSSFWorkbook();
	HSSFSheet sheet = workbook.createSheet();
	HSSFRow row = sheet.createRow(0);
	HSSFCell cell = row.createCell(0);
	cell.setCellValue("Gravado");
	
	//saves sheet
	ByteArrayOutputStream bos = new ByteArrayOutputStream();
	workbook.write(bos);
	bos.close();
	ByteArrayInputStream bis = new ByteArrayInputStream(bos.toByteArray());
	
	new File("target/.output-file/xls").mkdirs();
	WorksheetEncryptor.encrypt(
		bis, 
		new FileOutputStream("target/.output-file/xls/OpenProtected.xls"), 
		DocumentType.LEGACY, 
		"54321");
	bis.close();
	
	//read again and check
	HSSFWorkbook workbook2 = (HSSFWorkbook) WorksheetDecryptor.decrypt(
		new FileInputStream("target/.output-file/xls/OpenProtected.xls"),
		DocumentType.LEGACY, 
		"54321");
	Assert.assertEquals("Gravado", workbook2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
    }

    @Test(expected = UnsupportedOperationException.class)
    public void createXLSOpenAndModifyProtected() throws IOException {
	System.out.println("createXLSOpenAndModifyProtected");
	
	//creates sheet
	HSSFWorkbook workbook = new HSSFWorkbook();
	HSSFSheet sheet = workbook.createSheet();
	sheet.protectSheet("12345");
	
	HSSFRow row = sheet.createRow(0);
	HSSFCell cell = row.createCell(0);
	cell.setCellValue("Gravado");
	
	//saves sheet
	ByteArrayOutputStream bos = new ByteArrayOutputStream();
	workbook.write(bos);
	bos.close();
	ByteArrayInputStream bis = new ByteArrayInputStream(bos.toByteArray());
	
	new File("target/.output-file/xls").mkdirs();
	WorksheetEncryptor.encrypt(
		bis, 
		new FileOutputStream("target/.output-file/xls/OpenAndModifyProtected.xls"), 
		DocumentType.LEGACY, 
		"54321");
	bis.close();
	
	//read again and check
	HSSFWorkbook workbook2 = (HSSFWorkbook) WorksheetDecryptor.decrypt(
		new FileInputStream("target/.output-file/xls/OpenProtected.xls"),
		DocumentType.LEGACY, 
		"54321");
	Assert.assertEquals("Gravado", workbook2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
    }


}
