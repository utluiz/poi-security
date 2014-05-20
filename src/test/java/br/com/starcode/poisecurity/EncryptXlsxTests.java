package br.com.starcode.poisecurity;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import br.com.starcode.poisecurity.decryptor.WorksheetDecryptor;
import br.com.starcode.poisecurity.encryptor.WorksheetEncryptor;

public class EncryptXlsxTests {

    @Test
    public void createXLSXModifyProtected() throws IOException {
	System.out.println("createXLSXModifyProtected");
	
	//creates sheet
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet();
	XSSFRow row = sheet.createRow(0);
	XSSFCell cell = row.createCell(0);
	cell.setCellValue("Gravado");
	sheet.protectSheet("54321");
	
	//saves sheet
	new File("target/.output-file/xlsx").mkdirs();
	workbook.write(new FileOutputStream("target/.output-file/xlsx/ModifyProtected.xlsx"));
	
	//read again and check
	XSSFWorkbook workbook2 = new XSSFWorkbook(new FileInputStream("target/.output-file/xlsx/ModifyProtected.xlsx"));
	Assert.assertEquals("Gravado", workbook2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
    }
    
    @Test
    public void createXLSXOpenProtected() throws IOException {
	System.out.println("createXLSXOpenProtected");
	
	//creates sheet
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet();
	XSSFRow row = sheet.createRow(0);
	XSSFCell cell = row.createCell(0);
	cell.setCellValue("Gravado");
	
	//saves sheet
	ByteArrayOutputStream bos = new ByteArrayOutputStream();
	workbook.write(bos);
	bos.close();
	ByteArrayInputStream bis = new ByteArrayInputStream(bos.toByteArray());
	
	new File("target/.output-file/xlsx").mkdirs();
	WorksheetEncryptor.encrypt(
		bis, 
		new FileOutputStream("target/.output-file/xlsx/OpenProtected.xlsx"), 
		DocumentType.XML, 
		"54321");
	bis.close();
	
	//read again and check
	XSSFWorkbook workbook2 = (XSSFWorkbook) WorksheetDecryptor.decrypt(
		new FileInputStream("target/.output-file/xlsx/OpenProtected.xlsx"),
		DocumentType.XML, 
		"54321");
	Assert.assertEquals("Gravado", workbook2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
    }

    @Test
    public void createXLSXOpenAndModifyProtected() throws IOException {
	System.out.println("createXLSXOpenAndModifyProtected");
	
	//creates sheet
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet();
	sheet.protectSheet("54321");
	
	XSSFRow row = sheet.createRow(0);
	XSSFCell cell = row.createCell(0);
	cell.setCellValue("Gravado");
	
	//saves sheet
	ByteArrayOutputStream bos = new ByteArrayOutputStream();
	workbook.write(bos);
	bos.close();
	ByteArrayInputStream bis = new ByteArrayInputStream(bos.toByteArray());
	
	new File("target/.output-file/xlsx").mkdirs();
	WorksheetEncryptor.encrypt(
		bis, 
		new FileOutputStream("target/.output-file/xlsx/OpenAndModifyProtected.xlsx"), 
		DocumentType.XML, 
		"54321");
	bis.close();
	
	//read again and check
	XSSFWorkbook workbook2 = (XSSFWorkbook) WorksheetDecryptor.decrypt(
		new FileInputStream("target/.output-file/xlsx/OpenProtected.xlsx"),
		DocumentType.XML, 
		"54321");
	Assert.assertEquals("Gravado", workbook2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
    }

}
