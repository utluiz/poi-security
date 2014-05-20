package br.com.starcode.poisecurity;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import br.com.starcode.poisecurity.decryptor.WorksheetDecryptor;

public class DecryptXlsxTests {
    
    @Test
    public void testModifyProtectedXLSX() throws IOException {
	System.out.println("### testModifyProtectedXLSX");
	XSSFWorkbook workbook = new XSSFWorkbook(getClass().getResourceAsStream("xlsx/ModifyProtected.xlsx"));
	String conteudoCelula = workbook.getSheetAt(0).getRow(0).getCell(0)
		.getStringCellValue();
	Assert.assertEquals("Teste criptografia", conteudoCelula);
    }
    
    @Test
    public void testOpenProtectedXLSX() throws IOException {
	System.out.println("### testOpenProtectedXLSX");
	XSSFWorkbook workbook = (XSSFWorkbook) WorksheetDecryptor.decrypt(
		getClass().getResourceAsStream("xlsx/OpenProtected.xlsx"),
		DocumentType.XML, 
		"12345");
	String conteudoCelula = workbook.getSheetAt(0).getRow(0).getCell(0)
		.getStringCellValue();
	Assert.assertEquals("Teste criptografia", conteudoCelula);
    }

    @Test
    public void testOpenAndModifyProtectedXLSX() throws IOException {
	System.out.println("### testOpenAndModifyProtectedXLSX");
	XSSFWorkbook workbook = (XSSFWorkbook) WorksheetDecryptor.decrypt(
		getClass().getResourceAsStream("xlsx/OpenAndModifyProtected.xlsx"),
		DocumentType.XML, 
		"12345");
	String conteudoCelula = workbook.getSheetAt(0).getRow(0).getCell(0)
		.getStringCellValue();
	Assert.assertEquals("Teste criptografia", conteudoCelula);
    }
    
}
