package br.com.starcode.poisecurity;

import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import br.com.starcode.poisecurity.decryptor.WorksheetDecryptor;

public class DecryptXlsTests {
    
    @Test
    public void testModifyProtectedXLS() throws IOException {
	System.out.println("### testModifyProtectedXLS");
	HSSFWorkbook workbook = new HSSFWorkbook(getClass().getResourceAsStream("xls/ModifyProtected.xls"));
	String conteudoCelula = workbook.getSheetAt(0).getRow(0).getCell(0)
		.getStringCellValue();
	Assert.assertEquals("Teste criptografia", conteudoCelula);
    }
    
    @Test
    public void testOpenProtectedXLS() throws IOException {
	System.out.println("### testOpenProtectedXLS");
	HSSFWorkbook workbook = (HSSFWorkbook) WorksheetDecryptor.decrypt(
		getClass().getResourceAsStream("xls/OpenProtected.xls"),
		DocumentType.LEGACY, 
		"12345");
	String conteudoCelula = workbook.getSheetAt(0).getRow(0).getCell(0)
		.getStringCellValue();
	Assert.assertEquals("Teste criptografia", conteudoCelula);
    }

    @Test
    public void testOpenAndModifyProtectedXLS() throws IOException {
	System.out.println("### testOpenAndModifyProtectedXLS");
	HSSFWorkbook workbook = (HSSFWorkbook) WorksheetDecryptor.decrypt(
		getClass().getResourceAsStream("xls/OpenAndModifyProtected.xls"),
		DocumentType.LEGACY, 
		"12345");
	String conteudoCelula = workbook.getSheetAt(0).getRow(0).getCell(0)
		.getStringCellValue();
	Assert.assertEquals("Teste criptografia", conteudoCelula);
    }
    
}
