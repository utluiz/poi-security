package br.com.starcode.poisecurity.decryptor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;

import br.com.starcode.poisecurity.DocumentType;

public abstract class WorksheetDecryptor {

    public static Workbook decrypt(InputStream input, DocumentType worksheetType,
	    String password) throws IOException {
	if (DocumentType.LEGACY.equals(worksheetType)) {
	    return XlsDecryptor.decrypt(input, password);
	} else {
	    return XlsxDecryptor.decrypt(input, password);
	}
    }

    public static Workbook decrypt(String fileName, String password)
	    throws IOException {
	return decrypt(new FileInputStream(fileName),
		DocumentType.forFileName(fileName), password);
    }

    public static Workbook decrypt(File file, String password)
	    throws IOException {
	return decrypt(new FileInputStream(file), DocumentType.forFile(file),
		password);
    }

    private WorksheetDecryptor() {
    }

}
