package br.com.starcode.poisecurity.encryptor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import br.com.starcode.poisecurity.DocumentType;

public abstract class WorksheetEncryptor {

    public static void encrypt(InputStream input, OutputStream output, DocumentType worksheetType,
	    String password) throws IOException {
	if (DocumentType.LEGACY.equals(worksheetType)) {
	    XlsEncryptor.encrypt(input, output, password);
	} else {
	    XlsxEncryptor.encrypt(input, output, password);
	}
    }

    public static void encrypt(String fileName, OutputStream output, String password)
	    throws IOException {
	encrypt(new FileInputStream(fileName), output, 
		DocumentType.forFileName(fileName), password);
    }

    public static void encrypt(File file, OutputStream output, String password)
	    throws IOException {
	encrypt(new FileInputStream(file), output, DocumentType.forFile(file),
		password);
    }

    private WorksheetEncryptor() {
    }

}
