package br.com.starcode.poisecurity.encryptor;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class XlsEncryptor {

    public static HSSFWorkbook encrypt(InputStream input, OutputStream output, String password)
	    throws IOException {
	throw new UnsupportedOperationException("XLS format does not support encryption!");
    }

    private XlsEncryptor() {
    }

}
