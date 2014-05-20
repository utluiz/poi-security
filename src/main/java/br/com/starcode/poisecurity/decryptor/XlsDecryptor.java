package br.com.starcode.poisecurity.decryptor;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class XlsDecryptor {

    /**
     * When HSSF receive encrypted file, it tries to decode it with MSOffice
     * build-in password. Use static method setCurrentUserPassword(String
     * password) of org.apache.poi.hssf.record.crypto.Biff8EncryptionKey to set
     * password. It sets thread local variable. Do not forget to reset it to
     * null after text extraction.
     * 
     * (http://poi.apache.org/encryption.html)
     * 
     * @param input
     * @param password
     * @return
     * @throws IOException
     */
    public static HSSFWorkbook decrypt(InputStream input, String password)
	    throws IOException {

	try {

	    Biff8EncryptionKey.setCurrentUserPassword(password);

	    return new HSSFWorkbook(input);

	} catch (Exception ex) {
	    throw new RuntimeException("Unable to process encrypted document",
		    ex);
	} finally {
	    Biff8EncryptionKey.setCurrentUserPassword(null);
	}

    }

    private XlsDecryptor() {
    }

}
