package br.com.starcode.poisecurity.decryptor;

import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxDecryptor {

    public static XSSFWorkbook decrypt(InputStream input, String password)
	    throws IOException {

	POIFSFileSystem filesystem = new POIFSFileSystem(input);

	EncryptionInfo info = new EncryptionInfo(filesystem);
	Decryptor d = Decryptor.getInstance(info);

	try {
	    if (!d.verifyPassword(password)) {
		throw new RuntimeException(
			"Unable to process: document is encrypted");
	    }

	    InputStream dataStream = d.getDataStream(filesystem);

	    // parse dataStream
	    return new XSSFWorkbook(dataStream);

	} catch (GeneralSecurityException ex) {
	    throw new RuntimeException("Unable to process encrypted document",
		    ex);
	}

    }

    private XlsxDecryptor() {
    }

}
