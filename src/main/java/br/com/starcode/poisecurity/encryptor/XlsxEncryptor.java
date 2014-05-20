package br.com.starcode.poisecurity.encryptor;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class XlsxEncryptor {

    public static void encrypt(InputStream input, OutputStream output, String password)
	    throws IOException {

	try {
	    
        	POIFSFileSystem fs = new POIFSFileSystem();
        	EncryptionInfo info = new EncryptionInfo(fs, EncryptionMode.agile);
        	// EncryptionInfo info = new EncryptionInfo(fs, EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);
        
        	Encryptor enc = info.getEncryptor();
        	enc.confirmPassword(password);
        
        	OPCPackage opc = OPCPackage.open(input);
        	OutputStream os = enc.getDataStream(fs);
        	opc.save(os);
        	opc.close();
        
        	fs.writeFilesystem(output);
        	output.close();
        	
	} catch (GeneralSecurityException e) {
	    throw new RuntimeException(e);
	} catch (InvalidFormatException e) {
	    throw new RuntimeException(e);
	}

    }

    private XlsxEncryptor() {
    }

}
