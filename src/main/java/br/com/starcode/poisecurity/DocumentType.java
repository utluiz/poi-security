package br.com.starcode.poisecurity;

import java.io.File;

public enum DocumentType {

    XML, LEGACY;

    public static DocumentType forFile(File file) {
	return forFileName(file.getName());
    }

    public static DocumentType forFileName(String fileName) {
	/*int posExtension = fileName.lastIndexOf('.');
	if (posExtension >= 0) {
	    return DocumentType.valueOf(fileName.substring(posExtension + 1)
		    .toUpperCase());
	}*/
	if (fileName.toLowerCase().endsWith("x")) {
	    return XML;
	}
	return LEGACY;
    }

}
