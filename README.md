Poi Security
============

Working with password protected office documents using Apache POI.

This project is focused on Excel (XLS and XLSX) spreadsheets.

## Different types of documents

Today we have the new office XML based documents (XLSX, DOCX, etc).

We have also the legacy binary document format (XLS, DOC, etc.).

The current project includes features for both versions.

Apache POI doesn't work well with legacy format.

## Different levels of protection

Office documents offer protection against modifications (read only mode), i.e., you can open and read, but cannot make any change unless types the corret password.

There's also protection against opening the document, i. e., encryption that wouldn't allow anyone to open the document without a password. 

The former approach has low security level since any other software can actually read the document and unlock it. The latter has stronger security because the entire file is encrypted.

## The tests

I've tested POI reading both XLS and XLSX documents in the following scenarios:

1. Reading a file protected against modification
2. Reading a file protected against opening (encrypted)
3. Reading a file protected against modification and opening (encrypted)

Then I've tested POI creating both XLS and XLSX documents in the following scenarios:

1. Creating a file protected against modification
2. Creating a file protected against opening (encrypted)
3. Creating a file protected against modification and opening (encrypted)

## Results 

### Reading XML based documents (XLSX)

1. **Passed**. It doesn't need any additional implementation.
2. **Passed**. It was necessary to use `Decryptor` class.
3. **Passed**. It was necessary to use `Decryptor` class.

### Reading legacy binary documents (XLS)

1. **Passed**. It doesn't need any additional implementation.
2. **Passed**. It was necessary to use `Biff8EncryptionKey` class.
3. **Passed**. It was necessary to use `Biff8EncryptionKey` class.

### Writing XML Based documents (XLSX)

1. **Passed**. It was used the method `protectSheet` from POI API. 
2. **Passed**. It was necessary to use `Encryptor` class.
3. **Passed**. It was necessary to use `Encryptor` class and the method `protectSheet`.

### Writing legacy binary documents (XLS)

1. **Passed**. It was used the method `protectSheet` from POI API. 
2. **Not Passed**. Encryption not supported.
2. **Not Passed**. Encryption not supported.

## Notes

"It doesn't need any additional implementation" actually means POI reads the file as if it haven't any protection.

And I'll discuss the other cases in the following topics:

### `Decryptor` class for reading XLSX

In this project, the class `XlsxDecryptor` is responsible for loading documents protected against opening.

Actually, this class wraps the API provided by `org.apache.poi.poifs.crypt.Decryptor`.

`Decryptor` reads a binary `InputStream` and returns the actual content of the document.

It means this class isn't coupled to any kind of document.

### `Biff8EncryptionKey` class for reading XLS

In this project, the class `XlsDecryptor` is responsible for loading documents protected against opening.

Actually, this class wraps the API provided by `org.apache.poi.hssf.record.crypto.Biff8EncryptionKey`.

It's necessary to call the static method `setCurrentUserPassword(password)` to define the password for opening the file.

Then, `new HSSFWorkbook(input)` constructor will use that password.

It's also necessary to call the static method again with a `null` as argument to clear the password for future documents.

This solution is thread-safe because the password is stores in a ThreadLocal variable.

### `protectSheet` method for write-protect XLSX

`XSSFSheet` provides the method `protectSheet` so you can put a password for modifications in sheets.

You need to do this for each sheet in the Workbook.

There's not a method to protect the document as a whole.

### `Decryptor` class for writing encrypted XLSX

In this project, the class `XlsxEncryptor` is responsible for outputting documents protected against opening.

Actually, this class wraps the API provided by `org.apache.poi.poifs.crypt.Encryptor`.

`Encryptor` reads a `InputStream` (the original document) and outputs the binary encrypted version of it.

This class isn't coupled to any kind of document.

## Java Excel API (jexcel)

Java Excel API is an anternative to Apache POI. However, it's somewhat more limited, so I tested as a second option.

See: http://stackoverflow.com/questions/14980717/what-is-the-better-api-to-reading-excel-sheets-in-java-jxl-or-apache-poi

Unfortunately, **there's no password protection features** unless Sheet protection as we already have in POI.

## TODO

Contribute to POI project:

  - Implement `writeProtectWorkbook` in XSSF
  - Fix `writeProtectWorkbook` in HSSF

