<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, com.acesoft.aceoffix.wordreader.*,java.io.*"%>
<%

String filePath=request.getSession().getServletContext().getRealPath("wordSplit/doc/")+"/";
WordDocument doc=new WordDocument(request,response);
byte[] bWord;

DataRegion dr1=doc.openDataRegion("Paragraph1");
bWord=dr1.getFileBytes();
FileOutputStream fos1=new FileOutputStream(filePath+"Paragraph1.doc");
fos1.write(bWord);
fos1.flush();
fos1.close();

DataRegion dr2=doc.openDataRegion("Paragraph2");
bWord=dr2.getFileBytes();
FileOutputStream fos2=new FileOutputStream(filePath+"Paragraph2.doc");
fos2.write(bWord);
fos2.flush();
fos2.close();

DataRegion dr3=doc.openDataRegion("Paragraph3");
bWord=dr3.getFileBytes();
FileOutputStream fos3=new FileOutputStream(filePath+"Paragraph3.doc");
fos3.write(bWord);
fos3.flush();
fos3.close();

doc.close();
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Extracts data from the Word document</title>
    <style type="text/css">
        .TypeValue{ color:#0000FF; text-decoration:underline;}
    </style>
</head>
<body>
    
</body>
</html>
