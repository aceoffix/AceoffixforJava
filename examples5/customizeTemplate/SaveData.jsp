<%@ page language="java"
	import="java.util.*,java.text.*,com.acesoft.aceoffix.*,com.acesoft.aceoffix.wordreader.*"
	pageEncoding="UTF-8"%>

<%
	WordDocument doc = new WordDocument(request, response);
	
    try{
        DataRegion Name = doc.openDataRegion("Name");
        out.print("Get the value of ACE-Name from back end:" + Name.getValue());
    }
    catch(Exception e){
        out.print("No DataRegion named ACE-Name submitted from client side.");
    }
	doc.showPage(400, 300);
	doc.close();
%>
