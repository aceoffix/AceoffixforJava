<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, java.awt.*"%>
<%
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	aceCtrl.setTheme(ThemeType.Office2007);
	aceCtrl.setSaveFilePage("savefile.jsp");
	aceCtrl.addCustomToolButton("Save", "SaveDocument()", 1);
	aceCtrl.openDocument("doc/introduce.doc",OpenModeType.docNormalEdit,"Tom");
	out.println(aceCtrl.getHtmlCode("AceoffixCtrl1"));
%>
