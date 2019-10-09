<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, com.acesoft.aceoffix.wordwriter.*, java.awt.*"%>

<%
	AceoffixCtrl aceCtrl=new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	aceCtrl.setTheme(ThemeType.Office2007);
	aceCtrl.setBorderStyle(BorderStyleType.BorderThin);
	aceCtrl.addCustomToolButton("save","SaveDocument()",1);
	
	WordDocument doc = new WordDocument();

	DataRegion title = doc.createDataRegion("title",
			DataRegionInsertType.After, "[home]");
	title.setValue("Introduction\n");
	
	title.getFont().setBold(true);
	title.getFont().setSize(20);
	title.getFont().setName("Aharoni");
	title.getFont().setItalic(false);
	
	ParagraphFormat titlePara = title.getParagraphFormat();
	
	titlePara.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
	
	titlePara.setLineSpacingRule(WdLineSpacing.wdLineSpaceMultiple);

	
	DataRegion body = doc.createDataRegion("body",
			DataRegionInsertType.After, "title");
	
	body.getFont().setBold(false);
	body.getFont().setItalic(true);
	body.getFont().setSize(10);
	
	body.getFont().setName("Berlin Sans FB Demi");
	
	//body.getFont().setName("Times New Roman");
	body.getFont().setColor(Color.RED);
	
	body.setValue("Aceoffix is a flexible and professional web component for Microsoft Office, with the simplified interfaces and powerful functions that make it terrific not only for online editing and saving Office documents, but also importing and exporting data from database to Office documents. Aceoffix supports many Office document formats, such as *.doc, *.docx, *.xls, *.xlsx, *.ppt, *.pptx, *.xml and *.rtf. \n");
	
	ParagraphFormat bodyPara = body.getParagraphFormat();
	
	bodyPara.setLineSpacingRule(WdLineSpacing.wdLineSpaceAtLeast);
	bodyPara.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
	bodyPara.setFirstLineIndent(21);

	
	DataRegion body2 = doc.createDataRegion("body2",
			DataRegionInsertType.After, "body");
	body2.getFont().setBold(false);
	body2.getFont().setSize(12);
	body2.getFont().setName("Arial Rounded MT Bold");
	body2.setValue("Without installing Microsoft Office at the server side, web developers can easily embed and call Microsoft Office in web pages, just like using a Java or .Net control. Aceoffix edits the real Microsoft Office documents online without converting any formats. Intuitive examples with source code are included to speed up your development time.  In general management systems base on Browser/Server architecture, developers have to manage Word/Excel documents by downloading and uploading. With Aceoffix, you can not only online view, edit and save Office documents on web, but also access the contents of them. In addition, Aceoffix has many other powerful functions as well, like read-only control, authority control, editable region control, forced revision mode, generating formal documents and etc.\n");
	
	ParagraphFormat bodyPara2 = body2.getParagraphFormat();
	bodyPara2.setLineSpacingRule(WdLineSpacing.wdLineSpace1pt5);
	bodyPara2.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
	bodyPara2.setFirstLineIndent(21);

	
	DataRegion body3 = doc.createDataRegion("body3",
			DataRegionInsertType.After, "body2");
	body3.getFont().setBold(false);
	body3.getFont().setColor(Color.getHSBColor(0, 128, 228));
	body3.getFont().setSize(14);
	body3.getFont().setName("Broadway");
	body3.setValue("Aceoffix includes a group of easy-to-use components. They are the object modules that are developed based on commonly used functions of Word/Excel. These components have the complete objects hierarchies. Developer will be able to understand and handle them easily. With simple code, they can accomplish the functions of Microsoft Office that they could hardly achieve before. Developers will be able to create their own business components based on these components as well. Developers do not have to face with the complex interfaces of Microsoft Office COM automation and VBA (Visual Basic for Applications). They will be able to save their time with Aceoffix. We make it easy to call the components of Aceoffix, and we insist on continuing efforts to keep the simplest calling interfaces of Aceoffix.\n");
	ParagraphFormat bodyPara3 = body3.getParagraphFormat();
	bodyPara3.setLineSpacingRule(WdLineSpacing.wdLineSpaceDouble);
	bodyPara3.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
	bodyPara3.setFirstLineIndent(21);

	DataRegion body4 = doc.createDataRegion("body4",
			DataRegionInsertType.After, "body3");
	body4.setValue("[image]doc/image.png[/image]");
	
	ParagraphFormat bodyPara4 = body4.getParagraphFormat();
	bodyPara4.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);

	aceCtrl.bind(doc);
	
	aceCtrl.setJsFunction_AfterDocumentSaved("SaveOK()");

	aceCtrl.setMenubar(false);

	aceCtrl.setSaveFilePage("savefile.jsp");
	aceCtrl.openDocument("doc/merge.doc",OpenModeType.docNormalEdit,"Tom");
	

%>

	<!DOCTYPE HTML>
<!--[if lt IE 7]><html class="no-js lt-ie9 lt-ie8 lt-ie7"><![endif]-->
<!--[if IE 7]><html class="no-js lt-ie9 lt-ie8"><![endif]-->
<!--[if IE 8]><html class="no-js lt-ie9"><![endif]-->
<!--[if gt IE 8]><!--><html class="no-js"><!--<![endif]--><head>

<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

<!-- Mobile Specific Metas -->
<meta name="viewport" content="width=device-width, initial-scale = 1.0, user-scalable = no">
<meta http-equiv="X-UA-Compatible" content="IE=Edge">

<title>Acesoft Corporation - Only using code to create a formatted Word document from a blank file</title>

<!-- Theme styles -->
<link href="../css/acestyle.css" rel="stylesheet">
<link rel='stylesheet' id='contact-form-7-css'  href='../css/styles.css?ver=4.2.2' type='text/css' media='all' />
</head>

<body>
<!--[if lt IE 7]>
	<p class="chromeframe">You are using an <strong>outdated</strong> browser. Please <a href="http://browsehappy.com/">upgrade your browser</a> or <a href="http://www.google.com/chromeframe/?redirect=true">activate Google Chrome Frame</a> to improve your experience.</p>
	<![endif]-->

<div id="banner-wrap" class="clearfix">
	<div id="top-nav" class="container clearfix">
        <div class="grid_10 center alpha">
        	Web Development Platform for Microsoft Office
        </div>
        <div class="grid_2 right omega">
        	<a href="http://www.aceoffix.com/index.php/contact-us/"><img src="../image/contactw.png" alt="contact">&nbsp;&nbsp;Contact Us</a>
        </div>
	</div>
    <div id="nav-wrap" class="clearfix">
        <div class="container">
            <div id="logo" class="grid_3 alpha"><a href="http://www.aceoffix.com"><img src="../image/aceoffixlogo.png" height="48" width="229" alt="Aceoffix" class="scale-with-grid" /></a></div>
            <div class="grid_9 omega">
                <nav class="nav">
                    <ul class="nav-list"><li class="page_item page-item-2 current_page_item menu-item-8"><a href="http://www.aceoffix.com/">Home</a></li>
<li class="page_item menu-item-244"><a href="http://www.aceoffix.com/index.php/overview/">Products</a>
</li>
<li class="page_item menu-item-224"><a href="http://www.aceoffix.com/index.php/download/">Download</a>
</li>
<li class="page_item menu-item-236"><a href="http://www.aceoffix.com/index.php/recent-update/">Support</a>
</li>
<li class="page_item menu-item-205"><a href="http://www.aceoffix.com/index.php/purchase/">Purchase</a>
</li>
<li class="page_item menu-item-17"><a href="http://www.aceoffix.com/index.php/about-us/">About Us</a></li>
</ul>                </nav>
            </div>
        </div>
    </div>


    <div id="title-wrap" class="clearfix">
        <div class="container left">
        Current Page:  Example Center
            <h1>Example Center - Aceoffix V5</h1>
            
        </div>
    </div>
    		
    </div>
    <br/>
    <div class="container left">
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Only using code to create a formatted Word document from a blank file</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div class="container left">
	

<!--************** Start of AceoffixCtrl JavaScript ************************-->

    <script type="text/javascript">
		function SaveDocument(){
			document.getElementById("AceoffixCtrl1").SaveDocument();
		}
		function SaveOK(){
		document.getElementById("AceoffixCtrl1").Alert("Save Succeed");
		}
		</script>

<!--************** End of AceoffixCtrl JavaScript ************************-->
<div style="width:auto; height:800px;">
      <%= aceCtrl.getHtmlCode("AceoffixCtrl1")%>
      </div>
			       </div><br/>   
<div id="quicklink-wrap" class="clearfix">
	<div class="container">
    	<h1>Quick Link</h1>
       <div class="footer-column grid_3 alpha">
       		<h3>Product</h3>
			<a href="http://www.aceoffix.com/index.php/overview/">Overview</a>
<a href="http://www.aceoffix.com/index.php/functions-features/">Functions &#038; Features</a>
<a href="http://www.aceoffix.com/index.php/trial-license/">Trial &#038; License</a>
<a href="http://www.aceoffix.com/index.php/purchase/">Purchase</a>
       </div>
       <div class="footer-column grid_3 mobile-omega2">
			<h3>Download</h3>
			<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for ASP.NET</a>
<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for JAVA</a>
<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for PHP</a>
       </div>
       <div class="footer-column grid_3 mobile-alpha2">
			<h3>Support</h3>
			<a href="http://www.aceoffix.com/index.php/technical-faq/">Technical FAQ</a>
<a href="http://www.aceoffix.com/index.php/videos/">Videos</a>
<a href="http://www.aceoffix.com/index.php/recent-update/">Recent Update</a>
       </div>
       <div class="footer-column grid_3 omega">
			<h3>Follow Us</h3>
			<a href="https://www.facebook.com/Aceoffix-152929314763836/">Facebook</a>
<a href="https://twitter.com/Aceoffix">Twitter</a>
       </div>
    </div>
</div>

<div id="footer-wrap" class="clearfix">
    <div class="container center">
		<a href="http://www.aceoffix.com/">Home</a>
<a href="http://www.aceoffix.com/index.php/acesoft-customer-privacy-policy/">Privacy Policy</a>
<a href="http://www.aceoffix.com/index.php/acesoft-terms-of-use/">Terms of Use</a>
<a href="http://www.aceoffix.com/index.php/contact-us/">Contact Us</a>
<a href="http://www.aceoffix.com/index.php/sitemap/">Sitemap</a>
		<br>
		&copy; 2018 Acesoft Corporation. All rights reserved.
    </div>
</div>

</body>
</html>   
	