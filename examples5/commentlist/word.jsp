<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, java.awt.*"%>

<%
	
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	
	aceCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
	aceCtrl.setOfficeToolbars(false);
	aceCtrl.addCustomToolButton("New Comment", "InsertComment()", 3); 	
	aceCtrl.setSaveFilePage("savefile.jsp");
	aceCtrl.addCustomToolButton("Save", "SaveDocument()", 1);
	
	aceCtrl.openDocument("doc/introduce.doc",OpenModeType.docForcedRevision,"Tom");
	 
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

<title>Acesoft Corporation - Add new comments in two ways and navigate comment list</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Add new comments in two ways and navigate comment list</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div class="container left">
	

<!--************** Start of AceoffixCtrl JavaScript ************************-->
   <script language="javascript" type="text/javascript">
    function SaveDocument() {
        document.getElementById("AceoffixCtrl1").SaveDocument();
    }
    
   function AfterDocumentOpened() {
            refreshList();
        }
        
        function refreshList() {
            var sMac = "Function getComments() " + "\r\n"
                     + "Dim cmts As String " + "\r\n"
                     + "For i=1 To ActiveDocument.Comments.Count "+ "\r\n"
                     + "    cmts = cmts +ActiveDocument.Comments.Item(i).Author & \":\" & ActiveDocument.Comments.Item(i).Range.Text + \"||\" " + "\r\n"
                     + "Next" + "\r\n"
                     + "getComments = cmts" + "\r\n"
                     + "End Function ";

            var sComments = document.getElementById("AceoffixCtrl1").RunMacro("getComments", sMac);

            var arr = sComments.split("||");

            document.getElementById("ul_Comments").innerHTML = "";
            for (var i = 0; i < arr.length-1 ; i++) {
                document.getElementById("ul_Comments").innerHTML += "<li><a href='#' onclick='goToComment("+(i+1)+")'>"+arr[i]+"</a></li>"
            }
            
        }
        
        function getComment(index){
            var sMac = "Function getCmtTxt() " + "\r\n"
                     + "getCmtTxt = ActiveDocument.Comments.Item(" + index + ").Range.Text " + "\r\n"
                     + "End Function ";

            return document.getElementById("AceoffixCtrl1").RunMacro("getCmtTxt", sMac);
        }

        function goToComment(index)
        {
	        //refreshList();
	        var sMac = "function myfunc() " + "\r\n"
                     + "ActiveDocument.Comments.Item("+index+").Edit " + "\r\n"
                     + "End function ";

	        document.getElementById("AceoffixCtrl1").RunMacro("myfunc", sMac);
	        
        }

        function addComment(txt) {
            var sMac = "function myfunc() " + "\r\n"
                     + "Selection.Comments.Add Range:=Selection.Range " + "\r\n"
                     + "Selection.TypeText Text:=\"" + txt + "\" " + "\r\n"
                     + "On Error Resume Next " + "\r\n"
                     + "ActiveWindow.ActivePane.Close " + "\r\n"
                     + "End function ";
            document.getElementById("AceoffixCtrl1").RunMacro("myfunc", sMac);
        }
        function Button1_onclick() {
            addComment(document.getElementById("Text1").value);
            refreshList();
            document.getElementById("Text1").value = "";
        }

		function Button2_onclick() {
            refreshList();
        }

        function InsertComment() {
            document.getElementById("AceoffixCtrl1").WordInsertComment();
            var sMac = "function myfunc() " + "\r\n"
                     + "On Error Resume Next " + "\r\n"
                     + "ActiveWindow.ActivePane.Close " + "\r\n"
                     + "End function ";
            document.getElementById("AceoffixCtrl1").RunMacro("myfunc", sMac);
        }
</script>
<!--************** End of AceoffixCtrl JavaScript ************************-->
     
 <form id="form1" >
    <div  style=" width:1300px; height:700px;">
        <div id="Div_Comments" style=" float:left; width:200px; height:700px; border:solid 1px red;">
        <h3>Comment List</h3>
		<input id="Button2" type="button" value="Refresh" onclick="return Button2_onclick()" />
        <ul id="ul_Comments">
            
        </ul>
        </div>
        <div style=" width:1050px; height:700px; float:right;">
            <input id="Button1" type="button" value="Insert Comment" onclick="return Button1_onclick()" /> 
<input id="Text1" type="text" />
 <%= aceCtrl.getHtmlCode("AceoffixCtrl1")%>
     </div>
    </div>

    </form>   
    
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
    