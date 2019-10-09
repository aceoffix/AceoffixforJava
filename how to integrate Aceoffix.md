# How to integrate Aceoffix into your web project

1. Copy the examples5\WEB-INF\lib\aceoffix5.6.0.1.jar to the WEB-INF\lib folder of your website.
 

2. Add the following code to the web.xml of your web site:

        <servlet>
            <servlet-name>xafletrt</servlet-name> 
            <servlet-class>com.acesoft.aceoffix.xafletrt.Server</servlet-class> 
        </servlet>
        <servlet-mapping> 
            <servlet-name>xafletrt</servlet-name> 
            <url-pattern>/server.ace</url-pattern> 
        </servlet-mapping>  
        <servlet-mapping> 
            <servlet-name>xafletrt</servlet-name> 
            <url-pattern>/pluginsetup.exe</url-pattern>
        </servlet-mapping>
        <servlet-mapping>
            <servlet-name>xafletrt</servlet-name>
            <url-pattern>/aceoffix.js</url-pattern>
        </servlet-mapping>
        <servlet-mapping>
            <servlet-name>xafletrt</servlet-name>
            <url-pattern>/jquery.min.js</url-pattern>
        </servlet-mapping>
        <servlet-mapping>
            <servlet-name>xafletrt</servlet-name>
            <url-pattern>/acebstyle.css</url-pattern>
        </servlet-mapping>
        <mime-mapping>  
            <extension>mht</extension>  
            <mime-type>message/rfc822</mime-type>  
        </mime-mapping>

3. Add the following code to the parent page of which page you want to edit Office document.
Write the following code in the <head> tag. 

        <script type=text/javascript src="jquery.min.js"></script>
        <script type="text/javascript" src="aceoffix.js" id="ace_js_main"></script>
    
Note: The paths of jquery.min.js  and aceoffix.js are relative to the root of your website.
Write the following link to pop-up Aceoffix to edit Office document. We assume that the page which contains Aceoffix control is word/editword.jsp.

    <a href="javascript:AceBrowser.openWindowModeless('word/editword.jsp', 'width=1200px;height=800px;');" > Edit Word document online</a>
    
Write the following code in editword.jsp.

    <%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
    <%@ page import="com.acesoft.aceoffix.*"%>
    <%
    AceoffixCtrl aceCtrl1 = new AceoffixCtrl(request);
    aceCtrl1.setServerPage("../server.ace"); //Required
    aceCtrl1.setSaveFilePage("savefile.jsp");
    aceCtrl1.openDocument("../doc/editword.doc", OpenModeType.docNormalEdit, "John Scott");
    %>
    
Then write the code where you want to show Aceoffix in editword.jsp.

    <div style="width:auto; height:600px;">
        <%= aceCtrl1.getHtmlCode("AceoffixCtrl1")%>
    </div>

Write the savefile.jsp if your user want to save his/her document.

    <%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
    <%@ page import="com.acesoft.aceoffix.*"%>
    <%
    // This page is used to receive the file saved by AceoffixCtrl.
    FileRequest freq = new FileRequest(request, response);
    freq.saveToFile(request.getSession().getServletContext().getRealPath("doc/") + "\\" + freq.getFileName());
    freq.close();
    %>
    
4. Congratulations. Your user can open, edit, save the Word documents in your web project online now. You have finished the integration of Aceoffix. Please refer to the examples5 to learn how to use more features of Aceoffix.
 
