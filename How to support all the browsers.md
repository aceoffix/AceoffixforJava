# How to support all the browsers

Recently, Chrome and Firefox browsers claimed that they do not support NPAPI plugins. Aceoffix V4 and the earlier versions cannot run in the latest Chrome and Firefox browsers. So we develop a new feature to solve this problem in Aceoffix V5.

How to support all the browsers in Aceoffix V5? Take the following steps:

1. Copy the examples5\WEB-INF\lib folder to the WEB-INF\lib folder of your website and overwrite it.
Add the following code to the web.xml of your web site:

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

Note: The aceoffix5.jar, web.xml  and pluginsetup.exe have updated. You should use the new files from the “Aceoffix V5.6 Ultimate Edition for Java” download package.

2. Add the following JavaScript references to your web page.

        <script type=text/javascript src="jquery.min.js"></script>
        <script type="text/javascript" src="aceoffix.js" id="ace_js_main"></script>
	
Note: The paths of jquery.min.js  and aceoffix.js are relative to the root of your website.
 
![](https://github.com/aceoffix/AceoffixforJava/blob/master/examples5/image/support_1.png?raw=true)

3. Change the old link to new link.

Refer to examples5 /index.jsp to learn how to change the old link to new link.

 The old link:

 ![](https://github.com/aceoffix/AceoffixforJava/blob/master/examples5/image/support_2.png?raw=true)

For example: The old link is <a href=”word/editword.jsp”>Edit Word document</a>

The new link:

![](https://github.com/aceoffix/AceoffixforJava/blob/master/examples5/image/support_3.png?raw=true)
    
The new link is 

        <a href=”javascript:AceBrowser.openWindowModeless('word/editword.jsp', 'width=1200px;height=800px;');”>Edit Word document</a>
	
4. The web page with Aceoffix will pop up when you click the new link in Chrome or Firefox browsers.  This new link can also work in other browsers (e.g. IE,  Opera, Edge).

![](https://github.com/aceoffix/AceoffixforJava/blob/master/examples5/image/support_4.png?raw=true)

 

