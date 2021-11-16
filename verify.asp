
<!-- #include file="aspJSON1.19.asp"-->
<%
        Dim recaptcha_secret, sendstring, objXML
        recaptcha_secret = "secret_key_here"

        sendstring = "https://www.google.com/recaptcha/api/siteverify?secret=" & recaptcha_secret & "&response=" & Request.form("g-recaptcha-response")

        Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
        objXML.Open "GET", sendstring, False

        objXML.Send
        Set oJSON = New aspJSON
        result = (objXML.responseText)
        oJSON.loadJSON(result)
        If oJSON.data("success")=True then

        Response.Write ("Doğrulama Başarılı")

        else
        Response.Write ("Doğrulama Başarısız")
        end if
        
        Set objXML = Nothing
        Set oJSON = Nothing

%>
