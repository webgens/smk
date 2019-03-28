<%
IF U_NUM = "" THEN
		Response.Redirect "/ASP/Member/SubLogin.asp?ProgID=" & Server.URLEncode(ProgID)
		Response.end
END IF
%>