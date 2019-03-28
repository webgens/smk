<%
IF U_NUM = "" THEN
		Response.Redirect "/ASP/Mypage/Login.asp?ProgID=" & Server.URLEncode(ProgID)
		Response.end
END IF
%>