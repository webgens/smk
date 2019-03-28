<%
IF U_NUM = "" THEN
		Response.Redirect LOGIN_URL & "?ProgID=" & Server.URLEncode(ProgID)
		Response.end
END IF
%>