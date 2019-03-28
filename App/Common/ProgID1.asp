<%
DIM strPageLoca : strPageLoca = Request.ServerVariables("URL")
IF LCase(strPageLoca) = "/index.asp" THEN strPageLoca = ""
IF Request.ServerVariables("QUERY_STRING") <> "" THEN
		strPageLoca = strPageLoca & "?" & Request.ServerVariables("QUERY_STRING")
END IF

DIM ProgID
ProgID = strPageLoca
%>