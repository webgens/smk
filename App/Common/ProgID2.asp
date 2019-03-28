<%
DIM strPageLoca : strPageLoca = Replace(Request.ServerVariables("HTTP_REFERER"), "http://" & Request.ServerVariables("HTTP_HOST"), "")

DIM ProgID
ProgID = strPageLoca
%>