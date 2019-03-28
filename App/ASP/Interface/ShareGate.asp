<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%option Explicit%>
<%
'/****************************************************************************************/
'PushGate.asp - 엡에서 푸쉬메시지 클릭시 넘어오는 페이지
'Date		: 2014.11.04
'Update	: 
'Writer		: Kim YoungSik
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//-------------------------------------------------------------------------------
%>
<%
'/****************************************************************************************/
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
Dim cont						'
Dim goURL
Dim iswin
Dim ProductCode
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'
goURL	= request("cont")
ProductCode = Request("ProductCode")
'goUrl = Server.URLEncode(cont)
iswin = "1"
%>
<% If Trim(ProductCode) = "" Then %>
<%
	response.WRite "aaaaaaaaaaaaaaaa"
	Response.end
	 %>
<script type="text/javascript">
	location.replace('/index.asp?goUrl=<%=goUrl%>&iswin=<%=iswin%>');
</script>
<% Else %>
<script type="text/javascript">
	location.replace('/index.asp?ProductCode=<%=ProductCode%>');
</script>
<% End If %>

