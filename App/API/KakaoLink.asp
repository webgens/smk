<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%option Explicit%>
<%
'/****************************************************************************************/
'KakaoLink.asp - 카카오톡에서 넘어오는 페이지
'Date		: 2018.12.06
'Update	: 
'Writer		: Hong
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//-------------------------------------------------------------------------------

%>
<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include virtual = "/Common/ProgID1.asp" -->
<%
'/****************************************************************************************/
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

Dim PageCode						'
Dim ItemCode
Dim goUrl
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'
PageCode	= SqlFilter(request("PageCode"))
ItemCode	= SqlFilter(request("ItemCode"))

If PageCode = "P" Then
	goUrl = "/ASP/Product/ProductDetail.asp?ProductCode=" & ItemCode
Else
	goUrl  = "/index.asp"
End If
%>
<script type="text/javascript">
	location.replace('<%=goUrl%>');
</script>

