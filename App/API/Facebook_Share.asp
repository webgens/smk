<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'EventView.asp - 이벤트 내용
'Date		: 2019.01.12
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "EV"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%

'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

Dim ProductCode

Dim staus
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

ProductCode = Request("ProductCode")

staus = request("status")
%>

<script type="text/javascript">
	var u_no		 = "<%=U_NUM%>";
	var e_url		 = "<%=Server.URLEncode(ProgID)%>";
	var home_domain	 = "<%=HOME_DOMAIN%>";
	var isApp		 = "<%=U_ISAPP%>";
</script>
<!-- 페이스북 -->
<script src="/JS/jquery/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="/JS/dev/App.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<script>
function face()
{
	var url = "https://www.shoemarker.co.kr/ASP/Product/ProductDetail.asp?ProductCode=<%=ProductCode%>";
	var url1 = "https://app.shoemarker.co.kr/api/facebook_share.asp?status=Y";
	location.href='https://www.facebook.com/dialog/share?app_id=<%=FACEBOOK_LOGIN_CLIENTID%>&display=popup&href='+url+'&redirect_uri='+url1;
}
</script>
<!-- 페이스북 -->

<script type="text/javascript">
<% If staus = "" Then %>
face();
<% Else %>
APP_PopupHistoryBack();
<% End If %>

</script>
</body>
</html>