<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Gate.asp - 앱 Gate 페이지
'Date		: 2019.01.04
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정 ------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드 ---------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>
;jlk;lj;kl;lkj
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

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절
DIM sqlQuery

DIM i
DIM j
DIM x
DIM y

DIM IsApp
DIM objSC

Dim DB_DelFlag
Dim DB_DormancyFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<script src="/JS/jquery-1.11.1.min.js?ver=<%=U_DATE&U_TIME%>"></script>

<script type="text/javascript">
	var u_no		 = "";
	var e_url		 = "";
	var home_domain	 = "<%=HOME_URL%>";
	var isApp		 = "";
</script>
<script src="/JS/dev/App.js?ver=<%=U_DATE&U_TIME%>"></script>
<%
IsApp		= sqlFilter(Request("IsApp"))
		
'# APP 으로 실행시 U_ISAPP 쿠키를 생성한다.
IF IsApp = "Y" THEN
		Response.Cookies("U_ISAPP")		 = Encrypt("Y")
END IF

If ISNULL(U_NUM) OR ISEMPTY(U_NUM) Then
	U_NUM = ""
End If

If U_NUM <> "" Then
	SET oCmd = SErver.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_Member_Select_By_UserID"

			.Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 30, U_ID)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	SET oCmd = Nothing

	IF NOT oRs.EOF THEN
			DB_DelFlag			 = oRs("DelFlag")
			DB_DormancyFlag		 = oRs("DormancyFlag")
	ELSE
			DB_DelFlag			 = "Y"
			DB_DormancyFlag		 = "Y"
	END IF
	oRs.Close

	If DB_DelFlag = "Y" Or DB_DormancyFlag = "Y" Then
		Response.Cookies("UIP").Expires			 = Now - 1000
		Response.Cookies("UMFLAG").Expires		 = Now - 1000
		Response.Cookies("UNUM").Expires		 = Now - 1000
		Response.Cookies("UID").Expires			 = Now - 1000
		Response.Cookies("UNAME").Expires		 = Now - 1000
		Response.Cookies("UETYPE").Expires		 = Now - 1000
		Response.Cookies("UETYPE").Expires		 = Now - 1000
		Response.Cookies("UGROUP").Expires		 = Now - 1000
	Else
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Login_Insert"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , U_NUM)
				.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,  1, "A")
				.Parameters.Append .CreateParameter("@LoginIP",		 adVarChar,	 adParamInput, 15, U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
	End If

End If

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
<div id="intro" style="background:#fff url('/images/app_loading_screen.gif') center no-repeat; background-size:cover; position:fixed; left:0; top:0; width:100%; height:100%; z-index:99999; display:none;"></div>

<script type="text/javascript">
	$("#intro").show();

	//setTimeout("$('#intro').hide();", 3000);
</script>


<%IF IsApp = "Y" THEN%>
<script type="text/javascript">
	toApp({method:'getInstallationId'});
</script>
<%ELSE%>
<script type="text/javascript">
	//document.write "";
	location.replace("/");
</script>
<%END IF%>