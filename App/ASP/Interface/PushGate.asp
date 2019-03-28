<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%option Explicit%>
<%
'*****************************************************************************************'
'PushGate.asp - 엡에서 푸쉬메시지 클릭시 넘어오는 페이지
'Date		: 2018.08.16
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정 ------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

%>
<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체
	
DIM PushValue
DIM SplitPushValue
DIM Idx							'# 푸쉬메시지 idx
DIM PushType
DIM Push_Idx

DIM Link
DIM IsWin
DIM AlertCount
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


PushValue		 = SqlFilter(request("PushIdx"))

SplitPushValue	 = Split(PushValue, "AA")
PushType		 = SplitPushValue(0)
Idx				 = SplitPushValue(1)


SET oConn		 = ConnectionOpen()						'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'# 푸쉬 읽음 처리
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_App_Push_Result_Update_For_PushRead"

		.Parameters.Append .CreateParameter("@Idx",			 adInteger, adParamInput,  ,	 Idx)

		.Execute, , adExecuteNoRecords
END With
SET oCmd = Nothing


'# 푸쉬 내용
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_App_Push_Log_Select_By_IDX"

		.Parameters.Append .CreateParameter("@Idx", adInteger, adParamInput, , Idx)
END With
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		Link		 = oRs("Link")
ELSE
		Link		 = ""
END IF
oRs.Close

SET oRs = Nothing
oConn.Close
SET oConn = Nothing


IF ISNULL(Link) OR Link = "" THEN Link = ""
%>
<script src="/JS/App.js?<%=U_DATE%><%=U_TIME%>"></script>
<script type="text/javascript">
<!--
	//APP_BadgeUpdate(<%=ALERTCOUNT%>);
	setTimeout(function () {
		location.replace('/Index.asp?GoUrl=<%=Server.UrlEncode(Link)%>&IsWin=<%=IsWin%>');
	},10);
//-->
</script>

