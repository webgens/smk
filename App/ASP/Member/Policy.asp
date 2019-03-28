<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Policy.asp - 이용약관 가져오기
'Date		: 2018.11.30
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/ProgID1.asp" -->

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


DIM sType
Dim Title
Dim Contents
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



sType			 = Request("sType")
IF sType = "" THEN
%>
		<script type="text/javascript" src="/JS/dev/App.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
		<script type="text/javascript">
			APP_PopupHistoryBack_Alert("정보가 없습니다.<br />다시 시도하여 주십시오.");
		</script>
<%
		Response.End
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Policy_Select_By_IDX"

		.Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , sType)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF oRs.EOF THEN
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
		<script type="text/javascript" src="/JS/dev/App.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
		<script type="text/javascript">
			APP_PopupHistoryBack_Alert("정보가 없습니다.<br />다시 시도하여 주십시오.");
		</script>
<%
		Response.End
ELSE
		Title	 = oRs("Title")
		Contents = oRs("Contents")
END IF
oRs.Close
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/PopupTop.asp" -->



	   <div class="area-pop" id="PolicyPopup-pop">
			<div class="full terms">
				<div class="tit-pop">
					<p class="tit" id="popTitle"><%=Title%></p>
					<button class="btn-hide-pop" onclick="APP_PopupHistoryBack();">닫기</button>
				</div>

				<div class="container-pop">
					<div class="contents">
						<div class="wrap-agree-changedTerm" id="popCont">
							<%=Contents%>
						</div>
					</div>
				</div>
			</div>
		</div>



<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/PopupBottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>