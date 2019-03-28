<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CompanyInfo.asp - 사업자정보 확인
'Date		: 2018.11.07
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>
<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
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

DIM InfoUrl
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

InfoUrl = Trim(request("InfoUrl"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>
<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/PopupTop.asp" -->


	   <div class="area-pop" id="PolicyPopup-pop">
			<div class="full terms">
				<div class="tit-pop">
					<p class="tit" id="popTitle">사업자정보 확인</p>
					<button class="btn-hide-pop" onclick="APP_PopupHistoryBack();">닫기</button>
				</div>

				<div class="container-pop">
					<div class="contents">
						<div class="wrap-agree-changedTerm" id="popCont">
							<iframe id="ComInfo" src="<%=InfoUrl%>" style="display:block; border:0; width:100vw; height: 90vh; overflow:hidden;"></iframe>
						</div>
					</div>
				</div>
			</div>
		</div>


<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/PopupBottom.asp" -->

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
