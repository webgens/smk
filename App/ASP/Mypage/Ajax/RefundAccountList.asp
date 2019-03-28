<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'RefundAccountList.asp - 마이페이지 > 회원정보 수정 > 환불계좌 리스트
'Date		: 2019.01.06
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
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절


Dim IDX
Dim BankCode
Dim BankName
Dim AccountNum
Dim AccountName
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_RefundAccount_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParaminput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

Response.Write "OK|||||"

IF oRs.EOF THEN
%>
								<div class="h-line">
									<h2 class="h-level4">환불금액 계좌 관리</h2>
									<span class="h-date is-right">
										<button type="button" class="button-ty3 ty-bd-black" onclick="refundAccountAdd();">
											<span class="icon ico-add">계좌 추가하기</span>
										</button>
									</span>
								</div>
								<div class="add-account">
									<p>등록된 환불 계좌가 없습니다.</p>
								</div>
<%
ELSE
	IDX				= oRs("IDX")
	BankCode		= oRs("BankCode")
	BankName		= oRs("BankName")
	AccountNum		= oRs("AccountNum")
	AccountName		= oRs("AccountName")
%>
								<div class="h-line">
									<h2 class="h-level4">환불금액 계좌 관리</h2>
								</div>
								<form name="RefundAccountForm" id="RefundAccountForm" method="post">
								<input type="hidden" name="Idx" value="<%=IDX%>" />
								<div class="account-list">
									<div class="refund">
										<p class="bank"><%=BankName%></p>
										<div class="info-wrap">
											<p class="account"><%=LEFT(AccountNum,4) & "*****" & MID(AccountNum,10)%></p>
											<p class="holder">예금주 : <%=LEFT(AccountName,1) & "*" & MID(AccountName,3)%></p>
										</div>
										<div class="buttongroup is-space">
											<button type="button" onclick="refundAccountAdd();" class="button-ty2 is-expand ty-bd-gray">수정</button>
											<button type="button" onclick="refundAccountkDel();" class="button-ty2 is-expand ty-bd-gray">삭제</button>
										</div>
										<!--<div class="right-circle">
											<button type="button" class="closebtn">
												<span class="hidden">닫기</span>
											</button>
										</div>-->
									</div>
								</div>
								</form>
<%
END IF


oRs.Close
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>