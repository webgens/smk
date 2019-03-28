<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyInfoModifyPwChk.asp - 마이페이지 > 회원정보 수정(비밀번호 사전확인)
'Date		: 2018.12.18
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
PageCode1 = "05"
PageCode2 = "05"
PageCode3 = "03"
PageCode4 = "01"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

<%
'/****************************************************************************************'
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
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

Response.Write "OK|||||"
%>

        <div class="area-dim" style="z-index:101"></div>

        <div class="area-pop">
            <div class="alert">
				<form name="chkPwdForm1" id="chkPwdForm1" method="post" autocomplete="off">
                <!-- 나의 정보 수정 -->
                <div class="tit-pop">
                    <p class="tit">나의 정보 수정</p>
                </div>
                <div class="container-pop">
					<div class="myinfo-pwchk-info">
						<p class="tit no-border"></p>
						<fieldset>
							<div class="fieldset">
								<label for="join-pw" class="fieldset-label">비밀번호 확인</label>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="password" id="Pwd" name="Pwd" placeholder="사이트 접속 비밀번호 확인" onkeypress="if(event.keyCode == '13') { chk_MyPwd(); return false;}">
									</span>
								</div>
							</div>
						</fieldset>
					</div>
                            
                <!-- 비밀번호 확인 -->
					<div class="btns">
						<button type="button" onclick="chk_MyPwd();" class="button ty-red">확인</button>
						<button type="button" onclick="common_PopClose('DimDepth1');" class="button ty-black">취소</button>
					</div>
				</div>
				</form>
			</div>
		</div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>