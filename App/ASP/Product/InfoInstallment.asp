<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'InfoInstallment.asp - 무이자 카드혜택 안내
'Date		: 2019.01.08
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

DIM PCImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_InterestInfo_Select"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF oRs.EOF THEN
		PCImage = ""
ELSE
		PCImage = oRs("PCImage")
END IF
oRs.Close
%>
<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/PopupTop.asp" -->

	<section class="wrap-pop" style="display:block">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">무이자 할부 안내</p>
                    <button onclick="APP_PopupHistoryBack()" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents" style="text-align:center">
						<img src="<%=PCImage%>" alt="무이자 할부 안내" style="max-width:100%" />
                    </div>
                </div>
            </div>
        </div>
	</section>

<!-- #include virtual="/INC/FooterNone.asp" -->
<!-- #include virtual="/INC/PopupBottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing	
%>