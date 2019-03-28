<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'NoticeList.asp - 고객센터 > 공지사항 리스트
'Date		: 2019.01.06
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

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

DIM Page	 : Page = 1
DIM PageSize
DIM RecCnt
DIM PageCnt


DIM IDX
DIM TopFlag
DIM Title
Dim CreateDT
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


PageSize		 = Request("PageSize")
If PageSize		 = "" Then PageSize = 10


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

wQuery = "WHERE A.DelFlag = 'N' "
sQuery = "ORDER BY A.TopFlag DESC, A.CreateDT DESC "
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Notice_Select"

		.Parameters.Append .CreateParameter("@PAGE",		adInteger,	adParamInput,	  ,		Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	adInteger,	adParamInput,	  ,		PageSize)
		.Parameters.Append .CreateParameter("@WQUERY",		adVarchar,	adParamInput, 1000,		wQuery)
		.Parameters.Append .CreateParameter("@SQUERY",		adVarchar,	adParamInput,  100,		sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

RecCnt	 = oRs(0)

SET oRs = oRs.NextrecordSet

Response.Write "OK|||||"
IF oRs.EOF THEN
%>
                            <ul>
                                <li>
                                    <span class="cnt">
	                                    <span class="notice">등록된 공지글이 없습니다.</span>
                                    </span>
                                </li>
                            </ul>
<%
ELSE
%>
                            <ul>
<%
	j = 0
	DO UNTIL oRs.EOF
		IDX				= oRs("IDX")
		TopFlag			= oRs("TopFlag")
		Title			= oRs("Title")
		CreateDT		= oRs("CreateDT")
%>
                                <li onclick="APP_PopupGoUrl('/ASP/Customer/NoticeView.asp?Idx=<%=oRs("Idx")%>');">
                                    <a class="right-arrow-bg">
										<%IF TopFlag = "1" THEN%>
                                        <p class="left-circle important">중요</p>
										<%ELSE%>
											<% If DateDiff("d", CreateDT, Date) <= 15 Then %>
											<p class="left-circle new">NEW</p>
											<% Else %>
		                                    <p class="left-circle"><%=RecCnt-(pagesize*(page-1))-j%></p>
											<% End If %>
										<%END IF%>
                                        <span class="cnt">
                                            <span>공지</span>
                                        <span class="notice"><%=Title%></span>
                                        </span>
                                    </a>
<%
	j = j + 1
	oRs.MoveNext
	LOOP

END IF
%>
                            </ul>
							<form name="NoticeListForm" id="NoticeListForm">
								<input type="hidden" name="Idx" id="Idx" value="<%=IDX%>" />
								<input type="hidden" name="RecCnt" id="RecCnt" value="<%=RecCnt%>" />
								<input type="hidden" name="PageSize" id="PageSize" value="<%=PageSize%>" />
							</form>
<%
oRs.Close
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>