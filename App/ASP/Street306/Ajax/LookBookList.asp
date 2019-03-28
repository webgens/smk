<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'LookBookList.asp - LookBook리스트
'Date		: 2019.01.09
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
PageCode1 = "00"
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
DIM oRs1						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

DIM Page
DIM PageSize : PageSize = 10
DIM RecCnt
DIM PageCnt

Dim StoreProcName

DIM PCode
Dim ISTopN
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

PCode			 = sqlFilter(Request("PCode"))
Page			 = sqlFilter(Request("Page"))
PageSize		 = sqlFilter(Request("PageSize"))
ISTopN			 = sqlFilter(Request("ISTopN"))

IF Page			 = "" THEN Page			 = 1

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

wQuery = "WHERE A.DelFlag = 'N' "
sQuery = "ORDER BY IDX DESC "

If IsTopN = "Y" Then
	StoreProcName = "USP_Admin_EShop_Street306_LookBook_Select_HistoryBack"
Else
	StoreProcName = "USP_Admin_EShop_Street306_LookBook_Select"
End If
	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = StoreProcName

		.Parameters.Append .CreateParameter("@PAGE",		 adInteger, adParaminput,		, Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger, adParaminput,		, PageSize)
		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

RecCnt	 = oRs(0)

PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

SET oRs = oRs.NextrecordSet

Do While Not oRs.EOF
%>
                        <li class="card">
                            <a href="javascript:LookBookOpen('<%=oRs("IDX")%>');">
                                <img src="<%=oRs("PC_ListImage")%>" alt="<%=oRs("Title1")%>">
                                <div class="txt">
                                    <span><%=oRs("BrandName")%></span>
                                    <strong><%=oRs("Title1")%></strong>
                                    <p><%=oRs("Title2")%></p>
                                </div>
                            </a>
                        </li>
<%
	oRs.MoveNext
Loop
oRs.Close
%>
<%
Response.Write "|||||" & RecCnt & "|||||" & PageCnt

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
