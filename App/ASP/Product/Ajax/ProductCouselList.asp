<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductCounselList.asp - 상품문의 목록 페이지
'Date		: 2018.11.14
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Page											'# 페이지 넘버
DIM PageSize : PageSize = 5							'# 페이지 사이즈
DIM RecCnt											'# 전체 레코드 카운트
DIM PageCnt											'# 페이지 카운트

DIM ProductCode
DIM Values											'# 변수값들

DIM ContentsViewFlag
Dim Reply_Contents
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


ProductCode	= sqlFilter(Request("ProductCode"))

Page			= sqlFilter(Request("Page"))
IF Page			= "" THEN Page	= 1

Values			= ""
Values			= Values & "ProductCode="	& ProductCode



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1 = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
%>					

<%
wQuery	 = "WHERE A.DelFlag = 0 AND A.ProductCode = " & ProductCode & " "

sQuery	 = "ORDER BY A.Idx DESC"


Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
	.ActiveConnection = oConn
	.CommandType = adCmdStoredProc
	.CommandText = "USP_Admin_EShop_Product_QNA_Select"
	.Parameters.Append .CreateParameter("@PAGE",		 adInteger,	 adParamInput, ,		 Page)
	.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger,	 adParamInput, ,		 PageSize)
	.Parameters.Append .CreateParameter("@WQUERY",		 adVarChar,	 adParamInput, 1000,	 wQuery)
	.Parameters.Append .CreateParameter("@SQUERY",		 adVarChar,	 adParamInput, 100,		 sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing


RecCnt	 = oRs(0)
PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)


Set oRs = oRs.NextrecordSet
				
i = 0
IF oRs.EOF THEN
%>
                                            <p style="text-align:center;padding-bottom:20px;padding-top:20px;">등록된 내역이 없습니다.</p>
<%
ELSE
		Do Until oRs.EOF
				IF oRs("SecretFlag") = "0" OR oRs("CreateID") = U_ID THEN
						ContentsViewFlag = "Y"
				ELSE
						ContentsViewFlag = "N"
				END IF
%>
                                            <li>
                                                <div class="area-q-tit">
                                                    <p class="user"><%=MaskUserID(oRs("CreateID"))%></p>
													<%IF oRs("Reply_Flag") = "Y" THEN%>
													<p class="progress">답변완료</p>
													<% Else %>
													<p class="progress ready">답변대기</p>
													<% End If %>
                                                    <p class="tit"><%IF oRs("SecretFlag") = "1" THEN%><span class="ico-lock">비밀글</span><% End If %><%=oRs("Title")%></p>
                                                </div>

												<%IF ContentsViewFlag = "Y" THEN%>
                                                <div class="area-q-cont">
                                                    <p class="cont"><%=ReplaceDetails(oRs("Contents"))%></p>
                                                    <p class="data"><%=LEFT(oRs("CreateDT"),10)%></p>
                                                </div>
												<%IF oRs("Reply_Flag") = "Y" THEN%>
                                                <div class="btn-toggle" id="btn_replyview<%=i%>">
                                                    <a href="javascript:ProductCounsel_Reply('<%=i%>', 'O');">답변보기</a>
                                                </div>

                                                <div class="area-a" id="counsel_reply<%=i%>" style="display:none;">
                                                    <span class="badge">슈마커 답변</span>
                                                    <p class="answer">
                                                        <% If ISNULL(oRs("Reply_Contents")) Then Reply_Contents = "" Else Reply_Contents = oRs("Reply_Contents") End If %>
														<%=ReplaceDetails(Reply_Contents)%>
                                                    </p>
                                                </div>
                                                <div class="btn-toggle" id="btn_replyclose<%=i%>" style="display:none;">
                                                    <a href="javascript:ProductCounsel_Reply('<%=i%>', 'C');">닫기</a>
                                                </div>
												<% End If %>
												<%END IF%>
                                            </li>
<%
				oRs.MoveNext
				i = i + 1
		Loop 
END IF
oRs.Close

Response.Write "|||||" & RecCnt & "개의 문의|||||" & PageCnt

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>