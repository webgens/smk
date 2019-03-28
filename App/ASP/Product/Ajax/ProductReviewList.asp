<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductReviewList.asp - 상품리뷰 목록 페이지
'Date		: 2019.01.10
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
DIM PhotoFlag
DIM Values											'# 변수값들

DIM ReviewCount

Dim PhotoCount
Dim Photo
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
wQuery	 = "WHERE A.DelFlag = 'N' AND A.ProductCode = " & ProductCode & " "
sQuery	 = "ORDER BY A.BestFlag DESC, A.Idx DESC"

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
	.ActiveConnection = oConn
	.CommandType = adCmdStoredProc
	.CommandText = "USP_Front_EShop_Product_Review_Select"
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
                                        <p style="text-align:center;padding-bottom:20px;">등록된 내역이 없습니다.</p>
<%
ELSE
		Do Until oRs.EOF

			Set oCmd = Server.CreateObject("ADODB.Command")
			WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Product_Review_Image_Select_By_ReviewIdx"
				.Parameters.Append .CreateParameter("@ReviewIdx",	 adInteger,	 adParamInput, ,		 oRs("Idx"))
			END WITH
			oRs1.CursorLocation = adUseClient
			oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
			Set oCmd = Nothing

			PhotoCount = oRs1.RecordCount
			Photo = ""
			If Not oRs1.EOF Then
				Photo = "/Upload/Community/ProductReview/" & oRs1("FileName")
			End If
			oRs1.Close

%>
                                        <div class="reviewitems" style="position: relative;   padding: 15px 0;    border-bottom: 1px solid #e1e1e1;">
                                            <p class="writer"><%=MaskUserID(oRs("UserID"))%></p>
                                            <p class="star-score">
                                                <span class="score"><%=FormatNumber(oRs("AvgGrade"), 1)%></span>
                                                <span class="point val<%=Cint(FormatNumber(oRs("AvgGrade"), 1) * 10)%>"></span>
                                                <!-- 평점에 해당하는 값을 닷(.) 제외하고 val40 같은 형식으로 클래스 부여 (3.5점이면 val35) -->
                                            </p>
                                            <figure class="review-article clearfix">
												<% If Photo <> "" Then %>
                                                <a href="javascript:ReviewImageZoomOpen(<%=oRs("Idx")%>);" class="thumbnail" style="height:100%;">
													<% If PhotoCount > 1 Then %>
													<span class="thumb-more">+<%=PhotoCount%></span><!-- 이미지가 더 있을 경우 표시 -->
													<% End If %>
													<img src="<%=Photo%>" alt="">
												</a>
												<% End If %>

                                                <p class="text"><%=ReplaceDetails(oRs("Contents"))%></p>
                                                <p class="date"><%=DateDiff("d", CDate(Left(oRs("CreateDT"),10)), Date)%>일전</p>
                                            </figure>

                                            <dl class="consumer-score">
                                                <dt>사이즈</dt>
                                                <dd><%=FormatNumber(oRs("SizeGrade"), 1)%></dd>
                                                <dt>착용감</dt>
                                                <dd><%=FormatNumber(oRs("WearGrade"), 1)%></dd>
                                                <dt>디자인</dt>
                                                <dd><%=FormatNumber(oRs("DesignGrade"), 1)%></dd>
                                                <dt>품질</dt>
                                                <dd><%=FormatNumber(oRs("QualityGrade"), 1)%></dd>
                                            </dl>
                                        </div>
<%
				oRs.MoveNext
				i = i + 1
		Loop 
END IF
oRs.Close

Response.Write "|||||" & RecCnt & "|||||" & PageCnt

Set oRs = Nothing
Set oRs1 = Nothing
oConn.Close
Set oConn = Nothing
%>