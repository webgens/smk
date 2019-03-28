<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductImageZoom.asp - 상품이미지 확대 페이지
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

DIM Idx
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

Idx = sqlFilter(Request("Idx"))

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Product_Review_Image_Select_By_ReviewIdx"
		.Parameters.Append .CreateParameter("@ReviewIdx",	 adInteger,	 adParamInput, ,		 Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>            
                    <div class="contents wide">
                        <div class="swiper-container review-img">
                            <ul class="swiper-wrapper">
								<%
								Do While Not oRs.EOF	
								%>
                                <li class="swiper-slide">
                                    <div class="img">
                                        <img src="/Upload/Community/ProductReview/<%=oRs("FileName")%>" alt="">
                                    </div>
                                </li>
								<%
									oRs.MoveNext
								Loop	
								%>

                            </ul>

                            <div class="indicator">
                                <div class="swiper-button-next"></div>
                                <div class="swiper-button-prev"></div>
                                <div class="swiper-pagination"></div>
                            </div>
                        </div>
                    </div>

<%
oRs.Close
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>