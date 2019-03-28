<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ProductList.asp - 상품리스트
'Date		: 2019.01.05
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

DIM SearchType
DIM SCode1
DIM SCode2
DIM SCode3
DIM SBrandCode
DIM SSizeCD
DIM SPrice
DIM EPrice
DIM SColorCode
DIM SPickupFlag
DIM SFreeFlag
DIM SReserveFlag
Dim SSort
Dim ISTopN

Dim ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
SearchType		 = sqlFilter(Request("SearchType"))
SCode1			 = sqlFilter(Request("SCode1"))
SCode2			 = sqlFilter(Request("SCode2"))
SCode3			 = sqlFilter(Request("SCode3"))
SBrandCode		 = sqlFilter(Request("SBrandCode"))
SSizeCD			 = sqlFilter(Request("SSizeCD"))
SPrice			 = sqlFilter(Request("SPrice"))
EPrice			 = sqlFilter(Request("EPrice"))
SColorCode		 = sqlFilter(Request("SColorCode"))
SPickupFlag		 = sqlFilter(Request("SPickupFlag"))
SFreeFlag		 = sqlFilter(Request("SFreeFlag"))
SReserveFlag	 = sqlFilter(Request("SReserveFlag"))
SSort			 = sqlFilter(Request("SSort"))
ISTopN			 = sqlFilter(Request("ISTopN"))

IF Page			 = "" THEN Page			 = 1
IF SearchType	 = "" THEN SearchType	 = "S"
IF SCode1		 = "" THEN SCode1		 = "01"
IF SPrice		 = "" THEN SPrice		 = 0
IF EPrice		 = "" THEN EPrice		 = 100
If SSort		 = "" THEN SSort		 = "1"

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

wQuery = "WHERE A.SaleState = 'Y' "
IF SCode1 <> "" THEN
		wQuery = wQuery & "AND A.ProductCode IN ( "
		wQuery = wQuery & "						SELECT	 DISTINCT ProductCode "
		wQuery = wQuery & "						FROM	 UVW_EShop_Product_Category "
		wQuery = wQuery & "						WHERE	 CategoryCode1 = '" & SCode1 & "' "
		IF SCode2 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode2 = '" & SCode2 & "' "
		END IF
		IF SCode3 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode3 = '" & SCode3 & "' "
		END IF
		wQuery = wQuery & ") "
END IF

'브랜드
If SBrandCode <> "" Then
	wQuery = wQuery & " AND A.BrandCode IN (" & Replace(SBrandCode, "|", "'") & ") "
End If

'사이즈
If SSizeCD <> "" Then
	wQuery = wQuery & " AND (SELECT ISNULL(COUNT(S.StockCnt), 0) FROM UVW_EShop_Product_SizeCD_Available AS S WITH (NOLOCK) WHERE S.ProductCode = A.ProductCode AND S.SizeCD IN (" & Replace(SSizeCD, "|", "'") & ") AND S.StockCnt > 0) > 0 "
End If

'가격대
wQuery = wQuery & " AND SalePrice >= " & SPrice * 10000 & " "
wQuery = wQuery & " AND SalePrice <= " & EPrice * 10000 & " "

'컬러
If SColorCode <> "" Then
	Dim arrSColorCode
	arrSColorCode = Split(Replace(SColorCode, "|", ""), ",")
	wQuery = wQuery & " AND ("
	For i = 0 To UBound(arrSColorCode)
		If Trim(arrSColorCode(i)) <> "" Then
			If i > 0 Then
				wQuery = wQuery & " OR "
			End If
			wQuery = wQuery & " A.ColorCode LIKE '%" & Trim(arrSColorCode(i)) & "%' "
		End If
	Next
	wQuery = wQuery & ")"
End If

'매장픽업
If SPickupFlag = "Y" Then
	wQuery = wQuery & " AND A.PickupFlag = 'Y' "
End If

'배송비무료
If SFreeFlag = "Y" Then
	wQuery = wQuery & " AND C.StandardPrice <= SalePrice "
End If

'예약상품
If SReserveFlag = "Y" Then
	wQuery = wQuery & " AND A.ReserveFlag = 'Y' "
End If

Select CASE SSort
	CASE "1" : sQuery = "ORDER BY A.ProductCode DESC "
	CASE "2" : sQuery = "ORDER BY A.SaleQty DESC "
	CASE "3" : sQuery = "ORDER BY A.DiscountRate DESC "
	CASE "4" : sQuery = "ORDER BY A.SalePrice ASC "
	CASE "5" : sQuery = "ORDER BY A.SalePrice DESC "
End Select

If IsTopN = "Y" Then
	StoreProcName = "USP_Mobile_EShop_Product_Select_HistoryBack"
Else
	StoreProcName = "USP_Mobile_EShop_Product_Select"
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

If oRs.EOF Then
%>
						<li class="no-products" style="width:95%;">
							<p>해당하는 상품이 없습니다.</p>
						</li>
<%
Else
	Do While Not oRs.EOF
		If ISNULL(oRs("ImageUrl")) OR oRs("ImageUrl") = "" Then
			ProductImage = "/images/180_noimage.png"
		Else
			ProductImage = oRs("ImageUrl")
		End If
%>
                        <li>
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems" onclick="pushHash();">
                                <div class="badgegroup">
									<%=ProductBadgeNew(oRs("ProductCode"), oRs("DiscountRate"), oRs("ReserveFlag"), oRs("OPOFlag"), oRs("PickupFlag"), oRs("GiftCnt"), oRs("CouponIdx"))%>
                                </div>
                                <div class="thumbnail"><img src="<%=ProductImage%>" alt="<%=oRs("ProductName")%>"></div>
                                <p class="brand-name"><%=oRs("BrandName")%></p>
                                <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                            </a>
							<a nohref class="listitems">
                                <p class="optional-info">
                                    <button type="button" class="btn-size" onclick="SizeLayerOpen('<%=oRs("ProductCode")%>');">SIZE</button>
                                    <span class="icon ico-fav"><%=FormatNumber(oRs("WishCnt"), 0)%></span>
                                    <span class="icon ico-cmt"><%=FormatNumber(oRs("ReviewCnt"), 0)%></span>
                                </p>
							</a>
                        </li>
<%
		oRs.MoveNext
	Loop
End If
oRs.Close

Response.Write "|||||" & RecCnt & "|||||" & PageCnt

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
