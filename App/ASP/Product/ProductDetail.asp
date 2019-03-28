<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ProductDetail.asp - 상품상세
'Date		: 2018.12.26
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
PageCode1 = "02"
PageCode2 = "01"
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

DIM ProductCode

DIM ProductName
DIM ProductCD
DIM ProdCD
DIM ColorCD
DIM BrandCode
DIM BrandName
Dim BrandNameKor
DIM BrandBGImage
Dim BrandStoryImage
DIM ProductImage
DIM TagPrice
DIM SalePrice
DIM DCRate
DIM EmployeeSalePrice
DIM EmployeeDCRate
DIM MemberPrice
DIM ReserveFlag
DIM PickupFlag
DIM OPOFlag
DIM OffFlag
DIM ShopCD

DIM EmployeeFlag		: EmployeeFlag	= "N"		'# 임직원 여부
DIM EmployeeType		: EmployeeType	= "N"		'# N:비회원, P:일반회원, S:슈마커임직원, J:JD임직원

DIM ShopNM
DIM OutShopFlag
DIM StandardPrice
DIM DeliveryPrice
Dim RZipCode
Dim RAddr1
Dim RAddr2

DIM Material
DIM Color
DIM ProductSize
DIM Manufacturer
DIM Origin
DIM ImportCompany
DIM ASIncharge
DIM SafeQuality
DIM Warranty
DIM Caution
DIM Description
DIM ReentryFlag
DIM CouponTotalExceptFlag

DIM FreebieFlag		: FreebieFlag = "N"		'# 사은품 증정 여부
DIM RelationCount
	
DIM ImageUrl

DIM Picks
DIM tempSizeCD

DIM ReviewCnt
DIM AvgGrade
DIM AvgStar

Dim CategoryName1
Dim CategoryName2
Dim CategoryName3

DIM CouponIdx			 : CouponIdx		 = 0	'# 상품에 적용가능한 쿠폰
DIM CouponApplyPrice	 : CouponApplyPrice	 = 0	'# 상품에 적용한 쿠폰가

Dim	OffCnt				 : OffCnt			 = 0	'# 재고 총 수량
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


ProductCode		 = sqlFilter(Request("ProductCode"))


IF ProductCode = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=상품정보가 없습니다.&Script=APP_HistoryBack();"
		Response.End
END IF



SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1		 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



'# 상품 운영여부 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_For_Available_Check"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF oRs.BOF OR oRs.EOF THEN
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=없는 상품 정보 입니다.&Script=APP_HistoryBack();"
		Response.End
END IF
oRs.Close


IF U_NUM <> "" THEN
		'# 회원정보
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_Member_Select_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   ,		 U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing
																
		IF NOT oRs.EOF THEN
				EmployeeFlag			= oRs("EmployeeFlag")

				IF EmployeeFlag = "Y" THEN
						EmployeeType			= oRs("EmployeeType")
				ELSE
						EmployeeType			= "P"
				END IF
		END IF
		oRs.Close
END IF


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_By_ProductCode"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		ProductName				= oRs("ProductName")
		ProductCD				= oRs("ProductCD")
		ProdCD					= oRs("ProdCD")
		ColorCD					= oRs("ColorCD")
		BrandCode				= oRs("BrandCode")
		BrandName				= oRs("BrandName")
		BrandNameKor			= oRs("BrandNameKor")
		BrandBGImage			= "/Images/img/bg-type-pd-nike.png"				'# 상품상세페이지 백그라운드 이미지
		BrandStoryImage			= oRs("MobileStoryImage")
		ProductImage			= oRs("ProductImage")
		IF ProductImage = "" THEN ProductImage = "/Images/180_noimage.png"
		TagPrice				= oRs("TagPrice")
		SalePrice				= oRs("SalePrice")
		DCRate					= CInt(oRs("DCRate"))
		EmployeeSalePrice		= oRs("EmployeeSalePrice")
		EmployeeDCRate			= oRs("EmployeeDCRate")

		MemberPrice				= CInt(SalePrice * 90 / 100 / 10) * 10			'# 회원혜택가

		ReserveFlag				= oRs("ReserveFlag")
		PickupFlag				= oRs("PickupFlag")
		OPOFlag					= oRs("OPOFlag")
		OffFlag					= oRs("OffFlag")

		ShopCD					= oRs("ShopCD")

		Material				= oRs("Material")	
		Color					= oRs("Color")
		ProductSize				= oRs("ProductSize")
		Manufacturer			= oRs("Manufacturer")
		Origin					= oRs("Origin")
		ImportCompany			= oRs("ImportCompany")
		ASIncharge				= oRs("ASIncharge")
		SafeQuality				= oRs("SafeQuality")
		Warranty				= oRs("Warranty")
		Caution					= oRs("Caution")
		Description				= oRs("Description")
		ReentryFlag				= oRs("ReentryFlag")
		CouponTotalExceptFlag	= oRs("CouponTotalExceptFlag")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=없는 상품 정보 입니다.&Script=APP_HistoryBack();"
		Response.End
END IF
oRs.Close

If ISNULL(Material) Or Material	= "" Then	  Material			   = ""
If ISNULL(Color) Or Color = "" Then	  Color				   = ""
If ISNULL(ProductSize) Or ProductSize	= "" Then	  ProductSize		   = ""
If ISNULL(Manufacturer) Or Manufacturer		= "" Then	  Manufacturer		   = ""
If ISNULL(Origin) Or Origin			= "" Then	  Origin			   = ""
If ISNULL(ImportCompany) Or ImportCompany	= "" Then	  ImportCompany		   = ""
If ISNULL(ASIncharge) Or ASIncharge		= "" Then	  ASIncharge		   = ""
If ISNULL(SafeQuality) Or SafeQuality		= "" Then	  SafeQuality		   = ""
If ISNULL(Warranty) Or Warranty			= "" Then	  Warranty			   = ""
If ISNULL(Caution) Or	Caution			= "" Then	  Caution			   = ""
If ISNULL(Description) Or Description		= "" Then	  Description		   = ""
If ISNULL(ReentryFlag) Or ReentryFlag		= "" Then	  ReentryFlag		   = ""


'# 메인카테고리 정보
wQuery = "WHERE B.ProductCode = " & ProductCode & " AND B.MainFlag = 'Y' "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Category_Select_By_wQuery"

		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If oRs.EOF Then
	CategoryName1 = ""
	CategoryName2 = ""
	CategoryName3 = ""
Else
	CategoryName1 = oRs("CategoryName1")
	CategoryName2 = oRs("CategoryName2")
	CategoryName3 = oRs("CategoryName3")
End If
oRs.Close

'# 판매자 정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Store_Select_By_ShopCD"

		.Parameters.Append .CreateParameter("@ShopCD", adChar, adParamInput, 6, ShopCD)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("OutSaleFlag") <> "Y" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=판매하지 않는 판매자 입니다.&Script=APP_HistoryBack();"
				Response.End
		END IF

		ShopNM			= oRs("ShopNM")
		OutShopFlag		= oRs("OutShopFlag")
		StandardPrice	= oRs("StandardPrice")
		DeliveryPrice	= oRs("DeliveryPrice")
		RZipCode		= oRs("RZipCode")
		RAddr1			= oRs("RAddr1")
		RAddr2			= oRs("RAddr2")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=없는 판매자 정보 입니다.&Script=APP_HistoryBack();"
		Response.End
END IF
oRs.Close


'# 사은품정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_SubProduct_Event_Select_By_ProductCode"

		.Parameters.Append .CreateParameter("@ProductCode",		 adInteger, adParaminput,		, ProductCode)
End WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		FreebieFlag	= "Y"
END IF
oRs.Close



'# 관련상품 개수
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Relation_Select_For_Count"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		RelationCount	= oRs("RelationCount")
ELSE
		RelationCount	= 0
END IF
oRs.Close


'# 예약상품은 1+1, 매장픽업, 관련상품 미적용
IF ReserveFlag = "Y" THEN
		'# 예약상품 1+1은 가능하도록 변경 (2019-03-04 정승영대리 요청)
		'# OPOFlag				= "N"
		PickupFlag			= "N"
		RelationCount		= 0
END IF

'# 1+1상품은 매장픽업 미적용
IF OPOFlag = "Y" THEN
		PickupFlag			= "N"
END IF

Dim MemberNum
MemberNum = U_NUM
If MemberNum = "" Then MemberNum = 0

'# 찜한상품인지
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Pick_Select_By_ProductCode_MemberNum"
		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , MemberNum)
		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
Picks = "N"
If NOT oRs.EOF Then
	Picks = "Y"
End If	
oRs.Close
	
'# 최근 본 상품에 등록
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Latest_Insert"
		.Parameters.Append .CreateParameter("@ProductCode",			adInteger,	adParamInput,     ,	 ProductCode)
		.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 MemberNum)
		.Parameters.Append .CreateParameter("@GuestInfo",			adVarChar,	adParamInput,     255,	 U_GuestInfo)
		.Parameters.Append .CreateParameter("@Location",			adChar,		adParamInput,     1,	 "M")
		.Parameters.Append .CreateParameter("@YYYY",				adVarChar,	adParamInput,     4,	 Left(U_DATE, 4))
		.Parameters.Append .CreateParameter("@MM",					adVarChar,	adParamInput,     2,	 MID(U_DATE, 5, 2))
		.Parameters.Append .CreateParameter("@DD",					adVarChar,	adParamInput,     2,	 Right(U_DATE, 2))
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,     15,	 U_IP)
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

'재고 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_SizeCD_Select_With_EShop_Stock"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		i = 1
		Do Until oRs.EOF
				If oRs("StockCnt") < 1 Then
						tempSizeCD = tempSizeCD & oRs("SizeCD") & ","
				Else
						OffCnt = OffCnt + oRs("StockCnt")
				End If
				oRs.MoveNext
				i = i + 1
		Loop
END IF
oRs.Close

'상품 리뷰 전체 정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Review_Select_For_Product_AvgGrade"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
If Not oRs.EOF Then
	ReviewCnt = oRs("ReviewCnt")
	AvgGrade = FormatNumber(oRs("AvgGrade"), 1)
End If
oRs.Close

AvgStar = FormatNumber(AvgGrade, 1) * 10
If Len(AvgStar) = 2 Then
	If Right(AvgStar, 1) < 5 Then
		AvgStar = FormatNumber(fix(AvgGrade), 0) * 10
	ElseIf Cint(Right(AvgStar, 1)) = 5 Then
		AvgStar = FormatNumber(fix(AvgGrade), 0) * 10 + 5
	Else
		AvgStar = FormatNumber(AvgGrade, 0) * 10
	End If
Else
	AvgStar = "0"
End If

If ISNULL(BrandStoryImage) Or BrandStoryImage = "" Then BrandStoryImage = "/Images/tmp/@bg_detail1.png"



'# 상품 적용 쿠폰 - 일단 로그인 시에만 처리함 / 로그인 안했을 경우 쿠폰 다운로드 안보여줌
IF U_NUM <> "" THEN

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Select_For_PC_At_ProductDetail"
	
				.Parameters.Append .CreateParameter("@ProductCode",	 adInteger,	 adParamInput,  , ProductCode)
				IF U_NUM = "" THEN
				.Parameters.Append .CreateParameter("@LoginFlag",	 adChar,	 adParamInput, 1, "N")
				ELSE
				.Parameters.Append .CreateParameter("@LoginFlag",	 adChar,	 adParamInput, 1, "Y")
				END IF
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				CouponIdx			 = oRs("CouponIdx")
				CouponApplyPrice	 = oRs("ApplyPrice")
		END IF
		oRs.Close

END IF
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.ofh { overflow: hidden !important; }
		.detail-explanation .img-all img { width: 100%; }
		.selected-cont .cont .cost { display: inline-block; margin-left: 11px; font-size: 11px; color: #b4b4b4; }
		.selected-cont .cont .cost .employee { color: #e62019; }
		.selected-cont .cont .oneplusone { display: inline-block; margin-left: 11px; font-size: 11px; color: #e62019; }
		.onePlus-item-list { padding-top: 0; border-top: none; }
		.onePlus-item-list .inform .cont .price { display: block; font-size: 11px; }
		.onePlus-item-list .inform .cont .price>em { font-size: 14px; font-weight: 800; }
	</style>
	<script type="text/javascript" src="/ASP/Product/JS/ProductDetail.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

<%TopSubMenuTitle = ProductName%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <main id="container" class="container">

        <div class="sub_content">
            <section class="wrap-item-detail" >
                <div class="detail-tit">
                    <span class="tit-1"><%=BrandName%></span>
                    <span class="tit-2"><%=ProductName%></span>
                    <span class="tit-3">스타일컬러 : <%=ProdCD%>-<%=ColorCD%></span>
                </div>

                <div class="detail-img-list">
                    <div id="detailImg" class="swiper-container detail-img">
                        <ul class="swiper-wrapper">
							<%
							SET oCmd = Server.CreateObject("ADODB.Command")
							WITH oCmd
									.ActiveConnection	 = oConn
									.CommandType		 = adCmdStoredProc
									.CommandText		 = "USP_Front_EShop_Product_Image_Select_For_Thumbnail_List"

									.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , ProductCode)
									.Parameters.Append .CreateParameter("@SizeClass",	adChar,		adParamInput, 4, "0500")
							END WITH
							oRs.CursorLocation = adUseClient
							oRs.Open oCmd, , adOpenStatic, adLockReadOnly
							SET oCmd = Nothing

							IF NOT oRs.EOF THEN
									i = 1
									Do Until oRs.EOF
							%>
                            <li class="swiper-slide">
                                <img src="<%=oRs("ImageUrl")%>" alt="상품 상세 이미지 <%=i%>">
                            </li>
							<%
											oRs.MoveNext
											i = i + 1
									Loop
							END IF
							oRs.Close
							%>
                        </ul>

                        <div class="swiper-scrollbar"></div>
                    </div>

                    <div class="detail-img-add">
                        <div class="badgegroup">
							<%IF OPOFlag = "Y" THEN%>
                            <span class="badge plusOne">1+1</span>
							<%END IF%>
							<%IF PickupFlag = "Y" THEN%>
                            <span class="badge pickUp">매장<br>픽업</span>
							<%END IF%>
							<%IF FreebieFlag = "Y" THEN%>
                            <span class="badge freebie">사은품</span>
							<%END IF%>
							<%IF CDbl(SalePrice) >= CDbl(StandardPrice) THEN%>
                            <span class="badge pickUp">무료<br>배송</span>
							<%END IF%>
							<%IF CouponTotalExceptFlag  = "Y" THEN%>
                            <span class="badge noCoupon">쿠폰<br>불가</span>
							<%END IF%>
							<%IF CInt(RelationCount) > 0 THEN%>
                            <span class="badge pickUp">관련<br>용품</span>
							<%END IF%>
                        </div>

                        <button type="button" class="btn-zoom" onclick="ProductZoomOpen();">확대</button>
                    </div>
                </div>

                <div class="detail-inform-price">
                    <ul>
                        <li>
                            <span class="tit"><a href="javascript:APP_TopGo()">판매가</a></span>
                            <span class="inform"><em class="ty-price1"><span><%=FormatNumber(SalePrice,0)%></span>원</em><%IF DCRate > 0 THEN%><span class="pre-price"><%=FormatNumber(TagPrice,0)%></span><span class="offPer"><%=FormatNumber(DCRate,0)%>%</span><%END IF%></span>
                        </li>
				<%IF CouponTotalExceptFlag <> "Y" THEN%>
						<%IF U_NUM = "" THEN%>
                        <li>
                            <span class="tit">회원혜택가</span>
                            <span class="inform"><em class="ty-price2"><span><%=FormatNumber(MemberPrice,0)%></span>원</em><button type="button" onclick="APP_GoUrl('/ASP/Member/Join.asp');" class="btn-coupon">회원가입하고 쿠폰받기</button></span>
                        </li>
						<%ELSEIF U_NUM <> "" AND U_MFLAG <> "Y" THEN%>
                        <li>
                            <span class="tit">회원혜택가</span>
                            <span class="inform"><em class="ty-price2"><span><%=FormatNumber(MemberPrice,0)%></span>원</em><button type="button" onclick="APP_TopGoUrl('/ASP/Mypage/');" class="btn-coupon">정회원 전환하고 쿠폰받기</button></span>
                        </li>
						<%ELSE%>
							<%IF CDbl(CouponIdx) > 0 THEN%>
                        <li>
                            <span class="tit">쿠폰적용가</span>
							<span class="inform"><em class="ty-price2"><span><%=FormatNumber(CouponApplyPrice,0)%></span>원</em><button type="button" onclick="coupon_ProductCoupon('<%=CouponIdx%>');" class="btn-coupon">쿠폰</button></span>
                        </li>
							<%END IF%>
						<%END IF%>
				<%END IF%>



						<%IF OPOFlag <> "Y" AND EmployeeFlag = "Y" THEN%>
                        <li>
                            <span class="tit">임직원가</span>
                            <span class="inform"><em class="ty-price1"><span><%=FormatNumber(EmployeeSalePrice,0)%></span>원</em><%IF EmployeeDCRate > 0 THEN%><span class="pre-price"><%=FormatNumber(TagPrice,0)%></span><span class="offPer"><%=FormatNumber(EmployeeDCRate,0)%>%</span><%END IF%></span>
                        </li>
						<%END IF%>
                        <li>
                            <span class="tit">카드혜택</span>
                            <span class="inform"><button type="button" onclick="APP_PopupGoUrl('/ASP/Product/InfoInstallment.asp');" class="underline">무이자할부 안내</button></span>
                        </li>
                        <li>
                            <span class="tit">배송안내</span>
                            <span class="inform"><%=FormatNumber(StandardPrice,0)%>원 이상 무료배송 <span class="verline">개별배송 가능</span></span>
                        </li>
                    </ul>
					 <% If ReserveFlag <> "Y" AND Trim(tempSizeCD) <> "" AND ReentryFlag = "Y" Then %>
                    <button type="button" class="button" onclick="<% If U_Num = "" Then %>LoginChk();<% ELSEIf U_Num <> "" AND U_MFLAG <> "Y" Then %>regularLogin();<% Else %>Reentry_Open();<% End If %>">
						<span class="icon ico-bell">품절 사이즈 재입고 신청</span>
					</button>
					<% End If %>
                </div>

				<%
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Product_Select_For_OtherColor"

						.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , ProductCode)
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs.EOF THEN
						IF oRs.RecordCount > 1 THEN
				%>

                <h2 class="t-level4">컬러 <span class="color-length"><%=oRs.RecordCount%>개의 색상</span></h2>
				<style>
					.detail-inform-color ul li { margin-right: 0 !important; }
				</style>
                <div class="detail-inform-color" style="padding: 0 10px;">
                    <div class="swiper-container color-list">
                        <ul class="swiper-wrapper">
				<%
								i = 1
								Do Until oRs.EOF
				%>
                            <li class="swiper-slide">
                                <img src="<%=oRs("ImageUrl_0180")%>" alt="컬러 이미지 <%=i%>" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" />
                                <span class="point<%IF CStr(oRs("ProductCode")) = CStr(ProductCode) THEN%> had<%END IF%>"></span>
                            </li>
				<%
										oRs.MoveNext
										i = i + 1
								Loop
				%>
                        </ul>
                    </div>
                </div>
				<%
						END IF
				END IF
				oRs.Close
				%>

                <div class="detail-review">
					<!--
                    <div class="bg-img">
                        <img src="<%=BrandStoryImage%>" alt="<%=ProductName%>">
                    </div>
					//-->

                    <div class="satisfy">
                        <!-- 만족도 부분 배경색 넣는 곳 -->
                        <div class="average-score">
                            <h1>평균 고객 만족도</h1>
                            <p class="star-score sz-large">
                                <span class="point val<%=AvgStar%>"></span>
                                <!-- 3.5점 -->
                                <!-- 평점에 해당하는 값을 닷(.) 제외하고 val40 같은 형식으로 클래스 부여 (3.5점이면 val35) -->
                            </p>
                            <p class="score">
                                <%=FormatNumber(AvgGrade, 1)%>점<span>(<%=ReviewCnt%>개의 후기)</span>
                            </p>
                        </div>
                    </div>

                    <div class="sect">
                        <!-- 아코디언(상품후기, 상품정보 등) 배경색 넣는 곳 -->
                        <ul>
							<input type="hidden" id="ReviewPage" />
							<input type="hidden" id="CounselPage" />
                            <li id="detailAccord_1" class="review">
                                <div class="accordion-selector">
                                    <span class="tit">상품후기</span><span class="cont"><%=ReviewCnt%>개의 후기</span>
                                    <button type="button" class="btn-accord-ty1" data-target="detailAccord_1" onclick="ReviewList(1, '<%=ProductCode%>');" >열기</button>
                                </div>
                                <div class="accordion-panel">
                                    <div class="reviewlist" id="reviewList">
                                    </div>

                                    <div class="buttongroup">
                                        <button type="button" class="button more" id="review_morebtn" onclick="NextReviewList('<%=ProductCode%>');">
											<span class="icon is-right is-arrow-d1">후기 더보기</span>
										</button>
                                    </div>
                                </div>
                            </li>
                            <li id="detailAccord_2" class="inform">
                                <div class="accordion-selector">
                                    <span class="tit">상품정보</span><span class="cont">사이즈 / 소재 / 취급 시 주의사항</span>
                                    <button type="button" class="btn-accord-ty1" data-target="detailAccord_2">열기</button>
                                </div>
                                <div class="accordion-panel">
                                    <div class="informItems">
                                        <ul>
                                            <li>
                                                <p class="tit">사이즈</p>
                                                <p class="cont"><%=Replace(ProductSize, "|", ", ")%> &nbsp; <button type="button" class="size-chart" onclick="APP_PopupGoUrl('/ASP/Product/ProductSizeChart.asp');">SIZE CHART</button></p>
                                            </li>
                                            <li>
                                                <p class="tit">소재</p>
                                                <p class="cont"><%=Material%>&nbsp;</p>
                                            </li>
                                            <li>
                                                <p class="tit">제조자</p>
                                                <p class="cont"><%=Manufacturer%>&nbsp;</p>
                                            </li>
                                            <li>
                                                <p class="tit">제조국</p>
                                                <p class="cont"><%=Origin%>&nbsp;</p>
                                            </li>
                                            <li>
                                                <p class="tit">안전품질표시</p>
                                                <p class="cont">
                                                    <%=SafeQuality%>
                                                    <br>
                                                    <img src="/Images/ico/ico_kc.png" alt="안전품질표시 로고" class="ico_logo_kc">
                                                </p>
                                            </li>
                                            <li>
                                                <p class="tit">품질보증기준</p>
                                                <p class="cont">
                                                    <%=Warranty%>&nbsp;
                                                </p>
                                            </li>
                                            <li>
                                                <p class="tit">주의사항</p>
                                                <p class="cont">
                                                    <%=Caution%>&nbsp;
                                                </p>
                                            </li>
                                        </ul>
                                    </div>
                                </div>
                            </li>
                            <li id="detailAccord_3" class="question">
                                <div class="accordion-selector">
                                    <span class="tit">상품문의</span><span class="cont" id="ProductCounselCount">0개의 문의</span>
                                    <button type="button" class="btn-accord-ty1" data-target="detailAccord_3">열기</button>
                                </div>
                                <div class="accordion-panel">
                                    <div class="buttongroup" style="padding-right:23px;">
                                        <button type="button" class="button is-expand" style="line-height:15px;" onclick="<% If U_Num = "" Then %>conf_Login();<% Else %>popup_ProductQna('<%=ProductCode%>');<% End If %>">
											<span class="icon ico-question">문의 하기</span>
										</button>
                                    </div>

                                    <div class="qnaList">
                                        <ul id="productcounselList">
                                        </ul>
                                    </div>

                                    <div class="buttongroup">
                                        <button type="button" class="button more" id="productcounsel_more_btn" onclick="list_ProductCounselNext('<%=ProductCode%>');">
											<span class="icon is-right is-arrow-d1">문의 더보기</span>
										</button>
                                    </div>
                                </div>
                            </li>
                            <li id="detailAccord_4" class="orderShipping">
                                <div class="accordion-selector">
                                    <span class="tit">주문/배송</span><span class="cont"><%=FormatNumber(StandardPrice, 0)%>원 이상 무료배송 / 배송 안내</span>
                                    <button type="button" class="btn-accord-ty1" data-target="detailAccord_4">열기</button>
                                </div>
                                <div class="accordion-panel">
                                    <div class="informItems">
                                        <ul>
                                            <li>
                                                <p class="tit">당일출고</p>
                                                <p class="cont">오후 12시까지 결제건까지 해당되며, 그 이후에는 익일발송됩니다.</p>
                                            </li>
                                            <li>
                                                <p class="tit">배송소요일</p>
                                                <p class="cont">2~3일(주말/공휴일 제외)</p>
                                            </li>
                                            <li>
                                                <p class="tit">배송비</p>
                                                <p class="cont"><%=FormatNumber(StandardPrice, 0)%>원 이상 무료배송</p>
                                            </li>
                                            <li>
                                                <p class="tit">무통장입금</p>
                                                <p class="cont">3일 이내 입금이 완료되지 않을 시 주문이 자동취소 됩니다.</p>
                                            </li>
                                            <li>
                                                <p class="tit">배송업체</p>
                                                <p class="cont">대한통운 1588-1255</p>
                                            </li>
                                        </ul>
                                    </div>
                                    <div class="area-ps">
                                        <p>상품오염 및 불량 등의 검수작업이 지연될 시 발송지연될 수 있습니다. (최대 7일)</p>
                                    </div>
                                </div>
                            </li>
                            <li id="detailAccord_5" class="reActive">
                                <div class="accordion-selector">
                                    <span class="tit">교환/반품/AS</span><span class="cont">접수 안내</span>
                                    <button type="button" class="btn-accord-ty1" data-target="detailAccord_5">열기</button>
                                </div>
                                <div class="accordion-panel">
                                    <div id="tabs" class="tab" data-use="">
                                        <ul class="tab-selector">
                                            <li class="part-4"><a href="javascript:;" data-target="tabs-col1">교환/반품</a></li>
                                            <li class="part-4"><a href="javascript:;" data-target="tabs-col2">A/S</a></li>
                                            <li class="part-4"><a href="javascript:;" data-target="tabs-col3">심의</a></li>
                                            <li class="part-4"><a href="javascript:;" data-target="tabs-col4">고객센터</a></li>
                                        </ul>
                                        <div id="tabs-col1" class="tab-panel">
                                            <div class="informItems">
                                                <ul>
                                                    <li>
                                                        <p class="tit">반품배송지</p>
                                                        <address class="cont">
															(<%=RZipCode%>) <%=RAddr1 & " " & RAddr2%>
														</address>
                                                    </li>
                                                    <li>
                                                        <p class="tit">접수가능일</p>
                                                        <p class="cont">상품수령일로부터 7일 이내 접수</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">접수방법</p>
                                                        <p class="cont">마이페이지 > 주문/배송 또는 고객센터에서 신청<br>(고객센터: 080-030-2809)</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">배송비</p>
                                                        <p class="cont">
                                                            색상/사이즈 교환, 단순변심 : 고객님 부담<br> 
															제품불량, 오배송 : 회사측 부담<br> 
															대한통운 이용 : 왕복배송비 <%=FormatNumber(DeliveryPrice * 2, 0)%>원 동봉하여 발송<br> 
															타 택배 이용 : 택배비 결제 후 편도배송비 <%=FormatNumber(DeliveryPrice, 0)%>원<br> 
															동봉하여 발송<br />
															※ 합주문 (2개 이상의 상품을 장바구니에 담아 2개 동시결제) 후 1개 반품시 택배비 왕복 <%=FormatNumber(DeliveryPrice * 2, 0)%>원 발생
                                                        </p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">유의사항</p>
                                                        <p class="cont">
                                                            신속한 처리를 위해 동봉된 신청서에 올바른 고객 정보를 입력해 주세요.<br> 단, 온/오프라인 간 매장교환이나 반품은 불가합니다.
                                                        </p>
                                                    </li>
                                                </ul>
                                            </div>

                                            <div class="area-ps">
                                                <p>교환/반품 접수시 1:1 게시판 또는 고객센터로 연락주시면 보다 빠른 처리 도와드립니다.</p>
                                                <p>상품/택/박스/사은품 분실 및 훼손이나 외부착용 했을시, 제품수령 후 7일이 경과된 경우 환불이 불가능합니다.</p>
                                            </div>
                                        </div>
                                        <div id="tabs-col2" class="tab-panel">
                                            <div class="informItems">
                                                <ul>
                                                    <li>
                                                        <p class="tit">반품배송지</p>
                                                        <address class="cont">
															(<%=RZipCode%>) <%=RAddr1 & " " & RAddr2%>
														</address>
                                                    </li>
                                                    <li>
                                                        <p class="tit">AS 접수</p>
                                                        <p class="cont">상품주문일로부터 6개월 이내 접수</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">AS 기간</p>
                                                        <p class="cont">총 2주정도 소요되며, 완료시 개별연락 드립니다.</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">AS 비용</p>
                                                        <p class="cont">접착 및 봉제수선은 5,000원 정도이며, 내용에 따라 금액이 변동될 수 있습니다.</p>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                        <div id="tabs-col3" class="tab-panel">
                                            <div class="informItems">
                                                <ul>
                                                    <li>
                                                        <p class="tit">반품배송지</p>
                                                        <address class="cont">
															(<%=RZipCode%>) <%=RAddr1 & " " & RAddr2%>
														</address>
                                                    </li>
                                                    <li>
                                                        <p class="tit">심의접수</p>
                                                        <p class="cont">상품주문일로부터 1년 이내 접수</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">심의기간</p>
                                                        <p class="cont">약 한달정도 소요되며, 내용에 따라 기간연장될 수 있습니다.</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">심의방법</p>
                                                        <p class="cont">공정하고 정확한 제품불량 확인을 위해<br> 1차 : 브랜드 공식심의<br> 2차 : 한국소비생활연구원(사단법인)<br> 에 접수진행하고 있습니다.</p>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="area-ps">
                                                <p>단, 온라인 결제확인에 한하며 오프라인 구매는 해당 매장으로 요청바랍니다.</p>
                                            </div>
                                        </div>
                                        <div id="tabs-col4" class="tab-panel">
                                            <div class="informItems">
                                                <ul>
                                                    <li>
                                                        <p class="tit">상담시간</p>
                                                        <p class="cont">오전 10시 - 오후 5시 (주말/공휴일 제외)</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">점심시간</p>
                                                        <p class="cont">오후 12시 - 오후 1시</p>
                                                    </li>
                                                    <li>
                                                        <p class="tit">온라인 문의</p>
                                                        <p class="cont">고객센터 > 1:1문의 게시판을 이용하시면 더욱 빠른 답변을 받아보실 수 있습니다.<br>
                                                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/Qna.asp?QnaType=2');" class="mtm">1:1문의 바로가기</a></p>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </li>
                        </ul>
                    </div>
                </div>

				<%IF BrandCode = "NK" THEN%>
				<p><img src="/images/img/Nike_Dealer_Banner.jpg" style="width:100%"></p>
				<%END IF%>

                <div class="detail-explanation">
                    <div class="img-all">
						<%
						'# 사은품 증정 배너
						IF FreebieFlag	= "Y" THEN
								SET oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection	 = oConn
										.CommandType		 = adCmdStoredProc
										.CommandText		 = "USP_Front_EShop_SubProduct_Event_Select_For_Banner_By_ProductCode"

										.Parameters.Append .CreateParameter("@ProductCode",		 adInteger, adParaminput,		, ProductCode)
								End WITH
								oRs.CursorLocation = adUseClient
								oRs.Open oCmd, , adOpenStatic, adLockReadOnly
								SET oCmd = Nothing

								IF NOT oRs.EOF THEN
										Do Until oRs.EOF
						%>
						<img src="<%=oRs("MBanner")%>" alt="" />
						<%
												oRs.MoveNext
										Loop
								END IF
								oRs.Close
						END IF
						%>

						<%
						If InStr(LCase(Request.ServerVariables("HTTP_USER_AGENT")), "iphone") Or InStr(LCase(Request.ServerVariables("HTTP_USER_AGENT")), "ipad") Then
							If ProductCode = "10872" OR ProductCode = "10851" Then
							%>
							<iframe src="//contents.vrism.net/data/content/19FS5NI008/" allowfullscreen="" width="100%" height="600" frameborder="0" style="border:0" mozallowfullscreen="true" webkitallowfullscreen="true"></iframe>
						<%
							End If
						End If
						%>

						<%
						If InStr(LCase(Request.ServerVariables("HTTP_USER_AGENT")), "iphone") Or InStr(LCase(Request.ServerVariables("HTTP_USER_AGENT")), "ipad") Then
							If ProductCode = "10881" OR ProductCode = "10856" Then
							%>
							<iframe src="//contents.vrism.net/data/content/19FS5NI004/" allowfullscreen="" width="100%" height="600" frameborder="0" style="border:0" mozallowfullscreen="true" webkitallowfullscreen="true"></iframe>
						<%
							End If
						End If
						%>

                        <%'=Replace(Description, "/_DATA_/", "http://www.shoemarker.co.kr/_DATA_/")%>
                        <%=Replace(Description, "http://www.shoemarker.co.kr", "")%>
                    </div>
                </div>

<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_MDRecommend_Select_By_ProductCode"

		.Parameters.Append .CreateParameter("@ProductCode",		 adInteger, adParaminput,		, ProductCode)
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParaminput,		, 0)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then
%>
                <h2 class="detail-sub-tit">MD 추천상품</h2>

                <div class="detail-lately-item">
                    <div class="swiper-container item-group">
                        <ul class="swiper-wrapper">
						<%
						Do While Not oRs.EOF
						%>
                            <li class="swiper-slide">
                                <img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>');">
                            </li>
						<%
							oRs.MoveNext
						Loop	
						%>
                        </ul>
                    </div>
                </div>
<%
End If
oRs.Close	
%>

<%
DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN

wQuery = "WHERE BCode = '11' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MainBanner_Select_Top3_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF NOT oRs.EOF THEN
%>
                <h2 class="detail-sub-tit">관련 기획전</h2>
                <div class="ly-more-event">
                    <div class="swiper-container more-event">
                        <div class="swiper-wrapper">
<%
		Do Until oRs.EOF	
%>
                            <div class="swiper-slide">
                                <a href="javascript:void(0)" onclick="LinkgoUrl('<%=oRs("LinkUrl")%>')" class="listitems">
                                    <div class="thumbnail">
										<img src="<%=oRs("MobileImage1")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
                                    </div>

                                    <div class="inform">
                                        <span class="tit"><%=oRs("SubTitle1")%></span>
                                        <span class="date"><%=GetDateYMD2(LEFT(oRs("StartDT"), 8))%> ~ <%=GetDateYMD2(LEFT(oRs("EndDT"), 8))%></span>
                                    </div>
                                </a>
                            </div>
<%
				oRs.MoveNext
		Loop
%>
                        </div>
                        <!-- Add Scrollbar -->
                        <div class="area-scrollbar">
                            <div class="swiper-scrollbar"></div>

                        </div>
                    </div>
                </div>
<%
END IF
oRs.Close
%>

                <div class="brandShop-banner">
                    <div class="cont">
                        <span class="brand-name-en"><%=BrandName%></span>
                        <span class="brand-name-ko"><%=BrandNameKor%></span>
                    </div>

                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=BrandCode%>')" class="gotoThe">브랜드샵 바로가기</a>
                </div>
            </section>
        </div>
    </main>

	<form name="OrderForm" method="post">
		<input type="hidden" name="OrderType"		 value="" />
		<input type="hidden" name="DelvType"		 value="" />
		<input type="hidden" name="ProductCode"		 value="" />
		<input type="hidden" name="SizeCD"			 value="" />
		<input type="hidden" name="OrderCnt"		 value="" />
		<input type="hidden" name="SalePriceType"	 value="" />
		<input type="hidden" name="ProductType"		 value="" />
	</form>

<!-- #include virtual="/INC/Footer.asp" -->




	<input type="hidden" name="OPOFlag"			 id="OPOFlag"		 value="<%=OPOFlag%>"		 />
	<input type="hidden" name="EmployeeFlag"	 id="EmployeeFlag"	 value="<%=EmployeeFlag%>"	 />
	<input type="hidden" name="EmployeeType"	 id="EmployeeType"	 value="<%=EmployeeType%>"	 />
	<input type="hidden" name="ProductSeq"		 id="ProductSeq"	 value="0"					 />

	<input type="hidden" name="RelationCount"	 id="RelationCount"	 value="<%=RelationCount%>"	 />
	<input type="hidden" name="ReserveFlag"		 id="ReserveFlag"	 value="<%=ReserveFlag%>"	 />
	<input type="hidden" name="PickupFlag"		 id="PickupFlag"	 value="<%=PickupFlag%>"	 />




    <!-- PopUp -->
    <section class="wrap-pop" id="PopPurchase" style="display:block;height:0px;">
        <div class="area-dim" style="display:none"></div>



        <div class="area-pop">
            <div class="area-select-option">
                <!-- 옵션 선택  -->
                <div class="select-option">
                    <div class="footSize-table">
                        <div id="footSize_all" class="accordion">
                            <div class="selector">
                                <button type="button" class="btn-select clickEvt" data-target="footSize_all">
									<span>사이즈 선택</span>
								</button>
                            </div>
                            <div class="option">
                                <div class="pop-size">
									<%
									SET oCmd = Server.CreateObject("ADODB.Command")
									WITH oCmd
											.ActiveConnection	 = oConn
											.CommandType		 = adCmdStoredProc
											.CommandText		 = "USP_Front_EShop_Product_SizeCD_Select_With_EShop_Stock"

											.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , ProductCode)
									END WITH
									oRs.CursorLocation = adUseClient
									oRs.Open oCmd, , adOpenStatic, adLockReadOnly
									SET oCmd = Nothing

									IF NOT oRs.EOF THEN
											i = 1
											Do Until oRs.EOF
													If oRs("StockCnt") > 0 Then
									%>
                                    <span class="check-style"><input type="radio" name="chk-size" id="chk-size<%=i%>" value="<%=oRs("SizeCD")%>" <%IF oRs("StockCnt") > 0 THEN%>onclick="selectProduct('<%=ProductCode%>','<%=oRs("SizeCD")%>')"<%ELSE%>disabled<%END IF%> /><label for="chk-size<%=i%>"><span><%=oRs("SizeCD")%></span></label></span>
									<%
													End If

													oRs.MoveNext
													i = i + 1
											Loop
									END IF
									oRs.Close
									%>
                                </div>
                            </div>
                        </div>
                    </div>
				<%
				IF CInt(RelationCount) > 0 THEN
				%>
                    <div class="related-item">
                        <div id="relatedItem_all" class="accordion">
                            <div class="selector">
                                <button type="button" onclick="openRelation('<%=ProductCode%>')" class="btn-select">
									<span>관련용품 선택</span>
								</button>
                            </div>
                        </div>
                    </div>
				<%
				END IF
				%>
                </div>
                <!-- // 옵션 선택  -->

                <!-- 선택한 상품 -->
                <div class="selected-item">
                    <ul id="SelectProductList">
                    </ul>
                </div>
                <!-- // 선택한 상품 -->

                <!-- 총 금액 -->
                <div class="selected-bill">
                    <div class="total-price">
                        <span class="total">총 금액</span>
                        <span class="price"><em id="TotalPrice">0</em>원</span>
                    </div>

                    <div class="buttongroup purchase">
					<%IF OffFlag = "Y" OR OffCnt < 1 THEN%>
                        <button type="button" onclick="void(0)" class="button ty-gray wty6">품 절</button>
					<%ELSEIF NAVER_PAY_FLAG = "Y" THEN%>
						<%IF ReserveFlag = "Y" THEN%>
						<button type="button" onclick="addOrderSheet('R', 'P', 'C')" class="button ty-red wty5">예약상품 구매하기</button>
                        <button type="button" onclick="addOrderSheet('R', 'P', 'N')" class="button ty-n-pay wty1">N Pay</button>
						<%ELSEIF PickupFlag = "Y" THEN%>
                        <button type="button" onclick="addCart()" class="button ic-basket wty1">장바구니</button>
                        <button type="button" onclick="addOrderSheet('G', 'P', 'C')" class="button ty-red wty2">구매하기</button>
                        <button type="button" onclick="addOrderSheet('G', 'S', 'C')" class="button ty-picked wty2">매장픽업</button>
                        <button type="button" onclick="addOrderSheet('G', 'P', 'N')" class="button ty-n-pay wty1">N Pay</button>
						<%ELSE%>
                        <button type="button" onclick="addCart()" class="button ic-basket wty1">장바구니</button>
                        <button type="button" onclick="addOrderSheet('G', 'P', 'C')" class="button ty-red wty4">구매하기</button>
                        <button type="button" onclick="addOrderSheet('G', 'P', 'N')" class="button ty-n-pay wty1">N Pay</button>
						<%END IF%>
					<%ELSE%>
						<%IF ReserveFlag = "Y" THEN%>
						<button type="button" onclick="addOrderSheet('R', 'P', 'C')" class="button ty-red wty6">예약상품 구매하기</button>
						<%ELSEIF PickupFlag = "Y" THEN%>
                        <button type="button" onclick="addCart()" class="button ic-basket wty1">장바구니</button>
                        <button type="button" onclick="addOrderSheet('G', 'P', 'C')" class="button ty-red wty3">구매하기</button>
                        <button type="button" onclick="addOrderSheet('G', 'S', 'C')" class="button ty-picked wty3">매장픽업</button>
						<%ELSE%>
						<button type="button" onclick="addCart()" class="button ic-basket wty1">장바구니</button>
                        <button type="button" onclick="addOrderSheet('G', 'P', 'C')" class="button ty-red wty5">구매하기</button>
						<%END IF%>
					<%END IF%>
                    </div>
                </div>
                <!-- // 총 금액 -->
                <button type="button" class="btn-hide-select">닫기</button>
            </div>
        </div>
    </section>

    <section class="wrap-pop" id="ProductZoom">
        <div class="area-dim"></div>

        <div class="area-pop">
			<div class="full">
                <div class="tit-pop">
                    <button type="button" class="btn-hide-pop" onclick="ProductZoomClose()">닫기</button>
                </div>

                <div class="container-pop" id="slidecont">


					
                </div>
            </div>

        </div>
    </section>

    <section class="wrap-pop" id="ReviewImage">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full reviewImg">
                <button class="btn-hide-pop" onclick="ReviewImageZoomClose();">닫기</button>

                <div class="container-pop" id="reviewImage-pop">

                </div>
            </div>
        </div>
    </section>


    <section class="wrap-pop" id="ShareView">
        <div class="area-dim"></div>

        <div class="area-pop">
            <!-- 팝업 공유하기 -->
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">공유하기</p>
                    <button class="btn-hide-pop" type="button" onclick="ShareClose();">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents" style="text-align:center;">
                        <div class="wrap-share">
                            <a href="javascript:facebook_share();ShareClose();" class="facebook"><img src="/images/ico/ico_share_fac.png" alt="페이스북"></a>
                            <a href="javascript:instagram_share();ShareClose();" class="instagram"><img src="/images/ico/ico_share_ins.png" alt="인스타그램"></a>
                            <a href="javascript:kakao_share('ProductCode=<%=ProductCode%>', '<%=ProductName%>', 'http://m.shoemarker.co.kr<%=ProductImage%>', '슈마커', '<%=TagPrice%>', '<%=SalePrice%>', '<%=DCRate%>', '상품보러가기');ShareClose();" class="kakao"><img src="/images/ico/ico_share_kakao.png" alt="카카오톡"></a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>


    <!-- bnb-ty2 -->
    <article class="bnb-ty2" style="background-color: #ff201b;">
        <button type="button" class="btn-pick <%IF Picks = "Y" THEN%> on<%END IF%>" id="Pickup" onclick="PickCheck('<%=ProductCode%>', '<%=Picks%>');">찜하기</button>
        <button type="button" class="btn-share" onclick="ShareOpen();">공유하기</button>
        <button type="button" class="btn-buy"><%IF OffFlag="Y" Or OffCnt<1 THEN%>품 절<%ELSE%>구매하기<%END IF%></button>
    </article>
    <!-- bnb-ty2 -->

	<script type="text/javascript">
		$(function () {
			$('.bnb-ty2 .btn-buy').on('click', function () {
				$("#PopPurchase .area-dim").show();
				$("body").addClass("ofh");
				$('.area-select-option').addClass('is-block');
			});
			$('.area-select-option .btn-hide-select').on('click', function () {
				$('.area-select-option').removeClass('is-block');
				$("#PopPurchase .area-dim").hide();
				$("body").removeClass("ofh");
			});
		});

		function init_PurchaseBox() {
			$("#SelectProductList").html("");
			$("input:radio[name='chk-size']").prop("checked", false);
			$(".area-select-option .option").hide();
			$(".area-select-option .selector").removeClass('is-focus');
			$('.area-select-option').removeClass('is-block');
			$("#PopPurchase .area-dim").hide();
			$("body").removeClass("ofh");
		}



		/* 상품선택 */
		function selectProduct(productCode, sizeCD) {
			if (productCode.length == 0) {
				openAlertLayer("alert", "선택된 스타일컬러가 없습니다.", "closePop('alertPop', '');", "");
				return;
			}
			if (sizeCD.length == 0) {
				openAlertLayer("alert", "사이즈를 선택해 주십시오.", "closePop('alertPop', '');", "");
				return;
			}

			var opoFlag		 = $("#OPOFlag").val();
			var employeeFlag = $("#EmployeeFlag").val();
			var employeeType = $("#EmployeeType").val();
			
			if (opoFlag == "Y") {
				// 1+1 선택 레이어 열기
				$.ajax({
					type		 : "post",
					url			 : "/ASP/Product/Ajax/ProductOnePlusOneList.asp",
					async		 : false,
					data		 : "ProductCode=" + productCode + "&SizeCD=" + sizeCD,
					dataType	 : "text",
					success		 : function (data) {
									var splitData = data.split("|||||");
									var result = splitData[0];
									var cont = splitData[1];

									if (result == "OK") {
										$("#DimDepth1").html(cont);
										openPop('DimDepth1');
									}
									else if (result == "LOGIN") {
										PageReload();
									}
									else {
										openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
										return;
									}
					},
					error		 : function (data) {
									openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
					}
				});

			}
			else if (employeeFlag == "Y") {
				// 임직원가 구매여부 묻기
				openAlertLayer2("confirm", "임직원가로 구매 하시겠습니까?", "임직원가 구매", "일반가 구매", "closePop('confirmPop', '');addProduct('2','" + productCode + "','" + sizeCD + "','','');", "closePop('confirmPop', '');addProduct('1','" + productCode + "','" + sizeCD + "','','');");
			}
			else {
				// 하단 구매레이어에 상품 담기
				addProduct("1", productCode, sizeCD, "", "");
			}
		}

		/* 1+1 상품선택 */
		function selectOnePlusOne(productCode, sizeCD) {
			if ($("#ProductOnePlusOneLayer input[name='OProductCode']:checked").length == 0) {
				openAlertLayer("alert", "1+1 상품을 선택해 주십시오.", "closePop('alertPop', '');", "");
				return;
			}

			var oProductCode	 = $("#ProductOnePlusOneLayer input[name='OProductCode']:checked").val();
			var num				 = $("#ProductOnePlusOneLayer input[name='OProductCode']:checked").data("num");

			if ($("#ProductOnePlusOneLayer input[name='OSizeCD" + num + "']:checked").length == 0) {
				openAlertLayer("alert", "선택한 1+1상품의 사이즈를 선택해 주십시오.", "closePop('alertPop', '');", "");
				$("#ProductOnePlusOneLayer input[name='OSizeCD" + num + "']").focus();
				return;
			}
			var oSizeCD = $("#ProductOnePlusOneLayer input[name='OSizeCD" + num + "']:checked").val();

			// 하단 구매레이어에 상품 담기
			addProduct("1", productCode, sizeCD, oProductCode, oSizeCD);

			closePop('DimDepth1');
		}

		/* 관련상품 선택 레이어 열기 */
		function openRelation(productCode) {
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Product/Ajax/ProductRelationList.asp",
				async		 : false,
				data		 : "ProductCode=" + productCode,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									$("#DimDepth1").html(cont);
									openPop('DimDepth1');
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 관련상품 선택 */
		function selectRelation() {
			if ($("#ProductRelationLayer input[name='RProductCode']:checked").length == 0) {
				openAlertLayer("alert", "1111관련용품을 선택해 주십시오.", "closePop('alertPop', '');", "");
				return;
			}

			var rProductCode	 = $("#ProductRelationLayer input[name='RProductCode']:checked").val();
			var num				 = $("#ProductRelationLayer input[name='RProductCode']:checked").data("num");

			if ($("#ProductRelationLayer input[name='RSizeCD" + num + "']:checked").length == 0) {
				openAlertLayer("alert", "선택한 관련용품의 사이즈를 선택해 주십시오.", "closePop('alertPop', '');", "");
				$("#ProductRelationLayer input[name='RSizeCD" + num + "']").focus();
				return;
			}
			var rSizeCD = $("#ProductRelationLayer input[name='RSizeCD" + num + "']:checked").val();

			// 선택중복 체크
			var errChk = "N";
			$("#SelectProductList li").each(function () {
				var sProductCode = $(this).find("input[name='ProductCode']").val();
				var sSizeCD = $(this).find("input[name='SizeCD']").val();

				if (rProductCode == sProductCode && rSizeCD == sSizeCD) {
					common_msgPopOpen("", productName + "의 [" + sizeCD + "] 사이즈는 이미 선택하셨습니다.", "", "msgPopup", "N");
					errChk = "Y";
					return false;
				}
			});
			if (errChk == "Y") {
				return;
			}

			// 하단 구매레이어에 상품 담기
			addProduct("1", rProductCode, rSizeCD, "", "");

			closePop('DimDepth1');
		}

		/* 선택상품 하단 구매레이어 추가 */
		function addProduct(salePriceType, productCode, sizeCD, oProductCode, oSizeCD) {
			var seq = Number($("#ProductSeq").val()) + 1;

			if (oProductCode == "") { $("#ProductSeq").val(seq); }
			else					{ $("#ProductSeq").val(seq + 1); }

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Product/Ajax/ProductAdd.asp",
				async		 : false,
				data		 : "Seq=" + seq + "&SalePriceType=" + salePriceType + "&ProductCode=" + productCode + "&SizeCD=" + sizeCD + "&OProductCode=" + oProductCode + "&OSizeCD=" + oSizeCD,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									$("#SelectProductList").append(cont);

									// 총금액 계산
									computeTotalPrice();

									// 하단 구매 레이어 선택상품 스크롤 내리기
									$("#PopPurchase .selected-item").animate({ scrollTop: 10000 }, 100);
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 하단 선택상품 수량 증감 */
		function changeQty(seq, cnt) {
			// +,- 버튼 활성화 체크
			if (cnt > 0) {
				if ($(".Product" + seq).find("button.btn-plus").hasClass("dis")) {
					return;
				}
			} else {
				if ($(".Product" + seq).find("button.btn-minus").hasClass("dis")) {
					return;
				}
			}

			var orderCnt = Number($(".Product" + seq).find("span.product-length").text());
			orderCnt = orderCnt + cnt;
			// 최대,최소 구매수량 
			$(".Product" + seq).find("button.btn-minus").removeClass("dis");
			$(".Product" + seq).find("button.btn-plus").removeClass("dis");
			if (orderCnt <= 1) {
				orderCnt = 1;
				$(".Product" + seq).find("button.btn-minus").addClass("dis");
			} else if (orderCnt >= 99) {
				orderCnt = 100;
				$(".Product" + seq).find("button.btn-plus").addClass("dis");
			}
			$(".Product" + seq).find("span.product-length").text(orderCnt);

			var salePrice = Number($(".Product" + seq).find("input[name='SalePrice']").val());
			var orderPrice = salePrice * orderCnt;

			// 총금액에 콤마 넣기
			orderPrice = String(orderPrice);
			orderPrice = orderPrice.replace(/(\d)(?=(?:\d{3})+(?!\d))/g, "$1,");
			$(".Product" + seq).find("div.cost span.saleprice").text(orderPrice);

			computeTotalPrice();
		}

		/* 하단 선택상품 삭제 */
		function deleteProduct(seq) {
			$(".Product" + seq).remove();
			computeTotalPrice();
		}

		/* 하단 선택상품에 담긴 상품총액 */
		function computeTotalPrice() {
			var totalPrice = 0;
			$("#SelectProductList li").each(function () {
				if ($(this).find("div.cost span.saleprice").length == 1) {
					var salePrice = Number($(this).find("div.cost span.saleprice").text().replace(/[^\d]+/g, ''));
					totalPrice += salePrice;
				}
			});

			if (parseFloat(totalPrice) == 0) {
				$("input:radio[name='chk-size']").prop("checked", false);
			}

			// 총금액에 콤마 넣기
			totalPrice = String(totalPrice);
			totalPrice = totalPrice.replace(/(\d)(?=(?:\d{3})+(?!\d))/g, "$1,");

			$("#TotalPrice").html(totalPrice);
		}

		/* 선택된 상품을 주문하기위한 리스트로 만들기 */
		function makeProductList(orderType, delvType) {
			var productCodes	 = "";
			var sizeCDs			 = "";
			var orderCnts		 = "";
			var salePriceTypes	 = "";
			var productTypes	 = "";

			$("#SelectProductList li").each(function () {
				var productCode		= $(this).find("input[name='ProductCode']").val();
				var sizeCD			= $(this).find("input[name='SizeCD']").val();
				var orderCnt		= $(this).find("span.product-length").text();
				var salePriceType	= $(this).find("input[name='SalePriceType']").val();
				var productType		= $(this).find("input[name='ProductType']").val();

				if (productCodes == "") {
					productCodes	= productCode;
					sizeCDs			= sizeCD;
					orderCnts		= orderCnt;
					salePriceTypes	= salePriceType;
					productTypes	= productType;
				} else {
					productCodes	= productCodes		+ "," + productCode;
					sizeCDs			= sizeCDs			+ "," + sizeCD;
					orderCnts		= orderCnts			+ "," + orderCnt;
					salePriceTypes	= salePriceTypes	+ "," + salePriceType;
					productTypes	= productTypes		+ "," + productType;
				}
			});

			if (productCodes == "") {
				common_msgPopOpen("", "선택된 상품이 없습니다.", "", "msgPopup", "N");
				return false;
			}

			$("form[name='OrderForm'] input[name='OrderType']").val(orderType);
			$("form[name='OrderForm'] input[name='DelvType']").val(delvType);
			$("form[name='OrderForm'] input[name='ProductCode']").val(productCodes);
			$("form[name='OrderForm'] input[name='SizeCD']").val(sizeCDs);
			$("form[name='OrderForm'] input[name='OrderCnt']").val(orderCnts);
			$("form[name='OrderForm'] input[name='SalePriceType']").val(salePriceTypes);
			$("form[name='OrderForm'] input[name='ProductType']").val(productTypes);

			return true;
		}


		/* 장바구니 담기 */
		function addCart() {
			// 장바구니에 담을 상품리스트 작성하기
			if (makeProductList("G", "P") == false) {
				return;
			}

			//에이스카운터 장바구니 담기
			AM_PRODUCT(document.OrderForm.OrderCnt.value);

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Order/Ajax/CartProductAddOk.asp",
				async		 : false,
				data		 : $("form[name='OrderForm']").serialize(),
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									get_GNB_CartCount();
									init_PurchaseBox();
									openAlertLayer2("confirm", "선택하신 상품을 장바구니에 담았습니다.<br />장바구니로 이동 하시겠습니까?", "장바구니로 이동", "쇼핑 계속하기", "closePop('confirmPop', '');APP_GoUrl('/ASP/Order/CartList.asp');", "closePop('confirmPop', '');");
									return;
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 바로구매/예약구매/매장픽업 하기 */
		function addOrderSheet(orderType, delvType, payType) {
			// 주문서에 담을 상품리스트 작성하기
			if (makeProductList(orderType, delvType) == false) {
				return;
			}

			//에이스카운터 장바구니 담기
			AM_PRODUCT(document.OrderForm.OrderCnt.value);

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Order/Ajax/OrderSheet_ProductAddOk.asp",
				async		 : false,
				data		 : $("form[name='OrderForm']").serialize(), // "OrderType=" + orderType + "&ProductCode=" + productCodes + "&SizeCD=" + sizeCDs + "&OrderCnt=" + orderCnts + "&SalePriceType=" + salePriceTypes + "&ProductType=" + productTypes,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									init_PurchaseBox();
									APP_GoUrl("/ASP/Order/Order.asp?IsOrder=Yes&AccessType=ProductOrder&PayType=" + payType);
									return;
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		// 알림창 띄우기
		function alertPopup(alertType) {
			$.ajax({
				type: "post",
				url: "/Common/Ajax/AlertPopup.asp",
				async: false,
				data: "AlertType=" + alertType,
				dataType: "text",
				success: function (data) {
					$("#msgPopup").html(data);
					openPop('msgPopup');
				},
				error: function (data) {
					alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}

		/* 재입고 알림 Pop Up 호출 */
		function Reentry_Open() {
			$.ajax({
				type: "post",
				url: "/ASP/Product/Ajax/ProductReentryAdd.asp",
				async: false,
				data: "ProductCode=<%=ProductCode%>",
				dataType: "text",
				success: function (data) {
					$("#DimDepth1").html(data);
					openPop('DimDepth1');
				},
				error: function (data) {
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});

		}

		/* 재입고 알림 등록처리 */
		function Reentry_Insert() {
			var Reentry_SizeCD = alltrim(document.ReentryForm.Reentry_SizeCD.value);
			if (Reentry_SizeCD.length == 0) {
				common_msgPopOpen("", "사이즈를 선택해 주십시오.");
				return;
			}
			var Mobile1 = alltrim(document.ReentryForm.Mobile1.value);
			var Mobile2 = alltrim(document.ReentryForm.Mobile2.value);
			if (only_Num(Mobile2) == false) {
				common_msgPopOpen("", "휴대전화번호를 숫자로만 입력하여 주세요.", "", "Mobile2");
				return;
			}
			if (Mobile1 == "010") {
				if (Mobile2.length != 8) {
					common_msgPopOpen("", "휴대전화번호를 정확히 입력하여 주세요.", "", "Mobile2");
					return;
				}
			}
			else {
				if (Mobile2.length < 7) {
					common_msgPopOpen("", "휴대전화번호를 정확히 입력하여 주세요.", "", "Mobile2");
					return;
				}
			}

			if ($("input:checkbox[id='clause-agree']").prop("checked") == false)
			{
				common_msgPopOpen("", "개인 정보 이용에 동의를 하셔야 합니다.");
				return;
			}

			$.ajax({
				type: "post",
				url: "/ASP/Product/Ajax/ProductReentryAddOk.asp",
				async: false,
				data: $("form[name='ReentryForm']").serialize(),
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						common_msgPopOpen("", "재입고 알림이 등록 되었습니다.", "closePop('DimDepth1');");
						return;						
					}
					else {
						common_msgPopOpen("", cont);
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
				}
			});

		}

		//상품 이미지 확대
		function ProductZoomOpen() {
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Product/Ajax/ProductImageZoom.asp",
				async		 : false,
				data		 : "ProductCode=<%=ProductCode%>",
				dataType	 : "text",
				success		 : function (data) {
								$("#slidecont").html(data);
								$('#ProductZoom').show();

								var zoomControl = new Swiper('.zoom-control', {
									slidesPerView: 1,
									zoom: {
										maxRatio: 3,
										minRatio: 1,
										containerClass: 'swiper-zoom-container',
									},
									centeredSlides: true,
									observer: true,
									observeParents: true,
									pagination: {
										el: '.swiper-pagination',
										clickable: true
									},
									navigation: {
										nextEl: '.swiper-button-next',
										prevEl: '.swiper-button-prev',
									},
								});
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function ProductZoomClose() {
			$("#slidecont").html('');
			$('#ProductZoom').hide();
		}

		function ReviewList(page, productcode) {
			$("#ReviewPage").val(page);
			$.ajax({
				url			 : '/ASP/Product/Ajax/ProductReviewList.asp',
				data		 : "ProductCode=" + productcode + "&Page=" + page,
				async		 : false,
				type		 : 'get',
				dataType	 : 'html',
				success		 : function (data) {
								arrData	 = data.split("|||||");
								Data	 = arrData[0];
								RecCnt	 = arrData[1];
								PageCnt	 = arrData[2];

								$("#review_morebtn").show();
								if (parseInt(page) >= parseInt(PageCnt)) {
									$("#review_morebtn").hide();
								}

								// 목록 로딩시키기
								if (page == 1) {
									$("#reviewList").html(Data);
						
								} else {
									$("#reviewList").append(Data);
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function NextReviewList(productcode) {
			var page = $("#ReviewPage").val();
			page = parseInt(page) + 1;
			ReviewList(page, productcode);
		}

		//리뷰 이미지 확대
		function ReviewImageZoomOpen(idx) {
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Product/Ajax/ReviewImageZoom.asp",
				async		 : false,
				data		 : "Idx=" + idx,
				dataType	 : "text",
				success		 : function (data) {
								$("#reviewImage-pop").html(data);
								$('#ReviewImage').show();

								var reviewImg = new Swiper('.review-img', {
									centeredSlides: true,
									observer: true,
									observeParents: true,
									pagination: {
										el: '.swiper-pagination',
										clickable: true
									},
									navigation: {
										nextEl: '.swiper-button-next',
										prevEl: '.swiper-button-prev',
									},
								});
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function ReviewImageZoomClose() {
			$("#reviewimage-pop").html('');
			$('#ReviewImage').hide();
		}

		function ProductCounsel_Reply(i, stype) {
			if (stype == 'O') {
				$("#btn_replyview" + i).hide();
				$("#counsel_reply" + i).show();
				$("#btn_replyclose" + i).show();
			}
			else {
				$("#btn_replyview" + i).show();
				$("#counsel_reply" + i).hide();
				$("#btn_replyclose" + i).hide();
			}
		}



		//찜하기
		function PickCheck(productCode) {
			var onFlag = "N";
			if ($("#Pickup").hasClass("on")) {
				onFlag = "Y";
			}

			var ret = set_MyWishList(productCode, onFlag);
			var splitData = ret.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				if (onFlag == "Y") {
					common_msgPopOpen("", "찜한상품을 해제 하였습니다.");
					$("#Pickup").removeClass("on");
				}
				else {
					common_msgPopOpen("", "찜한상품으로 저장 되었습니다.");
					$("#Pickup").addClass("on");
				}
			}
			else if (result == "LOGIN") {
				common_msgPopOpen('SHOEMARKER', '로그인 후 이용가능합니다.<br/>로그인 하시겠습니까?', '$(\'#botLoginForm\').submit();', '', 'C');
			}
			else {
				common_msgPopOpen("", cont);
			}
		}

		//공유하기
		function ShareOpen() {
			$("#ShareView").show();
		}

		function ShareClose() {
			$("#ShareView").hide();
		}

		function LoginChk()
		{
			common_msgPopOpen('SHOEMARKER', '로그인 후 이용가능합니다.<br/>로그인 하시겠습니까?', '$(\'#botLoginForm\').submit();', '', 'C');
		}

		list_ProductCounsel(1, '<%=ProductCode%>');
	</script>

	<!-- 페이스북 -->
	<script type="text/javascript">
		function facebook_share()
		{
			APP_PopupGoUrl('/API/facebook_share.asp?ProductCode=<%=ProductCode%>', '0', '');
		}
	</script>
	<!-- 페이스북 -->

	<!-- 인스타그램 -->
	<script type="text/javascript">
		function instagram_share() {
			openExternal("https://www.instagram.com/shoemarker_official/");
		}
	</script>
	<!-- 인스타그램 -->



	<!-- WIDERPLANET  SCRIPT START 2019.1.8 -->
	<div id="wp_tg_cts" style="display:none;"></div>
	<script type="text/javascript">
		var wptg_tagscript_vars = wptg_tagscript_vars || [];
		wptg_tagscript_vars.push(
		(function () {
			return {
				wp_hcuid: "<%=U_Num%>",  	/*고객넘버 등 Unique ID (ex. 로그인  ID, 고객넘버 등 )를 암호화하여 대입.
					 *주의 : 로그인 하지 않은 사용자는 어떠한 값도 대입하지 않습니다.*/
				ti: "24585",
				ty: "Item",
				device: "mobile"
				, items: [{ i: "<%=ProductCode%>", t: "<%=ProductName%>" }] /* i:<?상품 식별번호  (Feed로 제공되는 식별번호와 일치하여야 합니다 .) t:상품명  */
			};
		}));
	</script>
	<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>
	<!-- // WIDERPLANET  SCRIPT END 2019.1.8 -->

	<!-- Google Tag Manager Variable (eMnet) -->
	<script type="text/javascript">
		var brandIds = [];
		brandIds.push('<%=ProductCode%>'); 
	</script>
	<!-- End Google Tag Manager Variable (eMnet) --> 

	<!-- AceCounter Mobile eCommerce (Product_Detail) v7.5 Start -->
	<script type="text/javascript">
		var m_pd = "<%=ProductName%>";
		var m_ct = "<%=BrandName%>";
		var m_amt = "<%=SalePrice%>";
	</script>

	<!-- Facebook Pixel Code -->
	<script>
	  fbq('track', 'ViewContent', {
		value: <%=SalePrice%>,
		currency: 'KRW',
	  });
	</script>
	<!-- End Facebook Pixel Code -->

	<!-- GA -->
	<script>
		gtag('event', 'view_item', {
		  "items": [
			{
			  "id": "<%=ProductCode%>",
			  "name": "<%=ProductName%>",
			  "list_name": "ProductDetail",
			  "brand": "<%=BrandName%>",
			  "category": "<%=CategoryName1%>/<%=CategoryName2%>/<%=CategoryName3%>",
			  "variant": "<%=ColorCD%>",
			  "list_position": 1,
			  "quantity": 1,
			  "price": '<%=SalePrice%>'
			}
		  ]
		});
	</script>
	<!-- GA --> 

	<!-- kakao pixel script //-->
	<script type="text/javascript" charset="UTF-8" src="//t1.daumcdn.net/adfit/static/kp.js"></script>
	<script type="text/javascript">
		kakaoPixel('5354511058043421336').pageView();
		kakaoPixel('5354511058043421336').viewContent({
			id: '<%=ProductCode%>'
		});
	</script>
	<!-- kakao pixel script //-->

<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs1 = Nothing
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
