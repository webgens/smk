<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetUseCoupon.asp - 주문서 보유쿠폰 리스트 폼 페이지
'Date		: 2018.12.28
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
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

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
DIM z

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderSheetIdx

DIM ProductCode
DIM ProductName
DIM SizeCD
DIM ProdCD
DIM ColorCD
DIM BrandName
DIM ProductImage
DIM OrderCnt
DIM TagPrice
DIM SalePriceType
DIM SalePrice
DIM DCRate
DIM UseCouponPrice
DIM UsePointPrice
DIM UseScashPrice
DIM OrderPrice

DIM rsCnt : rsCnt = 0
DIM arrRs

DIM CheckFlag
DIM MemberCouonIdx		: MemberCouonIdx		= ""
DIM DuplicateUseFlag	: DuplicateUseFlag		= ""
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx			= sqlFilter(Request("OrderSheetIdx"))




IF OrderSheetIdx = "" THEN
		Response.Write "FAIL|||||상품이 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_By_Idx"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput, 20,		U_CARTID)
		.Parameters.Append .CreateParameter("@Idx",		adInteger,	adParamInput,   ,		OrderSheetIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("SalePriceType") = "2" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||임직원 판매가 구매시 쿠폰을 사용하실 수 없습니다."
				Response.End
		END IF

		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		BrandName		= oRs("BrandName")
		SizeCD			= oRs("SizeCD")
		ProdCD			= oRs("ProdCD")
		ColorCD			= oRs("ColorCD")
		ProductImage	= oRs("ProductImage_180")
		OrderCnt		= oRs("OrderCnt")
		TagPrice		= oRs("TagPrice")
		SalePriceType	= oRs("SalePriceType")
		IF SalePriceType = "2" THEN
				SalePrice			= oRs("EmployeeSalePrice")
				DCRate				= oRs("EmployeeDCRate")
		ELSE
				SalePrice			= oRs("SalePrice")
				DCRate				= oRs("DCRate")
		END IF
		UseCouponPrice	= oRs("UseCouponPrice")
		UsePointPrice	= oRs("UsePointPrice")
		UseScashPrice	= oRs("UseScashPrice")
		OrderPrice		= CDbl(SalePrice) - CDbl(UseCouponPrice) - CDbl(UsePointPrice) - CDbl(UseScashPrice)

		IF ProductImage = "" THEN
				ProductImage	= "/Images/180_noimage.png"
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF
oRs.Close



'# 현재 주문 상품에 대해 적용한 쿠폰 정보
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Select_By_OrderSheetIdx"

		.Parameters.Append .CreateParameter("@MemberNum",		adInteger, adParamInput, ,	U_NUM)
		.Parameters.Append .CreateParameter("@OrderSheetIdx",	adInteger, adParamInput, ,	OrderSheetIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		rsCnt = oRs.RecordCount
		arrRs = oRs.GetRows(rsCnt)
END IF
oRs.Close


'# 0 : Idx
'# 1 : OrderSheetIdx
'# 2 : MemberCouponIdx
'# 3 : MemberNum
'# 4 : MoneyType
'# 5 : DiscountPrice
'# 6 : DiscountRate
'# 7 : DiscountRateOriginal
'# 8 : CouponIdx
'# 9 : CouponType
'# 10 : DuplicateUseFlag
'# 11 : CouponName
'# 12 : StartDT
'# 13 : EndDT

Function checkUseCoupon(ByVal memberCouponIdx)
		DIM retVal : retVal = False
		IF rsCnt > 0 THEN
				FOR z = 0 TO UBound(arrRs,2)
						IF CStr(memberCouponIdx) = CStr(arrRs(2,z)) THEN
								retVal = True
								EXIT FOR
						END IF
				NEXT
		END IF
		checkUseCoupon = retVal
End Function


Response.Write "OK|||||"
%>					
        <div class="area-pop" id="UseCoupon">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">쿠폰 적용</p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="coupon-lists">
                            <div class="coupon-list">
								<span class="sub"><%=FormatNumber(SalePrice,0)%>원</span>
									- 
								<span class="sub" id="UseCouponPrice"><%=FormatNumber(UseCouponPrice,0)%>원</span>
									- 
								<span class="sub" id="UsePointPrice"><%=FormatNumber(UsePointPrice,0)%>원</span>
									- 
								<span class="sub" id="UseScashPrice"><%=FormatNumber(UseScashPrice,0)%>원</span>
									= 
								<span class="sub" id="OrderPrice"><%=FormatNumber(OrderPrice,0)%>원</span>
                            </div>
<%
'해당 상품 및 사용자가 가지고 있는 쿠폰 목록을 검색
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Coupon_Member_Select_For_OrderSheet_Use"

		.Parameters.Append .CreateParameter("@MemberNum",		adInteger,	adParamInput,   ,	U_NUM)
		.Parameters.Append .CreateParameter("@OrderSheetIdx",	adInteger,	adParamInput,   ,	OrderSheetIdx)
		.Parameters.Append .CreateParameter("@Location",		adChar,		adParamInput,  1,	"M")
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

i = 0			
IF NOT oRs.EOF THEN
		Do Until oRs.EOF
				CheckFlag = checkUseCoupon(oRs("MemberCouponIdx"))
				IF CheckFlag THEN
					IF MemberCouonIdx = "" THEN
						MemberCouonIdx		 = oRs("MemberCouponIdx")
						DuplicateUseFlag	 = oRs("DuplicateUseFlag")
					ELSE
						MemberCouonIdx		 = MemberCouonIdx	& "," & oRs("MemberCouponIdx")
						DuplicateUseFlag	 = DuplicateUseFlag & "," & oRs("DuplicateUseFlag")
					END IF
				END IF
%>
                            <div class="coupon-list">
                                <div class="tit">
                                    <div class="inn">
                                        <div class="off"><%=FormatNumber(oRs("Discount"),0)%><%IF oRs("MoneyType") = "W" THEN%>원<%ELSE%>%<%END IF%> OFF</div>
                                        <div class="name"><%=oRs("CouponName")%></div>
                                    </div>
                                </div>
                                <div class="time-limit">
                                    <em>사용기한</em>
									<%=Replace(GetDateYMD(Left(oRs("StartDT"),8)), "-", ". ")%> ~<br>
									<%IF oRs("EndDT") = "999999999999" THEN%>
									제한기간없음
									<%ELSE%>
									<%=Replace(GetDateYMD(Left(oRs("EndDT"),8)), "-", ". ")%>
									<%END IF%>
                                </div>
                                <div class="checkboxgroup sel">
                                    <div class="inner">
                                        <span class="checkbox">
											<input type="hidden" name="DuplicateUseFlag<%=oRs("MemberCouponIdx")%>" value="<%=oRs("DuplicateUseFlag")%>" />
											<input type="hidden" name="CouponName<%=oRs("MemberCouponIdx")%>" value="<%=oRs("CouponName")%>" />
											<input type="checkbox" name="MemberCouponIdx<%=oRs("MemberCouponIdx")%>" id="MemberCouponIdx<%=oRs("MemberCouponIdx")%>" value="<%=oRs("MemberCouponIdx")%>" <%IF CheckFlag THEN%>checked="checked"<%END IF%> />
										</span>
                                        <label for="MemberCouponIdx<%=oRs("MemberCouponIdx")%>">선택</label>
                                    </div>
                                </div>
                            </div>
<%
				IF CheckFlag THEN
					Response.Write "<script>$(""#UseCoupon input[name='MemberCouponIdx"& oRs("MemberCouponIdx") &"']"").closest('.checkbox').addClass('is-checked');</script>"
				END IF
				oRs.MoveNext
				i = i + 1
		Loop
ELSE
%>
                            <div class="coupon-list">
                                <div class="area-empty">
									<span class="icon-empty"></span>
									<p class="tit-empty">보유중인 쿠폰이 없습니다</p>
                                </div>
                            </div>
<%
END IF
oRs.Close
%>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="applyCoupon(<%=OrderSheetIdx%>)" class="button ty-red">적용</button>
                    </div>
                </div>

				<form name="UseCouponForm" method="post">
					<input type="hidden" name="MemberCouponIdx" value="<%=MemberCouonIdx%>" />
					<input type="hidden" name="DuplicateUseFlag" value="<%=DuplicateUseFlag%>" />
				</form>

				<script type="text/javascript">
					$(function () {
						// checkbox 버튼 클릭 액션
						$('#UseCoupon .checkboxgroup input').on('click', function () {
							var $this = $(this);

							if ($this.prop('checked') === true) {
								$this.closest('.checkbox').addClass('is-checked');	//.siblings().removeClass('checked');
							} else {
								$this.closest('.checkbox').removeClass('is-checked');
							}

							checkCoupon("<%=OrderSheetIdx%>", $this.val());
						});
					})
				</script>
<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>