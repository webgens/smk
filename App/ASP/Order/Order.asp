<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Order.asp - 주문서
'Date		: 2018.12.28
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
PageCode2 = "02"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>
aaaaaaaa
<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
'# 비회원 주문을 위한 고려 사항
IF U_NUM = "" AND Request.Cookies("NON_ORDER") = "" THEN
		Response.Redirect "Login.asp?ProgID=" & Server.URLEncode("/ASP/Order/Order.asp?IsOrder=" & Request("IsOrder") & "&AccessType=" & Request("AccessType") & "&PayType=" & Request("PayType"))
		Response.End
END IF
    ''''''
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

DIM IsOrder						'//Yes:장바구니 혹은 바로주문에서 넘어오온 경우.
DIM AccessType					'//주문페이지 접근경로. Cart / ProductdOrder / Mileage / Coupon
DIM PayType


DIM ShoeMarkerPayUseFlag	: ShoeMarkerPayUseFlag = "N"		'# 슈마커페이 사용여부

DIM OrderCount			: OrderCount		= 0
DIM MultiDelvPossible	: MultiDelvPossible	= "Y"		'# 다중배송지 가능 여부
DIM MultiDelvFlag		: MultiDelvFlag		= "N"		'# 다중배송지 선택 여부
DIM ReceiverInfoFlag	: ReceiverInfoFlag	= "Y"		'# 배송지정보 표시여부

DIM ProductCode
DIM SalePrice
DIM DCRate
DIM SavePoint
DIM ProductImage

DIM PointRate
DIM OrderName
DIM OrderTel
DIM OrderTel1
DIM OrderTel2
DIM OrderTel3
DIM OrderHP
DIM OrderHP1
DIM OrderHP2
DIM OrderHP3
DIM OrderEmail
DIM OrderEmail1
DIM OrderEmail2
DIM OrderZipCode
DIM OrderAddr1
DIM OrderAddr2
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2

DIM MyAddressCount			: MyAddressCount		= 0

DIM TotalOrderCnt			: TotalOrderCnt			= 0
DIM TotalTagPrice			: TotalTagPrice			= 0
DIM TotalSalePrice			: TotalSalePrice		= 0
DIM TotalUseCouponPrice		: TotalUseCouponPrice	= 0
DIM TotalUseScashPrice		: TotalUseScashPrice	= 0
DIM TotalUsePointPrice		: TotalUsePointPrice	= 0
DIM TotalDeliveryPrice		: TotalDeliveryPrice	= 0
DIM TotalSavePoint			: TotalSavePoint		= 0
DIM ShopOrderCnt			: ShopOrderCnt			= 0
DIM ShopTagPrice			: ShopTagPrice			= 0
DIM ShopSalePrice			: ShopSalePrice			= 0
DIM ShopUseCouponPrice		: ShopUseCouponPrice	= 0
DIM ShopUseScashPrice		: ShopUseScashPrice		= 0
DIM ShopUsePointPrice		: ShopUsePointPrice		= 0
DIM ShopDeliveryPrice		: ShopDeliveryPrice		= 0
DIM ShopSavePoint			: ShopSavePoint			= 0

DIM arrTel1
DIM arrHP1
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

IsOrder				= sqlFilter(Request("IsOrder"))		
AccessType			= sqlFilter(Request("AccessType"))		
PayType				= sqlFilter(Request("PayType"))		

IF PayType = ""	THEN PayType = "C"
IF NAVER_PAY_FLAG <> "Y" AND PayType = "N" THEN PayType = "C"

MultiDelvFlag		= sqlFilter(Request("MultiDelvFlag"))		
IF MultiDelvFlag = "" THEN MultiDelvFlag = "N"



arrTel1	= ARRAY("02", "031", "032", "033", "041", "042", "043", "051", "052", "053", "054", "055", "061", "062", "063", "064", "070", "010", "011", "016", "017", "018", "019")
arrHP1	= ARRAY("010", "011", "016", "017", "018", "019")


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1		 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




'# 장바구니 UserID 변경
IF U_NUM <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Cart_Update_For_UserID"

				.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput,  20,	 U_CARTID)
				.Parameters.Append .CreateParameter("@UserID",		adVarChar,	adParamInput,  20,	 U_NUM)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
END IF


'# 주문서 체크
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_Select_For_OrderCount"

		.Parameters.Append .CreateParameter("@CartID",	 adVarChar,	 adParamInput, 20,		 U_CARTID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
																
IF NOT oRs.EOF THEN
		ProductCode	= oRs("ProductCode")

		OrderCount	= CInt(oRs("OrderCount"))

		'# 상품수가 한개이면 다중배송 불가능
		IF CDbl(oRs("DelvType_P")) = 1 THEN
				MultiDelvPossible	= "N"
		END IF

		'# 배송비가 있으면 다중배송 불가능
		IF CDbl(oRs("DeliveryPrice")) > 0 THEN
				MultiDelvPossible	= "N"
		END IF

		'# 전체주문상품이 한 개일 경우는 다중배송 불가
		IF OrderCount = 1 THEN
				MultiDelvPossible	= "N"
		END IF

		'# 전체가 매장픽업이면 다중배송 불가, 배송지정보 표시안함
		IF CDbl(oRs("OrderCount")) = CDbl(oRs("DelvType_S")) THEN
				MultiDelvPossible	= "N"
				ReceiverInfoFlag	= "N"
		END IF
END IF
oRs.Close


'# 다중배송지 불가능일 경우 단일배송지 선택으로 설정
IF MultiDelvPossible = "N" THEN
		MultiDelvFlag = "N"
END IF

'# 다중배송지 선택일 경우 배송지정보 표시안함
IF MultiDelvFlag = "Y" THEN
		ReceiverInfoFlag	= "N"
END IF


IF OrderCount = 0 THEN
		IF Request.serverVariables("HTTP_REFERER") = "" THEN
				Set oRs1 = Nothing : Set oRs = Nothing : oConn.Close : Set oConn = Nothing

				CALL AlertMessage2("주문내역이 없습니다", "location.href = '/';")
				Response.End
		ELSE
				Set oRs1 = Nothing : Set oRs = Nothing : oConn.Close : Set oConn = Nothing
		 
				CALL AlertMessage2("주문내역이 없습니다", "history.back();")
				Response.End
		END IF
END IF


PointRate			= 0
OrderName			= ""
OrderTel			= ""
OrderTel1			= ""
OrderTel2			= ""
OrderTel3			= ""
OrderHP				= ""
OrderHP1			= ""
OrderHP2			= ""
OrderHP3			= ""
OrderEmail			= ""
OrderEmail1			= ""
OrderEmail2			= ""
OrderZipCode		= ""
OrderAddr1			= ""
OrderAddr2			= ""
ReceiveZipCode		= ""
ReceiveAddr1		= ""
ReceiveAddr2		= ""

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
				PointRate			= oRs("PointRate")

				OrderName			= oRs("Name")

				OrderTel			= oRs("Tel")
				IF IsNull(OrderTel) THEN OrderTel = ""

				IF OrderTel <> "" THEN
						IF UBound(SPLIT(OrderTel,"-")) = 2 THEN
								OrderTel1					 = SPLIT(OrderTel, "-")(0)
								OrderTel2					 = SPLIT(OrderTel, "-")(1)
								OrderTel3					 = SPLIT(OrderTel, "-")(2)
						ELSEIF UBound(SPLIT(OrderTel,"-")) = 1 THEN
								OrderTel1					 = SPLIT(OrderTel, "-")(0)
								OrderTel2					 = SPLIT(OrderTel, "-")(1)
								OrderTel3					 = ""
						ELSEIF UBound(SPLIT(OrderTel,"-")) = 0 THEN
								OrderTel1					 = SPLIT(OrderTel, "-")(0)
								OrderTel2					 = ""
								OrderTel3					 = ""
						ELSE
								OrderTel1					 = OrderTel
								OrderTel2					 = ""
								OrderTel3					 = ""
						END IF
				END IF

				OrderHP				= oRs("HP")
				IF IsNull(OrderHP) THEN OrderHP = ""

				IF OrderHP <> "" THEN
						IF UBound(SPLIT(OrderHP,"-")) = 2 THEN
								OrderHP1					 = SPLIT(OrderHP, "-")(0)
								OrderHP2					 = SPLIT(OrderHP, "-")(1)
								OrderHP3					 = SPLIT(OrderHP, "-")(2)
						ELSEIF UBound(SPLIT(OrderHP,"-")) = 1 THEN
								OrderHP1					 = SPLIT(OrderHP, "-")(0)
								OrderHP2					 = SPLIT(OrderHP, "-")(1)
								OrderHP3					 = ""
						ELSEIF UBound(SPLIT(OrderHP,"-")) = 0 THEN
								OrderHP1					 = SPLIT(OrderHP, "-")(0)
								OrderHP2					 = ""
								OrderHP3					 = ""
						ELSE
								OrderHP1					 = OrderHP
								OrderHP2					 = ""
								OrderHP3					 = ""
						END IF
				END IF

				OrderEmail			= oRs("EMail")
				IF IsNull(OrderEmail) THEN OrderEmail = ""

				IF OrderEmail <> "" THEN
						IF UBound(SPLIT(OrderEmail,"@")) = 1 THEN
								OrderEmail1					 = SPLIT(OrderEmail, "@")(0)
								OrderEmail2					 = SPLIT(OrderEmail, "@")(1)
						ELSE
								OrderEmail1					 = OrderEmail
								OrderEmail2					 = ""
						END IF
				END IF

				OrderZipCode		= oRs("ZipCode")
				OrderAddr1			= oRs("Address1")
				OrderAddr2			= oRs("Address2")
				ReceiveZipCode		= oRs("ZipCode")
				ReceiveAddr1		= oRs("Address1")
				ReceiveAddr2		= oRs("Address2")
		END IF
		oRs.Close
END IF
%>

<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.delivery-info.multidelivery { border: none; }
		.delivery-info .radiogroup { border-bottom: none; }
		.delivery-info .prev-delivery { padding: 0 0 10px; }
		.delivery-info .prev-delivery p { line-height: 18px; padding: 8px 0; }
		.delivery-info .prev-delivery button { margin-left: 10px; }
		.delivery-info .fieldset .fieldset-row p { font-size: 11px; line-height: 18px; color: #646464; }

		.coupon .listitems { margin: 10px 0; }
		.cart-show .listitems .thumbnail { height: 112px; }
		.cart-show .listitems .price span { font-size: 14px; font-weight: 800; }
		.cart-show .listitems .price em { color: #e62019; margin-left: 10px; }
		.coupon .item-info .pickup { font-size: 11px; color: #646464; }
		.coupon .badge { position: absolute; right: 0; top: 0; display: block; border: 1px solid #e62019; padding: 4px 4px 2px; color: #e62019; font-size: 10px; font-weight: 800; }
		.coupon .badge.pickup { color: #ffffff; background : #e62019; border: 1px solid #e62019; }
		.coupon .badge.oneplusone { color: #ffffff; background : #282828; border: 1px solid #282828; }

		/* 쿠폰리스트 팝업 */
		.coupon-list .sel.checkboxgroup>.inner { float: none; width: 100%; }
		.coupon-list .sel .checkbox { position: absolute; top: 50px; left: 50%; margin-left: -7px; opacity: 0.7; }
		/*.coupon-list .sel label{display: inline-block;width: 100%;margin-top: 3px;font-size: 13px;font-weight: 800;color: #fff;text-align: center;}*/
		.coupon-list .sel .checkbox + label { opacity: .3; margin-left: 0; }
		.coupon-list .sel .checkbox.is-checked { opacity: 1; }
		.coupon-list .sel .checkbox.is-checked + label { opacity: 1; }

		#UsePoint .usage button { margin-top: 2px; font-size: 10px; border: 1px solid #282828; color: #282828; padding: 7px 8px 5px; }
		#UseScash .usage button { margin-top: 2px; font-size: 10px; border: 1px solid #282828; color: #282828; padding: 7px 8px 5px; }

		#PickupStore .formfield { padding: 10px 12px 0; }
	</style>
	<%IF Request.ServerVariables("HTTPS") = "on" Then%>
	<script src="//ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js"></script>
	<%ELSE%>
	<script src="//dmaps.daum.net/map_js_init/postcode.v2.js"></script>
	<%END IF%>
	<script type="text/javascript">
		function execDaumPostcode(zipCode, addr1, addr2) {
			new daum.Postcode({
				oncomplete: function(data) {
					// 팝업에서 검색결과 항목을 클릭했을때 실행할 코드를 작성하는 부분.
					//alert(data.addressType + "-" + data.userSelectedType + "\n\n1 : " + data.roadAddress + "\n2 : " + data.autoRoadAddress + "\n\n3 : " + data.jibunAddress + "\n4 : " + data.autoJibunAddress);

					// data.addressType : 주소검색방법 R=도로명검색, J=지번검색
					// data.userSelectedType : 주소선택구분 R=도로명주소선택, J=지번주소선택
					var roadAddress = data.roadAddress;
					var jibunAddress = data.jibunAddress;
					if (data.addressType == "R") {
						if (jibunAddress == "" && data.userSelectedType == "R") {
							jibunAddress = data.autoJibunAddress;
						}
					} else {
						if (roadAddress == "" && data.userSelectedType == "J") {
							roadAddress = data.autoRoadAddress;
						}
					}

					// 도로명 주소의 노출 규칙에 따라 주소를 조합한다.
					// 내려오는 변수가 값이 없는 경우엔 공백('')값을 가지므로, 이를 참고하여 분기 한다.
					var fullRoadAddr = roadAddress; // 도로명 주소 변수
					var extraRoadAddr = ''; // 도로명 조합형 주소 변수

					// 법정동명이 있을 경우 추가한다. (법정리는 제외)
					// 법정동의 경우 마지막 문자가 "동/로/가"로 끝난다.
					if(data.bname !== '' && /[동|로|가]$/g.test(data.bname)){
						extraRoadAddr += data.bname;
					}
					// 건물명이 있고, 공동주택일 경우 추가한다.
					if(data.buildingName !== '' && data.apartment === 'Y'){
					   extraRoadAddr += (extraRoadAddr !== '' ? ', ' + data.buildingName : data.buildingName);
					}
					// 도로명, 지번 조합형 주소가 있을 경우, 괄호까지 추가한 최종 문자열을 만든다.
					if(extraRoadAddr !== ''){
						extraRoadAddr = ' (' + extraRoadAddr + ')';
					}
					// 도로명, 지번 주소의 유무에 따라 해당 조합형 주소를 추가한다.
					if(fullRoadAddr !== ''){
						fullRoadAddr += extraRoadAddr;
					}

					// 우편번호와 주소 정보를 해당 필드에 넣는다.
					document.getElementById(zipCode).value = data.zonecode; //5자리 새우편번호 사용
					document.getElementById(addr1).value = fullRoadAddr;
					document.getElementById(addr2).focus();
					//document.getElementById('sample4_jibunAddress').value = jibunAddress;

					// iframe을 넣은 element를 안보이게 한다.
					// (autoClose:false 기능을 이용한다면, 아래 코드를 제거해야 화면에서 사라지지 않는다.)
					//document.getElementById('post_layer').style.display = 'none';
					closePop('PopupPostSearch');
				},
				width: '100%',
				height: '100%',
				maxSuggestItems: 5
			}).embed(document.getElementById('PopupPostContents'));

			openPop('PopupPostSearch');
		}
	</script>
	<script type="text/javascript" src="//dapi.kakao.com/v2/maps/sdk.js?appkey=<%=KAKAO_LOGIN_CLIENTID%>&libraries=services"></script>
	<script type="text/javascript">
		function load_Map(XPoint, YPoint, ShopNm) {
			
			$("#PickupStore .selected-store em").html(ShopNm);

			var mapContainer = document.getElementById('map'), // 지도를 표시할 div 
				mapOption = { 
					center: new daum.maps.LatLng(YPoint, XPoint), // 지도의 중심좌표
					level: 3 // 지도의 확대 레벨
				};

			// 지도를 표시할 div와  지도 옵션으로  지도를 생성합니다
			var map = new daum.maps.Map(mapContainer, mapOption); 
      
			if(XPoint.length <= 0 || YPoint.length <= 0){
				var XPoint = 127.044123;
				var YPoint = 37.502603;
			}

			// 마커의 이미지정보를 가지고 있는 마커이미지를 생성합니다
			var markerPosition = new daum.maps.LatLng(YPoint, XPoint); // 마커가 표시될 위치입니다

			// 마커를 생성합니다
			var marker = new daum.maps.Marker({
				position: markerPosition
			});

			// 마커가 지도 위에 표시되도록 설정합니다
			marker.setMap(map);  

			var iwContent = '<div style="width:150px;text-align:center;padding:6px 0;font-size:14px;">'+ ShopNm +'</div>', // 인포윈도우에 표출될 내용으로 HTML 문자열이나 document element가 가능합니다
				iwPosition = new daum.maps.LatLng(YPoint, XPoint), //인포윈도우 표시 위치입니다
				iwRemoveable = true; // removeable 속성을 ture 로 설정하면 인포윈도우를 닫을 수 있는 x버튼이 표시됩니다

			// 인포윈도우를 생성하고 지도에 표시합니다
			var infowindow = new daum.maps.InfoWindow({
				position : iwPosition, 
				content : iwContent
			});

			// 마커 위에 인포윈도우를 표시합니다. 두번째 파라미터인 marker를 넣어주지 않으면 지도 위에 표시됩니다
			infowindow.open(map, marker); 

		}
	</script>

<%TopSubMenuTitle = "주문서작성"%>
<!-- #include virtual="/INC/TopCart.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">
            <div class="cart cart-show">

				<form name="OrderForm" id="OrderForm" method="post" action="OrderAddOk.asp">

                <!-- 상품별 배송지 동일 선택 -->
				<%IF MultiDelvPossible = "Y" THEN%>
                <div class="delivery-info confirm multidelivery">
                    <div class="formfield">
                        <p class="cart-tit">배송지 지정</p>
                        <fieldset>
                            <div class="radiogroup">
                                <div class="inner">
                                    <span class="radio">
                                        <input type="radio" name="MultiDelvFlag" id="MultiDelv_N" value="N" onclick="location.replace('/ASP/Order/Order.asp?IsOrder=<%=IsOrder%>&AccessType=<%=AccessType%>&PayType=<%=PayType%>&MultiDelvFlag=N')" <%IF MultiDelvFlag = "N" THEN%>checked="checked"<%END IF%> />
                                    </span>
                                    <label for="MultiDelv_N">배송지가 동일합니다.</label>
                                </div>
                                <div class="inner">
                                    <span class="radio">
                                        <input type="radio" name="MultiDelvFlag" id="MultiDelv_Y" value="Y" onclick="location.replace('/ASP/Order/Order.asp?IsOrder=<%=IsOrder%>&AccessType=<%=AccessType%>&PayType=<%=PayType%>&MultiDelvFlag=Y')" <%IF MultiDelvFlag = "Y" THEN%>checked="checked"<%END IF%> />
                                    </span>
                                    <label for="MultiDelv_Y">상품별 배송지가 다릅니다.</label>
                                </div>
                            </div>
						</fieldset>
					</div>
				</div>
				<%END IF%>

				<!--주문상품 목록-->
                <div class="coupon" id="OrderSheetList">
                </div>
                <div class="order-price">
                    <strong>총 주문수량</strong>
                    <p><%=FormatNumber(OrderCount, 0)%>개</p>
                </div>

<%
DIM Readonly
'# IF U_NUM <> "" THEN
IF Trim(Replace(OrderHP,"-","")) <> "" AND Trim(Replace(OrderEmail,"@","")) <> "" THEN
		Readonly = "readonly=""readonly"""
ELSE
		Readonly = ""
END IF
%>
				<input type="hidden" name="OrderZipCode"	id="OrderZipCode"	value="<%=OrderZipCode%>"	/>
				<input type="hidden" name="OrderAddr1"		id="OrderAddr1"		value="<%=OrderAddr1%>"	/>
				<input type="hidden" name="OrderAddr2"		id="OrderAddr2"		value="<%=OrderAddr2%>"	/>

				<!--주문자 정보-->
                <div class="consumer-confirm confirm">
                    <div class="formfield">
                        <p class="cart-tit">주문고객 정보</p>
                        <fieldset class="">
                            <legend class="hidden">기본 정보 입력</legend>
                            <div class="fieldset">
                                <label for="OrderName" class="fieldset-label">이름</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="OrderName" id="OrderName" value="<%=OrderName%>" maxlength="10" placeholder="이름" <%=Readonly%> />
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <div class="fieldset ty-col2 pt0">
                                    <label for="OrderHP23" class="fieldset-label">휴대폰 번호</label>
                                    <div class="fieldset-row">
                                        <span class="select">
                                            <select name="OrderHP1" id="OrderHP1" title="휴대폰 국번 선택">
											<%IF OrderHP1 <> "" THEN%>
                                                <option value="<%=OrderHP1%>"><%=OrderHP1%></option>
											<%ELSE%>
												<%FOR i = 0 TO UBOUND(arrHP1)%>
                                                <option value="<%=arrHP1(i)%>"><%=arrHP1(i)%></option>
												<%NEXT%>
											<%END IF%>
                                            </select>
                                            <span class="value"><%=OrderHP1%></span>
                                        </span>
                                        <span class="input">
                                            <input type="text" name="OrderHP23" id="OrderHP23" value="<%=OrderHP2 & OrderHP3%>" maxlength="8" placeholder="휴대폰번호" <%=Readonly%> />
                                        </span>
                                    </div>
                                </div>
                            </div>
                            <div class="fieldset">
                                <div class="fieldset ty-col2 pt0">
                                    <label for="OrderTel23" class="fieldset-label">전화 번호</label>
                                    <div class="fieldset-row">
                                        <span class="select">
                                            <select name="OrderTel1" id="OrderTel1" title="전화 국번 선택">
											<%IF OrderTel1 <> "" THEN%>
                                                <option value="<%=OrderTel1%>"><%=OrderTel1%></option>
											<%ELSE%>
												<%FOR i = 0 TO UBOUND(arrTel1)%>
                                                <option value="<%=arrTel1(i)%>"><%=arrTel1(i)%></option>
												<%NEXT%>
											<%END IF%>
                                            </select>
                                            <span class="value"><%=OrderTel1%></span>
                                        </span>
                                        <span class="input">
                                            <input type="text" name="OrderTel23" id="OrderTel23" value="<%=OrderTel2 & OrderTel3%>" maxlength="8" placeholder="전화번호"  <%=Readonly%> />
                                        </span>
                                    </div>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="order-email" class="fieldset-label">이메일주소</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="OrderEmail" id="OrderEmail" value="<%=OrderEmail%>" maxlength="50" placeholder="이메일 주소를 입력해주세요." <%=Readonly%> />
                                    </span>
                                </div>
                                <div class="inf-type1">
                                    <p class="tit">이메일로 주문 진행상황을 안내해드립니다.</p>
                                </div>
                            </div>
                        </fieldset>
                    </div>
                </div>

<%
IF ReceiverInfoFlag = "Y" THEN
%>
                <!-- 상품별 배송지 동일 선택 -->
                <div class="delivery-info confirm">
                    <div class="formfield">
                        <p class="cart-tit">배송 정보</p>
                        <fieldset>
                            <legend class="hidden">배송지 정보 입력</legend>
                            <div class="prev-delivery">
								<p>
									<span class="checkbox">
										<input type="checkbox" id="SameOrderer" value="Y" onclick="setReceiveInfo('1')"<%IF U_NUM <> "" THEN%>checked="checked"<%END IF%> />
									</span>
									<label for="SameOrderer">주문고객정보와 동일</label>
								</p>
								<%IF U_NUM <> "" THEN%>
                                <button type="button" onclick="openMyAddress('')">배송 목록에서 선택</button>
								<%END IF%>
                            </div>
                            <div class="fieldset">
                                <label for="ReceiveName" class="fieldset-label">받는 사람</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="ReceiveName" id="ReceiveName" value="<%=OrderName%>" maxlength="10" placeholder="입력하세요">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <legend class="hidden">연락처 정보 입력</legend>
                                <div class="fieldset ty-col2 pt0">
                                    <label for="ReceiveHP23" class="fieldset-label">휴대폰번호</label>
                                    <div class="fieldset-row">
                                        <span class="select">
                                            <select name="ReceiveHP1" id="ReceiveHP1" title="휴대폰 국번 선택">
                                                <option value="">선택</option>
												<%FOR i = 0 TO UBOUND(arrHP1)%>
                                                <option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = OrderHP1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
												<%NEXT%>
                                            </select>
                                            <span id="SReceiveHP1" class="value"><%=OrderHP1%></span>
                                        </span>
                                        <span class="input">
                                            <input type="text" name="ReceiveHP23" id="ReceiveHP23" value="<%=OrderHP2 & OrderHP3%>" maxlength="8" placeholder="휴대폰번호의 앞 번호와 뒷 번호 입력">
                                        </span>
                                    </div>
                                </div>
                                <div class="fieldset ty-col2 pt0">
                                    <div class="more-num">
                                        <label for="ReceiveTel23" class="fieldset-label">전화번호</label>
                                        <span>(선택)</span>
                                    </div>
                                    <div class="fieldset-row">
                                        <span class="select">
                                            <select name="ReceiveTel1" id="ReceiveTel1" title="전화번호 국번 선택">
                                                <option value="">선택</option>
												<%FOR i = 0 TO UBOUND(arrTel1)%>
                                                <option value="<%=arrTel1(i)%>"<%IF arrTel1(i) = OrderTel1 THEN%> selected="selected"<%END IF%>><%=arrTel1(i)%></option>
												<%NEXT%>
                                            </select>
                                            <span id="SReceiveTel1" class="value"><%=OrderTel1%></span>
                                        </span>
                                        <span class="input">
                                            <input type="text" name="ReceiveTel23" id="ReceiveTel23" value="<%=OrderTel2 & OrderTel3%>" maxlength="8" placeholder="전화번호의 앞 번호와 뒷 번호 입력">
                                        </span>
                                    </div>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="delivery-address11" class="fieldset-label">배송 주소</label>
                                <div class="postnum">
                                    <button class="search-postnum" type="button" onclick="execDaumPostcode('ReceiveZipCode','ReceiveAddr1','ReceiveAddr2')"><span>우편번호 검색</span></button>
                                    <div class="fieldset-row delivery-num">
                                        <span class="input is-expand">
                                            <input type="text" name="ReceiveZipCode" id="ReceiveZipCode" value="<%=ReceiveZipCode%>" placeholder="우편번호" readonly="readonly" />
                                        </span>
                                    </div>
                                </div>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="ReceiveAddr1" id="ReceiveAddr1" value="<%=ReceiveAddr1%>" placeholder="주소 입력" readonly="readonly" />
                                    </span>
                                </div>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="ReceiveAddr2" id="ReceiveAddr2" value="<%=ReceiveAddr2%>" maxlength="25" placeholder="나머지 주소를 입력해주세요." />
                                    </span>
                                </div>
								<%IF U_NUM <> "" THEN%>
                                <div class="fieldset-row">
                                    <span class="checkbox">
                                        <input type="checkbox" name="MainFlag" id="MainFlag" value="Y" <%IF MyAddressCount = 0 THEN%>checked="checked"<%END IF%> />
                                    </span>
                                    <label for="detail-address11">기본배송지로 설정(회원정보 주소가 변경됩니다.)</label>
                                </div>
								<%END IF%>
                            </div>
                            <!-- 배송 요청사항 항목선택 시 -->
                            <div class="fieldset require">
                                <label for="Memo" class="fieldset-label">배송 요청사항</label>
                                <span class="select">
                                    <select name="selMemo" title="배송 요청사항 선택">
                                        <option value="">직접입력</option>
                                        <option value="1">경비실에 맡겨주세요</option>
                                        <option value="2">부재시 경비실에 맡겨주세요</option>
                                        <option value="3">부재시 문앞에 놔주세요</option>
                                        <option value="4">택배함에 넣어주세요</option>
                                        <option value="5">배송전 연락주세요</option>
                                    </select>
                                    <span class="value">직접입력</span>
                                </span>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Memo" id="Memo" value="" maxlength="40" placeholder="요청사항을 입력해주세요." />
                                    </span>
                                </div>
                            </div>
                        </fieldset>
                    </div>
                </div>
<%
ELSE
%>
                <div class="delivery-info confirm">
                    <div class="formfield">
                        <p class="cart-tit">배송 요청사항</p>
                        <fieldset>
                            <div class="fieldset require">
                                <span class="select">
                                    <select name="selMemo" title="배송 요청사항 선택">
                                        <option value="">직접입력</option>
                                        <option value="1">경비실에 맡겨주세요</option>
                                        <option value="2">부재시 경비실에 맡겨주세요</option>
                                        <option value="3">부재시 문앞에 놔주세요</option>
                                        <option value="4">택배함에 넣어주세요</option>
                                        <option value="5">배송전 연락주세요</option>
                                    </select>
                                    <span class="value"></span>
                                </span>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Memo" id="Memo" value="" maxlength="40" placeholder="요청사항을 입력해주세요." />
                                    </span>
                                </div>
                            </div>
                        </fieldset>
                    </div>
                </div>
<%
END IF
%>

                <div class="final-price" id="PaymentInfo">
                    <p class="cart-tit">최종 결제금액</p>
                    <div class="price">
                        <div class="price-info">
                            <div class="info-wrap">
                                <p class="price-tit">주문 상품 수</p>
                                <p class="price-value">0개</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">총 상품금액</p>
                                <p class="price-value">0원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">쿠폰 적용</p>
                                <p class="price-value">0원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">포인트 사용</p>
                                <p class="price-value">0원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">슈즈상품권 사용</p>
                                <p class="price-value">0원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">총 배송비</p>
                                <p class="price-value">0원</p>
                            </div>
                        </div>
                        <div class="order-price">
                            <strong>총 결제금액</strong>
                            <p>0원</p>
                        </div>
                    </div>
                </div>
                <div class="pay-method">
                    <p class="cart-tit">결제방법 선택</p>
                    <div class="pop-brand">
                        <span class="check-style"><input type="radio" name="PayType" id="PayType_C" value="C" onclick="setPayType()" <%IF PayType = "C" THEN%>checked="checked"<%END IF%> /><label for="PayType_C"><span>신용카드</span></label>
                        </span>
                        <span class="check-style"><input type="radio" name="PayType" id="PayType_B" value="B" onclick="setPayType()" <%IF PayType = "B" THEN%>checked="checked"<%END IF%> /><label for="PayType_B"><span>계좌이체</span></label>
                        </span>
                        <span class="check-style"><input type="radio" name="PayType" id="PayType_V" value="V" onclick="setPayType()" <%IF PayType = "V" THEN%>checked="checked"<%END IF%> /><label for="PayType_V"><span>무통장입금</span></label>
                        </span>
						<%IF NAVER_PAY_FLAG = "Y" THEN%>
                        <span class="check-style"><input type="radio" name="PayType" id="PayType_N" value="N" onclick="setPayType()" <%IF PayType = "N" THEN%>checked="checked"<%END IF%> /><label for="PayType_N"><span>네이버페이</span></label>
                        </span>
						<%END IF%>
                    </div>
                </div>

				<input type="hidden" name="USAFE_FLAG" value="<%=USAFE_FLAG%>" />
				<input type="hidden" name="USAFE_PAYTYPE" value="<%=USAFE_PAYTYPE%>" />

				<!-- 신용카드 -->
				<div class="inf-type1 paydesc" id="PayDesc_C">
                    <p class="tit"><button type="button" onclick="info_Installment()" class="underline">무이자할부 안내</button></p>
                </div>

				<!-- 계좌이체 -->
				<div class="inf-type1 paydesc" id="PayDesc_B">
                    <p class="tit">계좌이체 안내</p>
                    <ul>
                        <li class="bullet-ty1">계좌이체는 결제 금액이 회원님의 계좌에서 자동으로 이체되는 서비스 입니다.</li>
                        <li class="bullet-ty1">계좌이체는 범용 공인인증서나 은행 제한용 공인인증서가 이용하시는 컴퓨터에 저장되어 있어야만 사용가능합니다.</li>
                    </ul>
					<%IF USAFE_FLAG = "Y" AND USAFE_PAYTYPE = "B" THEN%>
                    <ul style="margin-top:20px;">
                        <li>
							<div class="agree-receive">
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">보증보험</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsurance" id="GuaranteeInsurance_Yes" value="Y">
												</span>
                                                <label for="GuaranteeInsurance_Yes">발행함</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsurance" id="GuaranteeInsurance_No" value="N">
												</span>
                                                <label for="GuaranteeInsurance_No">발행안함</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">개인정보 이용동의</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsuranceAgreement" id="GuaranteeInsuranceAgreement_Yes" value="Y">
												</span>
                                                <label for="GuaranteeInsuranceAgreement_Yes">동의함</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsuranceAgreement" id="GuaranteeInsuranceAgreement_No" value="N">
												</span>
                                                <label for="GuaranteeInsuranceAgreement_No">동의안함</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">생년월일</label>
                                    <div class="fieldset-row">
										<span class="input">
											<input type="number" name="USafeYear" style="width: 53px; height:22px;"/>
										</span>
										년
										<span class="select" style="width:70px; height:24px;">
											<select name="USafeMonth">
												<option value="">월</option>
												<%FOR i=1 TO 12%>
												<option value="<%=MakeZeroChr(i,2)%>"><%=i%> 월</option>
												<%NEXT %>
											</select>
											<span class="value" style="line-height:24px"></span>
										</span>
										<span class="select" style="width:70px; height:24px;">
											<select name="USafeDay">
												<option value="">일</option>
												<%FOR i=1 TO 31%>
												<option value="<%=MakeZeroChr(i,2)%>"><%=i%> 일</option>
												<%NEXT %>
											</select>
											<span class="value" style="line-height:24px"></span>
										</span>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">성별</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="USafeSex" id="USafeSex_1" value="1">
												</span>
                                                <label for="USafeSex_1">남</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="USafeSex" id="USafeSex_2" value="2">
												</span>
                                                <label for="USafeSex_2">여</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </li>
                        <li class="bullet-ty1">
							물품대금결제 시 구매자의 피해보호를 위해 “(주)서울보증보험”의 보증보험이 발급됩니다.<br/>
							증권이 발급되는 것의 의미는, 물품대금 결제 시에 소비자에게 서울보증보험의 쇼핑몰보증보험 계약체결서를 인터넷상으로 자동 발급하며, 피해발생 시 쇼핑몰보증보험으로써 완벽하게 보호받을 수 있습니다.<br/>
							또한, 입력하신 개인정보는 증권발급을 위해 필요한 정보이며 다른 용도로 사용되지 않습니다.<br/>
							(전자보증서비스의 보험료는 쇼핑몰에서 부담합니다)
                        </li>
                    </ul>
					<%END IF%>
                </div>

				<!-- 무통장입금 -->
				<div class="inf-type1 paydesc" id="PayDesc_V">
                    <p class="tit">무통장 입금</p>
                    <ul>
                        <li class="bullet-ty1">무통장 입금을 선택하시면 개인별 가상계좌가 부여됩니다.</li>
                        <li class="bullet-ty1">계좌로 정확한 주문금액(원 단위까지)을 입금해 주시기 바랍니다.</li>
                        <li class="bullet-ty1">입금이 완료되면 자동으로 결제 확인 처리됩니다.</li>
                        <li class="bullet-ty1">이체 가능수단 : 인터넷뱅킹, 텔레뱅킹, ATM/CD기, 통장, 카드이체 등</li>
                        <li class="bullet-ty1">가상계좌는 예약이체가 되지 않습니다.</li>
                        <li class="bullet-ty1">주문완료 후 <%=MALL_CLOSEDATE%>일이내 입금완료 되지 않으면 주문이 자동 취소 됩니다.</li>
                        <li class="bullet-ty1">현금영수증 발행은 입금완료 후 주문내역에서 발급받으실 수 있습니다.</li>
                    </ul>
					<%IF USAFE_FLAG = "Y" AND USAFE_PAYTYPE = "V" THEN%>
                    <ul style="margin-top:20px;">
                        <li>
							<div class="agree-receive">
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">보증보험</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsurance" id="GuaranteeInsurance_Yes" value="Y">
												</span>
                                                <label for="GuaranteeInsurance_Yes">발행함</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsurance" id="GuaranteeInsurance_No" value="N">
												</span>
                                                <label for="GuaranteeInsurance_No">발행안함</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">개인정보 이용동의</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsuranceAgreement" id="GuaranteeInsuranceAgreement_Yes" value="Y">
												</span>
                                                <label for="GuaranteeInsuranceAgreement_Yes">동의함</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="GuaranteeInsuranceAgreement" id="GuaranteeInsuranceAgreement_No" value="N">
												</span>
                                                <label for="GuaranteeInsuranceAgreement_No">동의안함</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">생년월일</label>
                                    <div class="fieldset-row">
										<span class="input">
											<input type="number" name="USafeYear" style="width: 53px; height:22px;"/>
										</span>
										년
										<span class="select" style="width:70px; height:24px;">
											<select name="USafeMonth">
												<option value="">월</option>
												<%FOR i=1 TO 12%>
												<option value="<%=MakeZeroChr(i,2)%>"><%=i%> 월</option>
												<%NEXT %>
											</select>
											<span class="value" style="line-height:24px"></span>
										</span>
										<span class="select" style="width:70px; height:24px;">
											<select name="USafeDay">
												<option value="">일</option>
												<%FOR i=1 TO 31%>
												<option value="<%=MakeZeroChr(i,2)%>"><%=i%> 일</option>
												<%NEXT %>
											</select>
											<span class="value" style="line-height:24px"></span>
										</span>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">성별</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="USafeSex" id="USafeSex_1" value="1">
												</span>
                                                <label for="USafeSex_1">남</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" name="USafeSex" id="USafeSex_2" value="2">
												</span>
                                                <label for="USafeSex_2">여</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </li>
                        <li class="bullet-ty1">
							물품대금결제 시 구매자의 피해보호를 위해 “(주)서울보증보험”의 보증보험이 발급됩니다.<br/>
							증권이 발급되는 것의 의미는, 물품대금 결제 시에 소비자에게 서울보증보험의 쇼핑몰보증보험 계약체결서를 인터넷상으로 자동 발급하며, 피해발생 시 쇼핑몰보증보험으로써 완벽하게 보호받을 수 있습니다.<br/>
							또한, 입력하신 개인정보는 증권발급을 위해 필요한 정보이며 다른 용도로 사용되지 않습니다.<br/>
							(전자보증서비스의 보험료는 쇼핑몰에서 부담합니다)
                        </li>
                    </ul>
					<%END IF%>
                </div>

				<!-- 네이버페이 -->
				<div class="inf-type1 paydesc" id="PayDesc_N">
                    <p class="tit">네이버페이 안내</p>
                    <ul>
                        <li class="bullet-ty1">네이버페이는 네이버ID로 별도 앱 설치 없이 신용카드 또는 은행계좌 정보를 등록하여 네이버페이 비밀번호로 결제할 수 있는 간편결제 서비스입니다.</li>
						<li class="bullet-ty1">주문 변경 시 카드사 혜택 및 할부 적용 여부는 해당 카드사 정책에 따라 변경될 수 있습니다.</li>
						<li class="bullet-ty1">결제 가능한 신용카드 : 신한, 삼성, 현대, BC, 국민, 하나, 롯데, NH농협, 씨티</li>
						<li class="bullet-ty1">결제 가능한 은행 : NH농협, 국민, 신한, 우리, 기업, SC제일, 부산, 경남, 수협, 우체국</li>
						<li class="bullet-ty1">네이버페이 카드 간편결제는 네이버페이에서 제공하는 카드사 별 무이자, 청구할인 혜택을 받을 수 있습니다.</li>
                    </ul>
                </div>


                <div class="fieldset final-confirm">
					<div class="personal-collect" style="height:130px; display:none">
                        <p class="tit2">개인정보 수집 범위</p>
                        <ul class="cnt" style="height:calc(100% - 24px)">
                            <li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.
								<p>- 필수항목: 성명(실명), 전화연락처(휴대폰 번호), 기타 본 서비스 이용과정에서 발생한 상품 또는 서비스 구매 내역, 접속 기록, 쿠키, 가입인증정보</p>
							</li>
                        </ul>
                    </div>
                    <span class="checkbox" style="display:none">
                        <input type="checkbox" name="AgreementFlag" id="AgreementFlag" value="Y" checked="checked" />
                    </span>
                    <label for="AgreementFlag" style="display:none">주문정보 및 결제 내용을 확인했으며, 이에 동의합니다. (필수)</label>
                    <div class="fieldset confirm-btn">
                        <a href="javascript:order();" class="button is-expand ty-red">결제하기</a>
                    </div>
                </div>

				</form>

            </div>
        </div>
    </main>

    <script type="text/javascript">
    	$(function () {
    		// 주문상품 리스트 가져오기
    		getOrderSheetList("<%=MultiDelvFlag%>");

    		// 결제수단 안내 설정
    		setPayType();

    		// 배송메세지 예시 선택시
    		$("form[name='OrderForm'] select[name='selMemo']").on("change", function () {
    			var _this = $("form[name='OrderForm'] select[name='selMemo'] option:selected");
    			if (_this.val() == "") {
    				$("form[name='OrderForm'] input[name='Memo']").val("");
    				$("form[name='OrderForm'] input[name='Memo']").prop("readonly", false);
    				$("form[name='OrderForm'] input[name='Memo']").focus();
    			} else {
    				$("form[name='OrderForm'] input[name='Memo']").val(_this.text());
    				$("form[name='OrderForm'] input[name='Memo']").prop("readonly", true);
				}
    		});

    		// 보증보험 발행여부 선택시
    		$("input[name='GuaranteeInsurance']").change(function () {
    			if ($(this).val() == "N") {
    				$("input[name='GuaranteeInsuranceAgreement']").attr("disabled", true);
    				$("input[name='USafeYear']").attr("disabled", true);
    				$("select[name='USafeMonth']").attr("disabled", true);
    				$("select[name='USafeDay']").attr("disabled", true);
    				$("input[name='USafeSex']").attr("disabled", true);
    			} else {
    				$("input[name='GuaranteeInsuranceAgreement']").removeAttr("disabled");
    				$("input[name='USafeYear']").removeAttr("disabled");
    				$("select[name='USafeMonth']").removeAttr("disabled");
    				$("select[name='USafeDay']").removeAttr("disabled");
    				$("input[name='USafeSex']").removeAttr("disabled");
    			}
    		});
    	});

    	/* 무이자할부안내 */
    	function info_Installment() {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Product/Ajax/InfoInstallment.asp",
    			async: false,
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

    	// 주문상품 리스트 가져오기
    	function getOrderSheetList(multiDelvFlag) {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetList.asp",
    			async: false,
    			data: "MultiDelvFlag=" + multiDelvFlag,
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var orderSheetList = splitData[1];
    				var paymentInfo = splitData[2];
    				var trackingInfo = splitData[3];

    				if (result == "OK") {
    					$("#OrderSheetList").html(orderSheetList);
    					$("#PaymentInfo").html(paymentInfo);

    					//구글 트래킹
    					arrtrack = trackingInfo.split("|||");
    					for (i = 0; i < arrtrack.length; i++) {

    						temptracinfo = arrtrack[i].split("|,|");

    						gtag('event', 'begin_checkout', {
    							"items": [
								  {
								  	"id": temptracinfo[1],
								  	"name": temptracinfo[2],
								  	"list_name": "Order",
								  	"brand": temptracinfo[3],
								  	"category": "",
								  	"variant": "",
								  	"list_position": 1,
								  	"quantity": 1,
								  	"price": temptracinfo[0]
								  }
    							],
    							"coupon": ""
    						});
    					}
    				}
    				else {
    					$("#OrderSheetList").html("");
    					$("#PaymentInfo").html("");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 쿠폰 조회
    	function openUseCoupon(orderSheetIdx) {
    		//location.href = "/ASP/Order/Ajax/OrderSheetUseCoupon.asp?" + "OrderSheetIdx=" + orderSheetIdx;
    		//return;

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetUseCoupon.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx,
    			dataType: "text",
    			success: function (data) {
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
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	/* 쿠폰 선택시 사용가능여부 체크 */
    	function checkCoupon(orderSheetIdx, idx) {
    		var memberCouponIdx		= $("#UseCoupon input[name='MemberCouponIdx" + idx + "']").val();
    		var duplicateUseFlag	= $("#UseCoupon input[name='DuplicateUseFlag" + idx + "']").val();
    		var couponName			= $("#UseCoupon input[name='CouponName" + idx + "']").val();

    		var sMemberCouponIdx	= document.UseCouponForm.MemberCouponIdx.value;
    		var sDuplicateUseFlag	= document.UseCouponForm.DuplicateUseFlag.value;

    		if ($("#UseCoupon input[name='MemberCouponIdx" + idx + "']").is(":checked") == true) {
    			if (duplicateUseFlag == "N") {
    				if (sMemberCouponIdx != "") {
    					common_msgPopOpen("주문서", "[" + couponName + "] 쿠폰은 다른 쿠폰과 중복으로 사용하실 수 없습니다.", "", "msgPopup", "N");
    					$("#UseCoupon input[name='MemberCouponIdx" + idx + "']").prop("checked", false);
    					$("#UseCoupon input[name='MemberCouponIdx" + idx + "']").closest('.checkbox').removeClass('is-checked');
    					return;
    				}
    			} else {
    				if (sDuplicateUseFlag.indexOf("N") > -1) {
    					common_msgPopOpen("주문서", "[" + couponName + "] 쿠폰은 단독사용 쿠폰과 같이 사용하실 수 없습니다.", "", "msgPopup", "N");
    					$("#UseCoupon input[name='MemberCouponIdx" + idx + "']").prop("checked", false);
    					$("#UseCoupon input[name='MemberCouponIdx" + idx + "']").closest('.checkbox').removeClass('is-checked');
    					return;
    				}
    			}

    			if (sMemberCouponIdx == "") {
    				sMemberCouponIdx = memberCouponIdx;
    				sDuplicateUseFlag = duplicateUseFlag;
    			} else {
    				sMemberCouponIdx = sMemberCouponIdx + "," + memberCouponIdx;
    				sDuplicateUseFlag = sDuplicateUseFlag + "," + duplicateUseFlag;
    			}
    		} else {
    			sMemberCouponIdx = "";
    			sDuplicateUseFlag = "";

    			$("#UseCoupon input[type='checkbox']:checked").each(function () {
    				memberCouponIdx = $(this).val();
    				duplicateUseFlag = $("#UseCoupon input[name='DuplicateUseFlag" + memberCouponIdx + "']").val();

    				if (sMemberCouponIdx == "") {
    					sMemberCouponIdx = memberCouponIdx;
    					sDuplicateUseFlag = duplicateUseFlag;
    				} else {
    					sMemberCouponIdx = sMemberCouponIdx + "," + memberCouponIdx;
    					sDuplicateUseFlag = sDuplicateUseFlag + "," + duplicateUseFlag;
    				}
    			});
    		}

    		document.UseCouponForm.MemberCouponIdx.value = sMemberCouponIdx;
    		document.UseCouponForm.DuplicateUseFlag.value = sDuplicateUseFlag;

    		//if (sMemberCouponIdx != "") {
    		useCoupon(orderSheetIdx, memberCouponIdx, sMemberCouponIdx, "N");
    		//}
    	}

    	// 쿠폰 리스트에서 쿠폰 적용하기 클릭시
    	function applyCoupon(orderSheetIdx) {
    		var sMemberCouponIdx = document.UseCouponForm.MemberCouponIdx.value;

    		useCoupon(orderSheetIdx, -1, sMemberCouponIdx, "Y");
    	}

    	/* 쿠폰 적용 */
    	function useCoupon(orderSheetIdx, memberCouponIdx, sMemberCouponIdx, applyFlag) {
    		var multiDelvFlag = $("#MultiDelvFlag").val();

    		//location.href = "/ASP/Order/Ajax/OrderSheetUseCouponCheck.asp?" + "OrderSheetIdx=" + orderSheetIdx + "&MemberCouponIdx=" + sMemberCouponIdx + "&ApplyFlag=" + applyFlag;
    		//return;

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetUseCouponCheck.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx + "&MemberCouponIdx=" + sMemberCouponIdx + "&ApplyFlag=" + applyFlag,
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
    					if (applyFlag == "N") {
    						var splitData1 = cont.split("|");
    						var useCouponPrice = splitData1[0];
    						var usePointPrice = splitData1[1];
    						var useScashPrice = splitData1[2];
    						var orderPrice = splitData1[3];
    						$("#UseCouponPrice").html(useCouponPrice + "원");
    						$("#UsePointPrice").html(usePointPrice + "원");
    						$("#UseScashPrice").html(useScashPrice + "원");
    						$("#OrderPrice").html(orderPrice + "원");
    					}
    					else {
    						getOrderSheetList(multiDelvFlag);
    						closePop('DimDepth1');
    					}
    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					if (applyFlag == "N") {
    						$("#UseCoupon input[name='MemberCouponIdx" + memberCouponIdx + "']").prop("checked", false);
    						$("#UseCoupon input[name='MemberCouponIdx" + memberCouponIdx + "']").closest('.list-type2 li').removeClass('checked');
    						alert(cont);

    						var sMemberCouponIdx = "";
    						var sDuplicateUseFlag = "";

    						$("#UseCoupon .list-type2 input[type='checkbox']:checked").each(function () {
    							memberCouponIdx = $(this).val();
    							duplicateUseFlag = $("#UseCoupon input[name='DuplicateUseFlag" + memberCouponIdx + "']").val();

    							if (sMemberCouponIdx == "") {
    								sMemberCouponIdx = memberCouponIdx;
    								sDuplicateUseFlag = duplicateUseFlag;
    							} else {
    								sMemberCouponIdx = sMemberCouponIdx + "," + memberCouponIdx;
    								sDuplicateUseFlag = sDuplicateUseFlag + "," + duplicateUseFlag;
    							}
    						});
    						document.UseCouponForm.MemberCouponIdx.value = sMemberCouponIdx;
    						document.UseCouponForm.DuplicateUseFlag.value = sDuplicateUseFlag;
    					}
    					else {
    						common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					}
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 포인트 조회
    	function openUsePoint(orderSheetIdx) {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetUsePoint.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx,
    			dataType: "text",
    			success: function (data) {
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
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 보유 포인트 전체 사용
    	function usePointAll() {
    		var usePointPrice = $("#UsePoint input[name='UsablePointPrice']").val();
    		usePointPrice = usePointPrice.replace(/(\d)(?=(?:\d{3})+(?!\d))/g, "$1,");

    		$("#UsePoint input[name='UsePointPrice']").val(usePointPrice);
    	}

    	function usePoint(orderSheetIdx) {
    		var multiDelvFlag = $("#MultiDelvFlag").val();

    		var usablePointPrice = $("#UsePoint input[name='UsablePointPrice']").val();
    		var usePointPrice = $("#UsePoint input[name='UsePointPrice']").val().replace(/,/g, "");

    		if (usePointPrice.length == 0) {
    			usePointPrice = "0";
    		} else if (only_Num(usePointPrice) == false) {
    			common_msgPopOpen("주문서", "적용할 포인트 금액은 숫자로만 입력해 주십시오.", "", "msgPopup", "N");
    			return;
    		} else if (Number(usePointPrice) > Number(usablePointPrice)) {
    			common_msgPopOpen("주문서", "적용가능 포인트보다 많이 사용하실 수 없습니다.", "", "msgPopup", "N");
    			return;
    		}

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetUsePointModifyOk.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx + "&UsePointPrice=" + usePointPrice,
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
    					getOrderSheetList(multiDelvFlag);
    					closePop('DimDepth1');

    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 슈즈상품권 조회
    	function openUseScash(orderSheetIdx) {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetUseScash.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx,
    			dataType: "text",
    			success: function (data) {
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
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 보유 슈즈상품권 전체 사용
    	function useScashAll() {
    		var useScashPrice = $("#UseScash input[name='UsableScashPrice']").val();
    		useScashPrice = useScashPrice.replace(/(\d)(?=(?:\d{3})+(?!\d))/g, "$1,");

    		$("#UseScash input[name='UseScashPrice']").val(useScashPrice);
    	}

    	function useScash(orderSheetIdx) {
    		var multiDelvFlag = $("#MultiDelvFlag").val();

    		var usableScashPrice = $("#UseScash input[name='UsableScashPrice']").val();
    		var useScashPrice = $("#UseScash input[name='UseScashPrice']").val().replace(/,/g, "");

    		if (useScashPrice.length == 0) {
    			useScashPrice = "0";
    		} else if (only_Num(useScashPrice) == false) {
    			common_msgPopOpen("주문서", "적용할 슈즈상품권 금액은 숫자로만 입력해 주십시오.", "", "msgPopup", "N");
    			return;
    		} else if (Number(useScashPrice) > Number(usableScashPrice)) {
    			common_msgPopOpen("주문서", "적용가능 슈즈상품권보다 많이 사용하실 수 없습니다.", "", "msgPopup", "N");
    			return;
    		}

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetUseScashModifyOk.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx + "&UseScashPrice=" + useScashPrice,
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
    					getOrderSheetList(multiDelvFlag);
    					closePop('DimDepth1');

    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 다중배송시 개별 배송지 입력창 열기
    	function openMultiReceiverInfo(orderSheetIdx) {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetReceiverInfo.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx,
    			dataType: "text",
    			success: function (data) {
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
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 개별 배송지 입력/수정
    	function setMultiReceiverInfo() {
    		var multiDelvFlag = $("#MultiDelvFlag").val();

    		/* 수취인 정보 체크*/
    		if ($("form[name='MultiReceiverInfo'] input[name='ReceiveName']").val() == "") { common_msgPopOpen("", "수취인 이름을 입력하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] input[name='ReceiveName']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] select[name='ReceiveTel1'] option:selected").val() == "") { common_msgPopOpen("주문서", "수취인 전화번호 앞자리를 선택하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] select[name='ReceiveTel1']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] input[name='ReceiveTel23']").val() == "") { common_msgPopOpen("", "수취인 전화번호 가운데와 뒷자리를 입력하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] input[name='ReceiveTel23']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] select[name='ReceiveHP1'] option:selected").val() == "") { common_msgPopOpen("", "수취인 휴대폰번호 앞자리를 선택하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] select[name='ReceiveHP1']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] input[name='ReceiveHP23']").val() == "") { common_msgPopOpen("", "수취인 휴대폰번호 가운데와 뒷자리를 입력하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] input[name='ReceiveHP23']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] input[name='ReceiveZipCode']").val() == "") { common_msgPopOpen("", "수취인 우편번호를 입력하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] input[name='ReceiveZipCode']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] input[name='ReceiveAddr1']").val() == "") { common_msgPopOpen("", "수취인 주소를 입력하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] input[name='ReceiveAddr1']").focus(); return; }
    		else if ($("form[name='MultiReceiverInfo'] input[name='ReceiveAddr2']").val() == "") { common_msgPopOpen("", "수취인 상세주소를 입력하세요.", "", "msgPopup", "N"); $("form[name='MultiReceiverInfo'] input[name='ReceiveAddr2']").focus(); return; }

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetReceiverInfoModifyOk.asp",
    			async: false,
    			data: $("form[name='MultiReceiverInfo']").serialize(),
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
    					getOrderSheetList(multiDelvFlag);
    					closePop('DimDepth1');

    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 배송지 목록
    	function openMyAddress(orderSheetIdx) {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/MyAddressList.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx,
    			dataType: "text",
    			success: function (data) {
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
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	// 배송지 설정
    	function setReceiveInfo(addrType) {
    		// 주문고객 정보와 동일
    		if (addrType == "1") {
    			$("#SameOrderer").prop("checked", true);
    			if ($("#RecentAddress button").length > 0) {
    				$("#RecentAddress button").removeClass("selected");
    			}

    			$("#AddressName").val($("#OrderName").val());
    			$("#ReceiveName").val($("#OrderName").val());
				$("#ReceiveTel1").val($("#OrderTel1").val());
				$("#SReceiveTel1").html($("#OrderTel1").val());
    			$("#ReceiveTel23").val($("#OrderTel23").val());
    			$("#ReceiveHP1").val($("#OrderHP1").val());
    			$("#SReceiveHP1").html($("#OrderHP1").val());
    			$("#ReceiveHP23").val($("#OrderHP23").val());
    			$("#ReceiveZipCode").val($("#OrderZipCode").val());
    			$("#ReceiveAddr1").val($("#OrderAddr1").val());
    			$("#ReceiveAddr2").val($("#OrderAddr2").val());
    		}
    			// 최근 배송지
    		else if (addrType == "2") {
    			$("#SameOrderer").prop("checked", false);
    			$("div.shipping-address div.tab-links button").removeClass("current");
    			$("div.shipping-address div.tab-links button").eq(0).addClass("current");
    			$("#RecentAddress button").removeClass("selected");
    			$("#RecentAddress").show();
    			$("div.shipping-address div.add-address").hide();
    			$("div.shipping-address div.add-address input[name='AddAddress']").prop("checked", false);

    			setMyAddress("btn", '', '0');
    		}
    			// 신규 배송지
    		else if (addrType == "3") {
    			$("#SameOrderer").prop("checked", false);
    			$("div.shipping-address div.tab-links button").removeClass("current");
    			$("div.shipping-address div.tab-links button").eq(1).addClass("current");
    			$("#RecentAddress").hide();
    			$("#RecentAddress button").removeClass("selected");
    			$("div.shipping-address div.add-address").show();
    			$("div.shipping-address div.add-address input[name='AddAddress']").prop("checked", true);

    			$("#AddressName").val("");
    			$("#ReceiveName").val("");
    			$("#ReceiveTel1").val("");
    			$("#ReceiveTel23").val("");
    			$("#ReceiveHP1").val("");
    			$("#ReceiveHP23").val("");
    			$("#ReceiveZipCode").val("");
    			$("#ReceiveAddr1").val("");
    			$("#ReceiveAddr2").val("");

    		}
    	}

    	// 배송지 목록에서 배송지선택시
    	// 배송지버튼 클릭시(listType = btn), 배송지목록 선택시(listType = list)
    	function setMyAddress(listType, orderSheetIdx, num) {
    		if (listType == "btn") {
    			$("#RecentAddress button").removeClass("selected");
    			$("#RecentAddress button.btn_AddressName_" + num).addClass("selected");

    			var addressName = $("#RecentAddress input[name='AddressName_" + num + "']").val();
    			var receiveName = $("#RecentAddress input[name='ReceiveName_" + num + "']").val();
    			var receiveTel = $("#RecentAddress input[name='ReceiveTel_" + num + "']").val();
    			var receiveHP = $("#RecentAddress input[name='ReceiveHP_" + num + "']").val();
    			var receiveZipCode = $("#RecentAddress input[name='ReceiveZipCode_" + num + "']").val();
    			var receiveAddr1 = $("#RecentAddress input[name='ReceiveAddr1_" + num + "']").val();
    			var receiveAddr2 = $("#RecentAddress input[name='ReceiveAddr2_" + num + "']").val();

    		} else {
    			if ($("#MyAddress input[name='MyAddress']:checked").length == 0) {
    				common_msgPopOpen("주문서", "배송지를 선택해 주십시오.", "", "msgPopup", "N");
    				return;
    			}
    			num = $("#MyAddress input[name='MyAddress']:checked").val();
    			var addressName = $("#MyAddress input[name='AddressName_" + num + "']").val();
    			var receiveName = $("#MyAddress input[name='ReceiveName_" + num + "']").val();
    			var receiveTel = $("#MyAddress input[name='ReceiveTel_" + num + "']").val();
    			var receiveHP = $("#MyAddress input[name='ReceiveHP_" + num + "']").val();
    			var receiveZipCode = $("#MyAddress input[name='ReceiveZipCode_" + num + "']").val();
    			var receiveAddr1 = $("#MyAddress input[name='ReceiveAddr1_" + num + "']").val();
    			var receiveAddr2 = $("#MyAddress input[name='ReceiveAddr2_" + num + "']").val();
    		}

    		var receiveTel1 = "";
    		var receiveTel2 = "";
    		var receiveTel3 = "";
    		var splitData = receiveTel.split("-");
    		if (splitData.length == 1) {
    			receiveTel1 = splitData[0];
    		}
    		else if (splitData.length == 2) {
    			receiveTel1 = splitData[0];
    			receiveTel2 = splitData[1];
    		}
    		else if (splitData.length == 3) {
    			receiveTel1 = splitData[0];
    			receiveTel2 = splitData[1];
    			receiveTel3 = splitData[2];
    		}

    		var receiveHP1 = "";
    		var receiveHP2 = "";
    		var receiveHP3 = "";
    		var splitData = receiveHP.split("-");
    		if (splitData.length == 1) {
    			receiveHP1 = splitData[0];
    		}
    		else if (splitData.length == 2) {
    			receiveHP1 = splitData[0];
    			receiveHP2 = splitData[1];
    		}
    		else if (splitData.length == 3) {
    			receiveHP1 = splitData[0];
    			receiveHP2 = splitData[1];
    			receiveHP3 = splitData[2];
    		}


    		if (orderSheetIdx == "") {
    			$("#AddressName").val(addressName);
    			$("#ReceiveName").val(receiveName);
    			$("#ReceiveTel1").val(receiveTel1);
    			$("#ReceiveTel23").val(receiveTel2 + receiveTel3);
    			$("#ReceiveHP1").val(receiveHP1);
    			$("#ReceiveHP23").val(receiveHP2 + receiveHP3);
    			$("#ReceiveZipCode").val(receiveZipCode);
    			$("#ReceiveAddr1").val(receiveAddr1);
    			$("#ReceiveAddr2").val(receiveAddr2);
    			closePop('DimDepth1');

    		} else {
    			var multiDelvFlag = $("#MultiDelvFlag").val();

    			$("form[name='MyAddressInfo'] input[name='OrderSheetIdx']").val(orderSheetIdx);
    			$("form[name='MyAddressInfo'] input[name='AddressName']").val(addressName);
    			$("form[name='MyAddressInfo'] input[name='ReceiveName']").val(receiveName);
    			$("form[name='MyAddressInfo'] input[name='ReceiveTel1']").val(receiveTel1);
    			$("form[name='MyAddressInfo'] input[name='ReceiveTel23']").val(receiveTel2 + receiveTel3);
    			$("form[name='MyAddressInfo'] input[name='ReceiveHP1']").val(receiveHP1);
    			$("form[name='MyAddressInfo'] input[name='ReceiveHP23']").val(receiveHP2 + receiveHP3);
    			$("form[name='MyAddressInfo'] input[name='ReceiveZipCode']").val(receiveZipCode);
    			$("form[name='MyAddressInfo'] input[name='ReceiveAddr1']").val(receiveAddr1);
    			$("form[name='MyAddressInfo'] input[name='ReceiveAddr2']").val(receiveAddr2);

    			$.ajax({
    				type: "post",
    				url: "/ASP/Order/Ajax/OrderSheetReceiverInfoModifyOk.asp",
    				async: false,
    				data: $("form[name='MyAddressInfo']").serialize(),
    				dataType: "text",
    				success: function (data) {
    					var splitData = data.split("|||||");
    					var result = splitData[0];
    					var cont = splitData[1];

    					if (result == "OK") {
    						getOrderSheetList(multiDelvFlag);
    						closePop('DimDepth1');

    					}
    					else if (result == "LOGIN") {
    						PageReload();
    					}
    					else {
    						common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    						return;
    					}
    				},
    				error: function (data) {
    					alert(data.responseText);
    					common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    				}
    			});
    		}
    	}

    	/* 픽업매장 검색창 열기 */
    	function openPickupStore(orderSheetIdx, productCode, sizeCD) {
    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/PickupStoreSearch.asp",
    			async: false,
    			data: "OrderSheetIdx=" + orderSheetIdx + "&ProductCode=" + productCode + "&SizeCD=" + sizeCD,
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
    					$("#DimDepth1").html(cont);
    					openPop('DimDepth1');
    					getPickupStoreList(1);
    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	/* 픽업매장 팝업에서 시도 변경시 구군 목록 가져오기 */
		/*
    	function chg_Sido() {

    		var sido = $("#PickupStore #Sido").val();
    		$("#PickupStore #Gugun option").each(function () {
    			if ($(this).index() > 0) {
    				$(this).remove();
    			}
    		});

    		if (sido != "") {
    			$.ajax({
    				type: "get",
    				url: "/Common/Ajax/GetGugunList.asp",
    				async: false,
    				data: "Sido=" + escape(sido),
    				dataType: "text",
    				success: function (data) {
    					var splitData = data.split("|||||");
    					var result = splitData[0];
    					var cont = splitData[1];

    					if (result == "OK") {
    						if (cont != "") {
    							var sGugun = cont.split(",");

    							for (var i = 0; i < sGugun.length; i++) {
    								$("#PickupStore #Gugun").append("<option value=\"" + sGugun[i] + "\">" + sGugun[i] + "</option>");
    							}
    						}
    					}
    					else {
    						common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					}
    				},
    				error: function (data) {
    					common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    				}
    			});
    		}
    	}
		*/

    	/* 픽업매장 리스트 가져오기 */
    	function getPickupStoreList(page) {

    		$("form[name='PickupStore'] input[name='Page']").val(page);

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/PickupStoreList.asp",
    			async: false,
    			data: $("form[name='PickupStore']").serialize(),
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
						/*
    					var splitData1 = cont.split("^^^^^");
    					var storeList = splitData1[0];
    					var storePaging = splitData1[1];
    					$("#StoreList").html(storeList);
    					$("#StorePaging").html(storePaging);
						*/
    					$("#StoreList").html(cont);

    					if ($("#StoreList input:radio[name='StoreCode']").length > 0) {
    						$("#StoreList input:radio[name='StoreCode']").eq(0).trigger("click");
    					} else {
    						// 기본위치는 슈마커
    						var x = 127.044123;
    						var y = 37.502603;
    						load_Map(x, y, '본사');
    					}
    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	/* 픽업매장 적용시 */
    	function setPickupStore() {
    		if ($("#StoreList input:radio[name='StoreCode']:checked").length == 0) {
    			common_msgPopOpen("주문서", "픽업할 매장을 선택해 주십시오.", "", "msgPopup", "N");
    			return;
    		}

    		if ($("#PickupStore input[name='PickupReceiveName']").val().length == 0) { common_msgPopOpen("", "수령인명을 입력해 주십시오.", "", "msgPopup", "N"); $("#PickupStore input[name='PickupReceiveName']").focus(); return; }
    		if ($("#PickupStore select[name='PickupReceiveHP1'] option:selected").val() == "") { common_msgPopOpen("", "수령인 휴대폰번호 앞자리를 선택하세요.", "", "msgPopup", "N"); $("#PickupStore select[name='PickupReceiveHP1']").focus(); return; }
    		if ($("#PickupStore input[name='PickupReceiveHP23']").val().length == 0) { common_msgPopOpen("", "수령인 휴대폰번호를 입력해 주십시오.", "", "msgPopup", "N"); $("#PickupStore input[name='PickupReceiveHP23']").focus(); return; }

    		var multiDelvFlag = $("#MultiDelvFlag").val();

    		$.ajax({
    			type: "post",
    			url: "/ASP/Order/Ajax/OrderSheetPickupStoreModifyOk.asp",
    			async: false,
    			data: $("form[name='PickupStore']").serialize(),
    			dataType: "text",
    			success: function (data) {
    				var splitData = data.split("|||||");
    				var result = splitData[0];
    				var cont = splitData[1];

    				if (result == "OK") {
    					getOrderSheetList(multiDelvFlag);
    					closePop('DimDepth1');

    				}
    				else if (result == "LOGIN") {
    					PageReload();
    				}
    				else {
    					common_msgPopOpen("주문서", cont, "", "msgPopup", "N");
    					return;
    				}
    			},
    			error: function (data) {
    				alert(data.responseText);
    				common_msgPopOpen("주문서", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
    			}
    		});
    	}

    	/* 결제수단 선택시 */
    	function setPayType() {
    		var payType = $("form[name='OrderForm'] input[name='PayType']:checked").val();
    		$(".paydesc").hide(); //.slideUp("fast");
    		$("#PayDesc_" + payType).slideDown("fast");
    	}

    	/* 주문취소 */
    	function orderCancel() {
			/*
    		if (confirm("주문을 취소 하시겠습니까?")) {
    			<%IF AccessType = "Cart" THEN%>
				location.href = "/ASP/Order/CartList.asp";
    			<%ELSEIF AccessType = "ProductOrder" THEN%>
				location.href = "/ASP/Product/ProductDetail.asp?ProductCode=<%=ProductCode%>";
    			<%END IF%>
        		history.back();
    		}
			*/
    	}

    	/* 결제하기 */
    	function order() {
    		var multiDelvFlag = $("#MultiDelvFlag").val();

    		/* 주문자 정보 체크*/
    		if ($("input[name='OrderName']").val() == "") { common_msgPopOpen("", "주문고객 이름을 입력하세요.", "", "OrderName", "N"); return; }
    		//else if ($("select[name='OrderTel1'] option:selected").val() == "") {common_msgPopOpen("", "주문고객 전화번호 앞자리를 선택하세요.", "", "msgPopup", "N");$("select[name='OrderTel1']").focus();return;}
    		//else if ($("input[name='OrderTel23']").val() == "") {common_msgPopOpen("", "주문고객 전화번호 가운데와 뒷자리를 입력하세요.", "", "msgPopup", "N");$("input[name='OrderTel23']").focus();return;}
			else if ($("select[name='OrderHP1'] option:selected").val() == "") { common_msgPopOpen("", "주문고객 휴대폰번호 앞자리를 선택하세요.", "", "msgPopup", "N"); $("select[name='OrderHP1']").focus(); return; }
			else if ($("input[name='OrderHP23']").val().length < 7) { common_msgPopOpen("", "주문고객 휴대폰번호 가운데와 뒷자리를 입력하세요.", "", "OrderHP23", "N"); return; }
    		else if (only_Num($("input[name='OrderHP23']").val()) == false) { common_msgPopOpen("", "주문고객 휴대폰번호 가운데와 뒷자리를 숫자로만 입력하세요.", "", "OrderHP23", "N"); return; }
			else if (checkEmail($("input[name='OrderEmail']").val()) == false) { common_msgPopOpen("", "주문고객 올바른 이메일 주소를 입력하세요.", "", "OrderEmail", "N"); return; }

    		var errFlag = false;
    		/* 매장픽업일 경우 픽업매장 선택 체크 */
    		$("#OrderSheetList li").each(function () {
    			var delvType = $(this).find("input[name='DelvType']").val();
    			/* 매장픽업일 경우 픽업매장 선택 체크 */
    			if (delvType == "S") {
    				if ($(this).find("input[name='PickupShopCD']").val().length == 0) {
    					common_msgPopOpen("주문서", "픽업매장을 선택해 주십시오.", "", "msgPopup", "N");
    					errFlag = true;
    					return false;
    				}
    			}
    				/* 일반택배 다중배송지일 경우 배송지 입력 체크 */
    			else if (multiDelvFlag == "Y") {
    				if ($(this).find("input[name='ProductReceiveName']").val().length == 0) { common_msgPopOpen("", "배송지 수취인명을 입력해 주십시오.", "", "msgPopup", "N"); errFlag = true; return false; }
    				if ($(this).find("input[name='ProductReceiveHP']").val().length == 0) { common_msgPopOpen("", "배송지 수취인휴대폰번호를 입력해 주십시오.", "", "msgPopup", "N"); errFlag = true; return false; }
    				if ($(this).find("input[name='ProductReceiveZipCode']").val().length == 0) { common_msgPopOpen("", "배송지 수취인 우편번호를 입력해 주십시오.", "", "msgPopup", "N"); errFlag = true; return false; }
    				if ($(this).find("input[name='ProductReceiveAddr1']").val().length == 0) { common_msgPopOpen("", "배송지 수취인 주소를 입력해 주십시오.", "", "msgPopup", "N"); errFlag = true; return false; }
    				if ($(this).find("input[name='ProductReceiveAddr2']").val().length == 0) { common_msgPopOpen("", "배송지 수취인 상세주소를 입력해 주십시오.", "", "msgPopup", "N"); errFlag = true; return false; }
    			}
    				/* 일반택배 단일배송지일 경우 배송지 입력 체크 */
    			else {
    				/* 수취인 정보 체크*/
    				if ($("#ReceiveName").val() == "") { common_msgPopOpen("", "수취인 이름을 입력하세요.", "", "ReceiveName", "N"); errFlag = true; return false; }
    				//else if ($("#ReceiveTel1 option:selected").val() == "") {common_msgPopOpen("", "수취인 전화번호 앞자리를 선택하세요.", "", "msgPopup", "N");$("#ReceiveTel1").focus();errFlag = true; return false;}
    				//else if ($("#ReceiveTel23").val() == "") {common_msgPopOpen("", "수취인 전화번호 가운데와 뒷자리를 입력하세요.", "", "msgPopup", "N");$("#ReceiveTel23").focus();errFlag = true; return false;}
					else if ($("#ReceiveHP1 option:selected").val() == "") { common_msgPopOpen("", "수취인 휴대폰번호 앞자리를 선택하세요.", "", "msgPopup", "N"); $("#ReceiveHP1").focus(); errFlag = true; return false; }
					else if ($("#ReceiveHP23").val().length < 7) { common_msgPopOpen("", "수취인 휴대폰번호 가운데와 뒷자리를 입력하세요.", "", "ReceiveHP23", "N"); errFlag = true; return false; }
    				else if (only_Num($("#ReceiveHP23").val()) == false) { common_msgPopOpen("", "수취인 휴대폰번호 가운데와 뒷자리를 숫자로만 입력하세요.", "", "ReceiveHP23", "N"); errFlag = true; return false; }
    				else if ($("#ReceiveZipCode").val() == "") { common_msgPopOpen("", "수취인 우편번호를 입력하세요.", "", "ReceiveZipCode", "N"); errFlag = true; return false; }
    				else if ($("#ReceiveAddr1").val() == "") { common_msgPopOpen("", "수취인 주소를 입력하세요.", "", "ReceiveAddr1", "N"); errFlag = true; return false; }
    				else if ($("#ReceiveAddr2").val() == "") { common_msgPopOpen("", "수취인 상세주소를 입력하세요.", "", "ReceiveAddr2", "N"); errFlag = true; return false; }
    			}
    		});

    		if (errFlag == true) {
    			return;
    		}

    		/* 결제수단 체크*/
    		if ($("input[name='PayType']:checked").length == 0) { common_msgPopOpen("", "결제수단을 선택하여 주십시오.", "", "msgPopup", "N"); return; }

    		var payType = $("input[name='PayType']:checked").val();
    		var usafeFlag = $("input[name='USAFE_FLAG']").val();
    		var usafePayType = $("input[name='USAFE_PAYTYPE']").val();

    		if (usafeFlag == "Y" && payType == usafePayType) {
    			if ($("input[name='GuaranteeInsurance']:checked").length == 0) {
    				alert("보증보험 발행여부를 선택하세요");
    				common_msgPopOpen("", "결제수단을 선택하여 주십시오.", "", "msgPopup", "N");
    				$("input[name='GuaranteeInsurance']:eq(0)").focus();
    				return;
    			}

    			if ($("input[name='GuaranteeInsurance']:checked").val() == "Y") {
    				if ($("input[name='GuaranteeInsuranceAgreement']:checked").length == 0 || $("input[name='GuaranteeInsuranceAgreement']:checked").val() == "N") {
    					common_msgPopOpen("", "보증보험을 이용하시려면 개인정보 이용에 동의를 하셔야 발행이 됩니다.", "", "msgPopup", "N");
    					$("input[name='GuaranteeInsuranceAgreement']:eq(0)").focus();
    					return;
    				}

    				var USafeYear = $("input[name='USafeYear']").val();
    				if (USafeYear == "") {
    					common_msgPopOpen("", "생년월일을 입력하세요.", "", "msgPopup", "N");
    					$("input[name='USafeYear']").focus();
    					return;
    				}

    				var USafeMonth = $("select[name='USafeMonth'] option:selected").val();
    				if (USafeMonth == "") {
    					common_msgPopOpen("", "생년월일을 입력하세요.", "", "msgPopup", "N");
    					$("select[name='USafeMonth']").focus();
    					return;
    				}

    				var USafeDay = $("select[name='USafeDay'] option:selected").val();
    				if (USafeDay == "") {
    					common_msgPopOpen("", "생년월일을 입력하세요.", "", "msgPopup", "N");
    					$("select[name='USafeDay']").focus();
    					return;
    				}

    				if ($("input:radio[name='USafeSex']:checked").length < 1) {
    					common_msgPopOpen("", "성별을 입력하세요.", "", "msgPopup", "N");
    					$("input:radio[name='USafeSex']:first").focus();
    					return;
    				}

    				if (!isValidDate(USafeYear + USafeMonth + USafeDay)) {
    					common_msgPopOpen("", "생년월일을 정확하게 입력하세요.", "", "msgPopup", "N");
    					$("input[name='USafeYear']").focus();
    					return;
    				}
    			}

    		}

    		/* 개인정보활용동의 체크*/
    		if ($("input[name='AgreementFlag']").is(":checked") == false) { common_msgPopOpen("", "개인정보 이용에 동의하여 주십시오.", "", "msgPopup", "N"); return; }




			//document.OrderForm.submit();






			var formData = $("#OrderForm").serialize();
			$.ajax({
				type		 : "POST",
				url			 : "/ASP/Order/Ajax/OrderAddOK.asp",
				data		 : formData,
				cache		 : false,
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var msg			 = splitData[1];
								var orderCode	 = splitData[2];
								
								if (result == "OK") {
									APP_PopupGoUrl("/ASP/Order/OrderPayment.asp?OrderCode=" + orderCode, '0', '');
								}
								else {
									openAlertLayer('alert', msg, 'closeAlertLayer("alertPop");APP_HistoryBack();', '');
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "주문 처리중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
								return;
				}
			});	


    	}

    	/**
		 * yyyymmdd 형식의 날짜값을 입력받아서 유효한 날짜인지 체크한다.
		 * ex) isValidDate("20070415");
		 */
    	function isValidDate(iDate) {
    		if (iDate.length != 8) {
    			return false;
    		}

    		oDate = new Date();
    		oDate.setFullYear(iDate.substring(0, 4));
    		oDate.setMonth(parseInt(iDate.substring(4, 6)) - 1);
    		oDate.setDate(iDate.substring(6));

    		if (oDate.getFullYear() != iDate.substring(0, 4)
				|| oDate.getMonth() + 1 != iDate.substring(4, 6)
				|| oDate.getDate() != iDate.substring(6)) {

    			return false;
    		}

    		return true;
    	}
	</script>

<!-- Facebook Pixel Code -->
<script>
  fbq('track', 'InitiateCheckout');
</script>
<!-- End Facebook Pixel Code -->


<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs1 = Nothing
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>