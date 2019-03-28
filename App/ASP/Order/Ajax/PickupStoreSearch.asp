<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'PickupStoreSearch.asp - 픽업매장 검색 페이지
'Date		: 2018.12.30
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

DIM OrderSheetIdx
DIM ProductCode
DIM SizeCD

DIM OrderName
DIM OrderHP
DIM OrderHP1
DIM OrderHP2
DIM OrderHP3

DIM ReceiveName
DIM ReceiveHP
DIM ReceiveHP1
DIM ReceiveHP2
DIM ReceiveHP3

DIM arrTel1
DIM arrHP1
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx	= sqlFilter(Request("OrderSheetIdx"))
ProductCode		= sqlFilter(Request("ProductCode"))
SizeCD			= sqlFilter(Request("SizeCD"))




IF ProductCode = "" OR SizeCD = "" THEN
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF


arrTel1	= ARRAY("02", "031", "032", "033", "041", "042", "043", "051", "052", "053", "054", "055", "061", "062", "063", "064", "070", "010", "011", "016", "017", "018", "019")
arrHP1	= ARRAY("010", "011", "016", "017", "018", "019")


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


OrderName			= ""
OrderHP				= ""
OrderHP1			= ""
OrderHP2			= ""
OrderHP3			= ""

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
				OrderName			= oRs("Name")
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
		END IF
		oRs.Close
END IF



ReceiveName			= ""
ReceiveHP			= ""
ReceiveHP1			= ""
ReceiveHP2			= ""
ReceiveHP3			= ""


'# 주문서 상품 내역
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
		ReceiveName			= oRs("ReceiveName")
		ReceiveHP			= oRs("ReceiveHP")
		IF IsNull(ReceiveHP) THEN ReceiveHP = ""
		IF ReceiveHP <> "" THEN
				IF UBound(SPLIT(ReceiveHP,"-")) = 2 THEN
						ReceiveHP1					 = SPLIT(ReceiveHP, "-")(0)
						ReceiveHP2					 = SPLIT(ReceiveHP, "-")(1)
						ReceiveHP3					 = SPLIT(ReceiveHP, "-")(2)
				ELSEIF UBound(SPLIT(ReceiveHP,"-")) = 1 THEN
						ReceiveHP1					 = SPLIT(ReceiveHP, "-")(0)
						ReceiveHP2					 = SPLIT(ReceiveHP, "-")(1)
						ReceiveHP3					 = ""
				ELSEIF UBound(SPLIT(ReceiveHP,"-")) = 0 THEN
						ReceiveHP1					 = SPLIT(ReceiveHP, "-")(0)
						ReceiveHP2					 = ""
						ReceiveHP3					 = ""
				ELSE
						ReceiveHP1					 = ReceiveHP
						ReceiveHP2					 = ""
						ReceiveHP3					 = ""
				END IF
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||선택한 주문상품이 없습니다."
		Response.End
END IF
oRs.Close


IF IsNull(ReceiveName) OR ReceiveName = "" THEN
		ReceiveName			= OrderName
		ReceiveHP			= OrderHP
		ReceiveHP1			= OrderHP1
		ReceiveHP2			= OrderHP2
		ReceiveHP3			= OrderHP3
END IF


Response.Write "OK|||||"
%>					
        <!-- 픽업 매장 선택 POP -->
        <div class="area-pop" id="PickupStore">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">픽업 매장 선택하기</p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop pickUp-store">
                    <div class="contents">
						<form name="PickupStore" method="post">
							<input type="hidden" name="Page" value="" />
							<input type="hidden" name="OrderSheetIdx" value="<%=OrderSheetIdx%>" />
							<input type="hidden" name="ProductCode" value="<%=ProductCode%>" />
							<input type="hidden" name="SizeCD" value="<%=SizeCD%>" />
							<input type="hidden" name="Sido" value="" />
							<input type="hidden" name="Gugun" value="" />

						<div class="formfield">
							<div class="fieldset">
								<label for="recieve-name31" class="fieldset-label">수령하시는분</label>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="PickupReceiveName" id="PickupReceiveName" value="<%=ReceiveName%>" placeholder="이름을 입력하세요">
									</span>
								</div>
							</div>
							<div class="fieldset">
								<legend class="hidden">연락처 정보 입력</legend>
								<div class="fieldset ty-col2 pt0">
									<label for="PickupReceiveHP23" class="fieldset-label">휴대폰번호</label>
									<div class="fieldset-row">
										<span class="select">
											<select name="PickupReceiveHP1" title="휴대폰 국번 선택">
                                                <option value="">선택</option>
												<%FOR i = 0 TO UBOUND(arrHP1)%>
                                                <option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = ReceiveHP1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
												<%NEXT%>
											</select>
											<span class="value"><%=ReceiveHP1%></span>
										</span>
										<span class="input">
											<input type="text" name="PickupReceiveHP23" id="PickupReceiveHP23" value="<%=ReceiveHP2 & ReceiveHP3 %>" placeholder="휴대폰의 앞 번호와 뒷 번호 입력">
										</span>
									</div>
								</div>
							</div>
						</div>

                        <div class="area-map">
                            <div class="ly-map" id="map">
                            </div>

                            <span class="selected-store" style="z-index:1">선택매장 : <em>없음</em></span>
                        </div>

                        <div class="store-list">
							<ul id="StoreList">
                            </ul>
                        </div>

						</form>
                    </div>

                    <div class="buttongroup is-expand">
                        <button type="button" onclick="closePop('DimDepth1')" class="button ty-black">취소</button>
                        <button type="button" onclick="setPickupStore()" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			$(function () {
				$('#PickupStore .select').each(function (i, el) {
					FormSelect.build(el);
					$(el).find('select').on('change', function () {
						FormSelect.change(this);
					});
					$(el).find('select').on('focus blur click', function () {
						FormSelect.focusin(this);
					});
				});
			});
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>