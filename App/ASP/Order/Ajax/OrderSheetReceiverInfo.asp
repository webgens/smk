<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetReceiverInfo.asp - 주문서 배송지 입력 창 폼 페이지
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

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderSheetIdx

DIM AddressName
DIM ReceiveName
DIM ReceiveTel
DIM ReceiveTel1
DIM ReceiveTel2
DIM ReceiveTel3
DIM ReceiveHP
DIM ReceiveHP1
DIM ReceiveHP2
DIM ReceiveHP3
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM MainFlag

DIM arrTel1
DIM arrHP1
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx			= sqlFilter(Request("OrderSheetIdx"))




IF OrderSheetIdx = "" THEN
		Response.Write "FAIL|||||입력된 주문상품이 없습니다."
		Response.End
END IF


arrTel1	= ARRAY("02", "031", "032", "033", "041", "042", "043", "051", "052", "053", "054", "055", "061", "062", "063", "064", "070", "010", "011", "016", "017", "018", "019")
arrHP1	= ARRAY("010", "011", "016", "017", "018", "019")


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


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
		AddressName			= oRs("AddressName")
		ReceiveName			= oRs("ReceiveName")
		ReceiveTel			= oRs("ReceiveTel")
		ReceiveHP			= oRs("ReceiveHP")
		ReceiveZipCode		= oRs("ReceiveZipCode")
		ReceiveAddr1		= oRs("ReceiveAddr1")
		ReceiveAddr2		= oRs("ReceiveAddr2")
		MainFlag			= oRs("MainFlag")

		IF IsNull(ReceiveName)	THEN ReceiveName = ""
		IF IsNull(ReceiveTel)	THEN ReceiveTel = ""
		IF ReceiveTel <> "" THEN
				IF UBound(SPLIT(ReceiveTel,"-")) = 2 THEN
						ReceiveTel1					 = SPLIT(ReceiveTel, "-")(0)
						ReceiveTel2					 = SPLIT(ReceiveTel, "-")(1)
						ReceiveTel3					 = SPLIT(ReceiveTel, "-")(2)
				ELSEIF UBound(SPLIT(ReceiveTel,"-")) = 1 THEN
						ReceiveTel1					 = SPLIT(ReceiveTel, "-")(0)
						ReceiveTel2					 = SPLIT(ReceiveTel, "-")(1)
						ReceiveTel3					 = ""
				ELSEIF UBound(SPLIT(ReceiveTel,"-")) = 0 THEN
						ReceiveTel1					 = SPLIT(ReceiveTel, "-")(0)
						ReceiveTel2					 = ""
						ReceiveTel3					 = ""
				ELSE
						ReceiveTel1					 = ReceiveTel
						ReceiveTel2					 = ""
						ReceiveTel3					 = ""
				END IF
		END IF


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



Response.Write "OK|||||"
%>					
        <div class="area-pop" id="MultiReceiverInfo">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">배송지 <%IF ReceiveName = "" THEN%>입력<%ELSE%>수정<%END IF%></p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents delivery-info" style="border-bottom:none;">
						<form method="post" name="MultiReceiverInfo">
						<input type="hidden" name="OrderSheetIdx" value="<%=OrderSheetIdx%>" />
						<div class="formfield">
							<div class="fieldset">
								<label for="recieve-name31" class="fieldset-label">받는 사람</label>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="ReceiveName" id="MultiReceiveName" value="<%=ReceiveName%>" placeholder="이름을 입력하세요">
									</span>
								</div>
							</div>
							<div class="fieldset">
								<legend class="hidden">연락처 정보 입력</legend>
								<div class="fieldset ty-col2 pt0">
									<label for="MultiReceiveHP23" class="fieldset-label">휴대폰번호</label>
									<div class="fieldset-row">
										<span class="select">
											<select name="ReceiveHP1" title="휴대폰 국번 선택">
                                                <option value="">선택</option>
												<%FOR i = 0 TO UBOUND(arrHP1)%>
                                                <option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = ReceiveHP1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
												<%NEXT%>
											</select>
											<span class="value"><%=ReceiveHP1%></span>
										</span>
										<span class="input">
											<input type="text" name="ReceiveHP23" id="MultiReceiveHP23" value="<%=ReceiveHP2 & ReceiveHP3 %>" placeholder="휴대폰의 앞 번호와 뒷 번호 입력">
										</span>
									</div>
								</div>
								<div class="fieldset ty-col2 pt0">
									<div class="more-num">
										<label for="MultiReceiveTel23" class="fieldset-label">전화번호</label>
										<span>(선택)</span>
									</div>
									<div class="fieldset-row">
										<span class="select">
											<select name="ReceiveTel1" title="전화번호 국번 선택">
                                                <option value="">선택</option>
												<%FOR i = 0 TO UBOUND(arrTel1)%>
                                                <option value="<%=arrTel1(i)%>"<%IF arrTel1(i) = ReceiveTel1 THEN%> selected="selected"<%END IF%>><%=arrTel1(i)%></option>
												<%NEXT%>
											</select>
											<span class="value"><%=ReceiveTel1%></span>
										</span>
										<span class="input">
											<input type="text" name="ReceiveTel23" id="MultiReceiveTel23" value="<%=ReceiveTel2 & ReceiveTel3 %>" placeholder="전화번호의 앞 번호와 뒷 번호 입력">
										</span>
									</div>
								</div>
							</div>
							<div class="fieldset">
								<label for="MultiReceiveAddr2" class="fieldset-label">배송 주소</label>
								<div class="postnum">
									<button class="search-postnum" type="button" onclick="execDaumPostcode('MultiReceiveZipCode','MultiReceiveAddr1','MultiReceiveAddr2')"><span>우편번호 검색</span></button>
									<div class="fieldset-row delivery-num">
										<span class="input is-expand">
											<input type="text" name="ReceiveZipCode" id="MultiReceiveZipCode" value="<%=ReceiveZipCode%>" placeholder="우편번호" readonly="readonly" />
										</span>
									</div>
								</div>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="ReceiveAddr1" id="MultiReceiveAddr1" value="<%=ReceiveAddr1%>" placeholder="주소 입력" readonly="readonly" />
									</span>
								</div>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="ReceiveAddr2" id="MultiReceiveAddr2" value="<%=ReceiveAddr2%>" placeholder="나머지 주소를 입력해주세요.">
									</span>
								</div>
								<%IF U_NUM <> "" THEN%>
								<div class="fieldset-row">
									<span class="checkbox">
										<input type="checkbox" name="MainFlag" id="MultiMainFlag" value="Y" <%IF MainFlag = "Y" THEN%>checked="checked"<%END IF%> />
									</span>
									<label for="MultiMainFlag">기본배송지로 설정</label>
								</div>
								<%END IF%>
							</div>
						</div>
						</form>
					</div>

                    <div class="btns">
                        <button type="button" onclick="setMultiReceiverInfo()" class="button ty-red">확인</button>
                    </div>
				</div>
			</div>
		</div>

		<script type="text/javascript">
			$('#MultiReceiverInfo .select').each(function (i, el) {
				FormSelect.build(el);
				$(el).find('select').on('change', function () {
					FormSelect.change(this);
				});
				$(el).find('select').on('focus blur click', function () {
					FormSelect.focusin(this);
				});
			});
		</script>
<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>