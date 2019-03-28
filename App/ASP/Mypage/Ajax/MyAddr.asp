<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyAddr.asp - 배송지 추가/수정
'Date		: 2018.12.07
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM idx

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
DIM RecCnt
DIM AddrType
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

idx				 = sqlFilter(Request("idx"))
AddrType		 = sqlFilter(Request("AddrType"))
IF idx="" THEN idx = 0

arrTel1	= ARRAY("02", "031", "032", "033", "041", "042", "043", "051", "052", "053", "054", "055", "061", "062", "063", "064", "070")
arrHP1	= ARRAY("010", "011", "016", "017", "018", "019")


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

'/****************************************************************************************/
'회원 배송지정보 SELECT START
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MyAddress_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParaminput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

RecCnt	 = oRs.RecordCount
oRs.Close()

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
	.ActiveConnection = oConn
	.CommandType = adCmdStoredProc
	.CommandText = "USP_Front_EShop_MyAddress_Select_By_Idx"

	.Parameters.Append .CreateParameter("@MemberNum",	adInteger, adParaminput, , U_NUM)
	.Parameters.Append .CreateParameter("@Idx",			adInteger, adParaminput, , idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
	AddressName		= oRs("AddressName")
	ReceiveName		= oRs("ReceiveName")
	ReceiveTel		= oRs("ReceiveTel")
	IF IsNull(ReceiveTel) THEN ReceiveTel = ""
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
	ReceiveHP		= oRs("ReceiveHP")
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
	ReceiveZipCode	= oRs("ReceiveZipCode")
	ReceiveAddr1	= oRs("ReceiveAddr1")
	ReceiveAddr2	= oRs("ReceiveAddr2")
	MainFlag		= oRs("MainFlag")

	IF ReceiveTel1 = "" THEN
		ReceiveTel1 = "선택"
	END IF
	IF ReceiveHP1 = "" THEN
		ReceiveHP1 = "선택"
	END IF
END IF
Response.Write "OK|||||"
'/****************************************************************************************/
'회원 배송지정보 SELECT END
'-----------------------------------------------------------------------------------------------------------'
%>

	    <!-- PopUp -->
        <div class="area-dim"></div>

		<form method="post" name="MyAddress" id="MyAddress">
		<input type="hidden" name="Idx" value="<%=Idx%>" />
		<input type="hidden" name="RecCnt" value="<%=RecCnt%>" />
		<input type="hidden" name="AddrType" value="<%=AddrType%>" />
        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">배송지 <%IF idx > 0 THEN%>수정<%ELSE%>추가<%END IF%></p>
                    <button type="button" class="btn-hide-pop" onclick="common_PopClose('DimDepth1')">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents">
                        <fieldset class="area-delivery-addr">
                            <legend class="hidden">배송지명 정보 입력</legend>
                            <div class="fieldset">
                                <label for="join-name" class="fieldset-label">배송지명</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" id="AddressName" name="AddressName" placeholder="배송지명을 입력해주세요." value="<%=AddressName %>" maxlength="10">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="join-birth" class="fieldset-label">받는 분</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" id="ReceiveName" name="ReceiveName" placeholder="받는 분 성함을 입력해주세요." value="<%=ReceiveName %>" maxlength="25">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset ty-col2">
                                <label for="join-phone" class="fieldset-label">연락처</label>
                                <div class="fieldset-row">
                                    <span class="select">
                                        <select name="ReceiveTel1" id="ReceiveTel1" title="휴대폰 국번 선택">
	                                        <option value="">선택</option>
											<%FOR i = 0 TO UBOUND(arrTel1)%>
											<option value="<%=arrTel1(i)%>"<%IF arrTel1(i) = ReceiveTel1 THEN%> selected="selected"<%END IF%>><%=arrTel1(i)%></option>
											<%NEXT%>
                                        </select>
                                        <span class="value"><%=ReceiveTel1%></span>
                                    </span>
                                    <span class="input">
                                        <input type="text" id="ReceiveTel23" name="ReceiveTel23" title="나머지 번호 입력" value="<%=ReceiveTel2 & ReceiveTel3 %>" maxlength="8">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset ty-col2">
                                <label for="join-phone" class="fieldset-label">휴대폰</label>
                                <div class="fieldset-row">
                                    <span class="select">
                                        <select name="ReceiveHP1" id="ReceiveHP1" title="휴대폰 국번 선택">
	                                        <option value="">선택</option>
											<%FOR i = 0 TO UBOUND(arrHP1)%>
											<option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = ReceiveHP1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
											<%NEXT%>
                                        </select>
                                        <span class="value"><%=ReceiveHP1%></span>
                                    </span>
                                    <span class="input">
                                        <input type="text" id="ReceiveHP23" name="ReceiveHP23" title="나머지 번호 입력" value="<%=ReceiveHP2 & ReceiveHP3 %>" maxlength="8">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset ty-col2">
                                <label for="enter-addr" class="fieldset-label">배송 주소</label>
                                <div class="fieldset-row">
                                    <button type="button" class="button ty-black" onclick="execDaumPostcode('MyReceiveZipCode','MyReceiveAddr1')">우편번호 검색</button>
                                    <span class="input">
                                        <input type="text" id="MyReceiveZipCode" name="ReceiveZipCode" title="우편번호" readonly="readonly" value="<%=ReceiveZipCode %>" onclick="execDaumPostcode('MyReceiveZipCode','MyReceiveAddr1')">
                                    </span>
                                    <span class="input is-expand double">
                                        <input type="text" id="MyReceiveAddr1" name="ReceiveAddr1" title="주소 입력" readonly="readonly" value="<%=ReceiveAddr1 %>">
                                    </span>
                                    <span class="input is-expand double">
                                        <input type="text" id="MyReceiveAddr2" name="ReceiveAddr2" title="상세주소 입력" placeholder="나머지 주소를  입력해주세요" value="<%=ReceiveAddr2 %>" maxlength="50">
                                    </span>
									<span class="checkbox">
										<input type="checkbox" id="MainFlag" name="MainFlag" title="기본배송지 지정" value="Y" <%IF MainFlag = "Y" THEN%>checked="checked"<%END IF%>>
									</span>
									<label for="MainFlag" class="lab-keep-addr">기본배송지 지정</label>
                                </div>
                            </div>
                        </fieldset>
                    </div>
                    <div class="btns">
                        <button type="button" class="button ty-red" onclick="chk_MyAddr();">확인</button>
                    </div>
                </div>

            </div>
        </div>
		</form>
	    <!-- // PopUp -->

		<script>


			var formInit = function(){
				$('.checkbox').each(function (i, el) {
					FormCheckbox.build(el);
					$(el).find('input').on('change', function () {
						FormCheckbox.change(this);
						if ($(this).data('allchk') != undefined) {
							FormCheckbox.allchk(this);
						} else if ($(this).data('allparts') != undefined) {
							FormCheckbox.allparts(this);
						}
					});
					$(el).find('input').on('focus blur click', function () {
						FormCheckbox.focusin(this);
					});
				});
				$('.select').each(function(i, el){
					FormSelect.build(el);
					$(el).find('select').on('change', function(){
						FormSelect.change(this);
					});
					$(el).find('select').on('focus blur click', function(){
						FormSelect.focusin(this);
					});
				});
			};

			$(document).ready(function(){
				formInit();
			});

		</script>
<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>