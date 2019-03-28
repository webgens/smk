<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductReentryAdd.asp - 상품 재입고 알림 등록
'Date		: 2019.01.09
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



DIM ProductCode
DIM RIdx
DIM ProductName
DIM BrandName
DIM tempSizeCD
DIM ProductImage

'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

ProductCode				= sqlFilter(Request("ProductCode"))

IF U_Num = "" THEN
		Response.Write "FAIL|||||로그인 정보가 없습니다."
		Response.End
END IF

IF ProductCode = "" THEN
		Response.Write "FAIL|||||상품 정보가 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


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
		BrandName				= oRs("BrandName")
		ProductImage			= oRs("ProductImage")
		IF ProductImage = "" THEN ProductImage = "/Images/60_noimage.png"
		RIdx					= oRs("RIdx")

		IF RIdx = "" THEN
			oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
			Response.Write "FAIL|||||재입고 알림 정보가 없습니다."
			Response.End
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 상품 정보 입니다."
		Response.End
END IF
oRs.Close

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
				End If
				oRs.MoveNext
				i = i + 1
		Loop
END IF
oRs.Close

Response.Write "OK|||||"
%>
    <!-- PopUp -->
    <div class="area-dim"></div>

	<form name="ReentryForm" id="ReentryForm" method="post">
	<input type="hidden" name="ProductCode" value="<%=ProductCode%>" />
	<input type="hidden" name="RIdx"		value="<%=RIdx%>" />
    <div class="area-pop">
        <!-- 팝업 재입고알림 신청 -->
        <div class="full">
            <div class="tit-pop">
                <p class="tit">재입고알림 신청</p>
                <button type="button" class="btn-hide-pop" onclick="closePop('DimDepth1');">닫기</button>
            </div>

            <div class="container-pop">
                <div class="contents">
                    <div class="wrap-reReception">
                        <div class="ly-items">
                            <div class="thumbNail">
                                <img src="<%=ProductImage%>" alt="<%=ProductName%>">
                            </div>
                            <div class="cont">
                                <span class="brand"><%=BrandName%></span>
                                <span class="line"><%=ProductName%></span>
                            </div>
                        </div>

                        <p class="tit-ty2">품절사이즈 선택</p>
                        <div class="pop-size">
                            <div class="inner">
                                <%
									IF tempSizeCD<>"" THEN
										tempSizeCD = Split(tempSizeCD, ",")
										For x = 0 To UBound(tempSizeCD)
											If Trim(tempSizeCD(x)) <> "" Then
								%>
								<span class="check-style"><input type="radio" id="select<%=tempSizeCD(x)%>" name="Reentry_SizeCD" value="<%=tempSizeCD(x)%>"><label for="select<%=tempSizeCD(x)%>"><span><%=tempSizeCD(x)%></span></label>
                                </span>
								<%
											End If
										Next
									END IF
								%>
                            </div>
                        </div>

                        <p class="tit-ty2">연락처</p>
                        <div class="fieldset ty-col2">
                            <label for="join-phone" class="fieldset-label hidden">연락처 입력</label>
                            <div class="fieldset-row">
                                <span class="select">
                                <select name="Mobile1" title="휴대폰 국번 선택">
									<option value="010">010</option>
									<option value="011">011</option>
									<option value="016">016</option>
									<option value="017">017</option>
									<option value="018">018</option>
									<option value="019">019</option>
                                </select>
                                <span class="value">010</span>
                                </span>
                                <span class="input">
                                <input type="text" id="Mobile2" name="Mobile2" maxlength="8" placeholder="휴대폰 번호를 입력해주세요.">
                            </span>
                            </div>
                        </div>

                        <div class="inf-type1">
                            <p class="tit">알려드립니다.</p>
                            <ul>
                                <li class="bullet-ty1">2주 이내 재입고 없을 경우 본 신청정보는 삭제됩니다.</li>
                                <li class="bullet-ty1">재입고 시 조기품절, 가격변동이 있을 수 있습니다.</li>
                            </ul>
                        </div>

                        <div class="term-agree">
                            <div class="fieldset">
                                <span class="checkbox"><input type="checkbox" id="clause-agree" name="clause-agree" data-allparts="clause-agree"></span>
                                <label for="agreement">개인정보 수집 및 이용에 대한 동의</label>
                                <a href="javascript:PolicyView(26);" class="icon is-notext ico-go"></a>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="btns">
                    <button type="button" class="button ty-red" onclick="Reentry_Insert();">신청하기</button>
                </div>
            </div>
        </div>
        <!-- // 팝업 재입고알림 신청 -->
    </div>
	</form>
    <!-- // PopUp -->


	<script>
		//ajax 이용시 라디오버튼 disabled 처리되는 문제로 추가 (2018.12.18 DJ)
		/*
		var FormRadio = {
			build : function(el){
				if($(el).find('input').is(':disabled')){
					$(el).addClass('is-disabled');
				}
				if($(el).find('input').prop('readonly')){
					$(el).addClass('is-readonly');
				}
				if($(el).find('input').is(':checked')){
					$(el).addClass('is-checked');
				}
			},
			change : function(el){
				var groupName = $(el).attr('name');
				$('[name=' + groupName + ']').parent().removeClass('is-checked');
				$('[name=' + groupName + ']:checked').parent().addClass('is-checked');
			},
			focusin : function(el){
				if($(el).is(':focus')){
					$(el).parent().addClass('is-focus');
				} else {
					$(el).parent().removeClass('is-focus');
				}
			}
		};
		*/

		var FormSelect = {
			build : function(el){
				$('.value', el).text($('option:selected', el).text());
				if($('select', el).is(':disabled')){
					$(el).addClass('is-disabled');
				}
				if($('select', el).prop('readonly')){
					$(el).addClass('is-readonly');
				}
			},
			change : function(el){
				$(el).parent().find('.value').text($('option:selected', el).text());
			},
			focusin : function(el){
				if($(el).is(':focus')){
					$(el).parent().addClass('is-focus');
				} else {
					$(el).parent().removeClass('is-focus');
				}
			}
		};

		$(document).ready(function(){
			formInit();
		});

	</script>


<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>