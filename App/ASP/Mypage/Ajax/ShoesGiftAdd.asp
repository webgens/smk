<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ShoesGiftAdd.asp - 슈즈 상품권 입력 폼
'Date		: 2018.12.10
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
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM CPNO
DIM ProductCost
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


CPNO			 = TRIM(Decrypt(Request.Cookies("CPNO")))
ProductCost		 = TRIM(Decrypt(Request.Cookies("ProductCost")))

IF CPNO = "" OR ProductCost = "" THEN
		Response.Write "FAIL|||||슈즈 상품권 정보가 없습니다. 다시 입력하여 주십시오."
		Response.End
END IF


Response.Write "OK|||||"
%>

        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">슈즈 상품권 등록</p>
                    <button type="button" onclick="common_PopClose('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>
				<form name="addShoesGift" id="addShoesGift" method="post" autocomplete="off">
                <div class="container-pop">
                    <div class="contents">
                        <fieldset class="refund-account" style="background: #f3f3f3; padding: 20px 30px;">
                            <legend class="hidden">슈즈 상품권 등록</legend>
                            <div class="fieldset" style="margin-bottom:0;">
                                <label for="sel_bank" class="fieldset-label" style="float:left;width:100px;height:40px;line-height:40px;margin-bottom:0;">슈즈 상품권 번호</label>
                                 <span style="display:inline-block;height:40px;line-height:40px;"><%=CPNO%></span>
                            </div>
                            <div class="fieldset" style="margin-bottom:0;">
                                <label for="sel_bank" class="fieldset-label" style="float:left;width:100px;height:40px;line-height:40px;margin-bottom:0;">슈즈 상품권 금액</label>
                                <span style="display:inline-block;height:40px;line-height:40px;"><%=FormatNumber(ProductCost, 0)%> 원</span>
                            </div>
                        </fieldset>

                        <div class="inf-type1" style="margin-top:30px;">
                            <p class="tit">슈즈 상품권 전환 규약.</p>
                            <ul>
                                <li class="bullet-ty1">슈즈상품권은 온라인전용 포인트로써 오프라인에서는 사용할 수 없습니다.</li>
                                <li class="bullet-ty1">슈즈상품권으로 전환된 상품권은 사용완료 처리되며, 복구 되지 않습니다.</li>
                                <li class="bullet-ty1">슈즈상품권으로 전환된 금액은 현금으로 환불되지 않습니다.</li>
                                <li class="bullet-ty1">슈즈상품권의 사용유효기한은 등록 후 5년 입니다.</li>
                            </ul>
                        </div>



						<ul class="agreement-list">
							<li>
								<div class="fieldset">
									<span class="checkbox">
										<input type="checkbox" id="Agr" name="Agr" data-allparts="agreement" value="Y">
									</span>
									<label for="Agr">슈즈 상품권 전환 규약 동의</label>
									<a href="#" class="icon is-notext ico-go"></a>
								</div>
							</li>
						</ul>



                    </div>


                    <div class="btns" style="position: fixed;width: 100%;bottom: 0;">
                        <button type="button" onclick="ins_ShoesGift()" class="button ty-red">확인</button>
                        <button type="button" onclick="common_PopClose('DimDepth1')" class="button ty-black">취소</button>
                    </div>
					</form>
                </div>

            </div>
        </div>		<script type="text/javascript">
			$(function () {
				formInit();
			});
		</script>