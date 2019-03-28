<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'RefundAccountAdd.asp - 마이페이지 > 회원정보수정 > 계좌정보 입력/수정
'Date		: 2019.01.19
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
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

<%
'/****************************************************************************************'
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

Dim Idx
Dim BankCode
Dim BankName
Dim AccountNum
Dim AccountName
Dim CreateDT

Dim BankInfoArr
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

Idx					= sqlFilter(Request("Idx"))


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_BankCode_Select_For_SelectBox"

END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF oRs.EOF THEN
	oRs.Close : Set oRs = Nothing : oConn.Close : Set oConn = Nothing
	Response.Write "FAIL|||||은행정보가 없습니다."
ELSE
	BankInfoArr = oRs.GetRows()
END IF
oRs.Close

IF Idx <> "" THEN
	Set oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection = oConn
			.CommandType = adCmdStoredProc
			.CommandText = "USP_Front_EShop_Member_RefundAccount_Select_By_MemberNum"

			.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	Set oCmd = Nothing

	IF oRs.EOF THEN
		oRs.Close : Set oRs = Nothing : oConn.Close : Set oConn = Nothing
		Response.Write "FAIL|||||사용자 계좌정보가 일치하지 않습니다."
	ELSE
		Idx				= oRs("Idx")
		BankCode		= oRs("BankCode")
		BankName		= oRs("BankName")
		AccountNum		= oRs("AccountNum")
		AccountName		= oRs("AccountName")
		CreateDT		= oRs("CreateDT")
	END IF
END IF
IF BankName = "" THEN
	BankName = "은행을 선택하세요."
END IF


Response.Write "OK|||||"
%>

		 <!-- PopUp -->
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">환불계좌 <%IF Idx="" THEN%>등록<%ELSE%>수정<%END IF%></p>
                    <button type="button" class="btn-hide-pop" onclick="common_PopClose('DimDepth1');">닫기</button>
                </div>

				<form name="RefundAccountAdd" id="RefundAccountAdd" method="post">
				<input type="hidden" name="Idx" value="<%=Idx%>" />
                <div class="container-pop">
                    <div class="contents">
                        <fieldset class="refund-account">
                            <legend class="hidden">환불계좌 정보 <%IF Idx="" THEN%>등록<%ELSE%>수정<%END IF%></legend>
                            <div class="fieldset">
                                <label for="sel_bank" class="fieldset-label">은행선택</label>
                                <span class="select is-expand">
                                    <select name="BankCode" title="휴대폰 국번 선택" id="BankCode">
                                        <option value="">은행을 선택하세요.</option>
										<%
											FOR i=0 TO Ubound(BankInfoArr,2)
												IF BankInfoArr(2, i) = "N" THEN
										%>
										<option value="<%=BankInfoArr(0, i)%>" <%IF BankInfoArr(0, i) = BankCode THEN Response.Write " selected"%>><%=BankInfoArr(1, i)%></option>
										<%
												END IF
											NEXT
										%>
                                    </select>
                                    <span class="value"><%=BankName%></span>
                                </span>
                            </div>
                            <div class="fieldset">
                                <label for="enter_account_num" class="fieldset-label">계좌번호 입력</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" id="AccountNum" name="AccountNum" maxlength="20" placeholder="-없이 입력해주세요." value="<%=AccountNum%>">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="enter_name" class="fieldset-label">예금주</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" id="AccountName" name="AccountName" maxlength="25" placeholder="예금주명을 정확히 입력해 주세요." value="<%=AccountName%>">
                                    </span>
                                </div>
                            </div>
                        </fieldset>

                        <div class="inf-type1">
                            <p class="tit">알려드립니다.</p>
                            <ul>
                                <li class="bullet-ty1">계좌는 주문자명의 계좌로만 등록이 가능합니다.</li>
                                <li class="bullet-ty1">계좌등록이 안되실 경우, 고객센터 (080-030-2809)로 문의해주세요.</li>
                            </ul>
                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" class="button ty-red" onclick="refundAccountAddOk();">등록하기</button>
                    </div>
                </div>
				</form>

            </div>
        </div>
	    <!-- // PopUp -->


		<script>
			var formInit = function(){
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
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>