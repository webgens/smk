<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'SnsSDupInfoView.asp - SNS 정회원 통합 - 기존 ID 검색
'Date		: 2018.12.20
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
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절


DIM RecCnt

DIM SDupInfo
DIM ParentSDupInfo
DIM JoinType
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

JoinType = Decrypt(Request.Cookies("JoinType"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

'/****************************************************************************************/
'회원 ID정보 SELECT START
'-----------------------------------------------------------------------------------------------------------'
IF JoinType = "U" THEN
	SDupInfo = Decrypt(Request.Cookies("SDupInfo"))

	Set oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Member_Select_By_sDusInfo"

		.Parameters.Append .CreateParameter("@@sDupInfo",	 adVarChar, adParaminput, 64, SDupInfo)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	Set oCmd = Nothing
ELSE
	ParentSDupInfo = Decrypt(Request.Cookies("ParentSDupInfo"))

	Set oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Member_Select_By_ParentsDusInfo"

		.Parameters.Append .CreateParameter("@@sDupInfo",	 adVarChar, adParaminput, 64, ParentSDupInfo)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	Set oCmd = Nothing
END IF

RecCnt	 = oRs.RecordCount

IF oRs.EOF THEN
	oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	Response.Write "FAIL|||||해당 정보가 없습니다."
	Response.End
ELSE
	Response.Write "OK|||||"
END IF
'/****************************************************************************************/
'회원 배송지정보 SELECT END
'-----------------------------------------------------------------------------------------------------------'
%>

	<form method="post" name="MyIdCombine" id="MyIdCombine">
	<input type="hidden" name="JoinType" id="JoinType" value="<%=JoinType%>" />
	<input type="hidden" name="SDupInfo" id="SDupInfo" value="<%=SDupInfo%>" />
	<input type="hidden" name="ParentSDupInfo" id="ParentSDupInfo" value="<%=ParentSDupInfo%>" />
    <!-- PopUp -->
    <section class="wrap-pop">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">사용자 계정통합</div>
                    <button class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty2">
                    <!-- 팝업 스타일 변경으로 'mypage-ty2'클래스 명 추가 -->
                    <div class="contents">
                        <div class="wrap-mtom">
                            <div class="reason">
                                <span class="select">
									<select id="" name="CombineID" title="계정선택">
										<option value="">계정 선택</option>
<%
	i=1
	Do While Not oRs.EOF
%>
										<option value='<%=oRs("UserID")%>'>회원ID : <%=oRs("UserID")%> / 성명 : <%=oRs("Name")%> <%IF JoinType <> "U" THEN Response.Write " / 보호자성명 : "&oRs("ParentName") %></option>
<%
		i = i + 1
		oRs.MoveNext
	Loop
%>
									</select>
									<span class="value">계정 선택</span>
                                </span>
                            </div>

                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" onclick="chk_MyIdCombine();" class="button ty-red">확인</button>
                        <button type="button" onclick="common_PopClose('DimDepth1');" class="button ty-black">취소</button>
                    </div>
                </div>
            </div>
        </div>
    </section>
    <!-- // PopUp -->
	</form>


	<script>
		//ajax 이용시 라디오버튼 disabled 처리되는 문제로 추가 (2018.12.18 DJ)							
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

			if ($('#remaintime').get(0) != undefined) {
				setInterval(timedeal, 1);
			}
		});
	</script>

<%
oRs.Close
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>