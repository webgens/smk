<%
'*****************************************************************************************'
'Common.asp - 기본 함수 페이지
'*****************************************************************************************'


'-----------------------------------------------------------------------------------------'
'암호화/복호화
'Date : 2018.10.05
'-----------------------------------------------------------------------------------------'
FUNCTION Encrypt(val)
		DIM Enc
		DIM EncCode

		If ISNULL(val) OR isEmpty(val) Then val = ""

		SET Enc = Server.createobject("AES.Security")
		EncCode = Enc.Encrypt(val)
		SET Enc = Nothing

		Encrypt = EncCode
END FUNCTION

FUNCTION Decrypt(val)
		DIM Dec
		DIM DecCode : DecCode = ""

		If ISNULL(val) OR isEmpty(val) Then val = ""

		IF LEN(val) > 0 THEN
				SET Dec = Server.createobject("AES.Security")
				DecCode = Dec.Decrypt(val)
				SET Dec = Nothing
		END IF

		Decrypt = TRIM(DecCode)
END FUNCTION




'-----------------------------------------------------------------------------------------'
'URL DECODING
'Date : 2018.10.08
'-----------------------------------------------------------------------------------------'
FUNCTION URLDecode(val)
		DIM aList
		DIM sLength
		DIM fi
		DIM sT
		DIM fDepth
		DIM fY
		DIM sR

		SET aList = CreateObject("System.Collections.ArrayList")
		sLength = LEN(val)

		FOR fi = 1 to sLength
				sT = MID(val, fi, 1)
				IF sT = "%" THEN
						IF fi + 2 <= sLength THEN
								aList.Add CByte("&H" & MID(val, fi + 1, 2))
								fi = fi + 2
						END IF
				ELSE
						aList.Add ASC(sT)
				END IF
		NEXT

		fDepth = 0
		FOR EACH fY IN aList.ToArray()
				IF fY AND &h80 THEN
						IF (fY AND &h40) = 0 THEN
								IF fDepth = 0 THEN Err.Raise 5
								val = val * 2 ^ 6 + (fY AND &h3f)
								fDepth = fDepth - 1
								IF fDepth = 0 THEN
										sR = sR & chrw(val)
										val = 0
								END IF
						ELSEIF (fY AND &h20) = 0 THEN
								IF fDepth > 0 THEN Err.Raise 5
								val = fY AND &h1f
								fDepth = 1
						ELSEIF (fY AND &h10) = 0 THEN
								IF fDepth > 0 THEN Err.Raise 5
								val = fY AND &h0f
								fDepth = 2
						ELSE
								Err.Raise 5
						END IF
				ELSE
						IF fDepth > 0 THEN Err.Raise 5
						sR = sR & chrw(fY)
				END IF
		NEXT
		IF fDepth > 0 THEN Err.Raise 5
		URLDecode = sR
END FUNCTION



'-----------------------------------------------------------------------------------------'
'중복문자체크
'Date : 2018.12.03
'-----------------------------------------------------------------------------------------'
FUNCTION chk_SameChr(ByVal val, ByVal vlen)
		DIM retVal : retVal = True
		DIM fi
		DIM fj : fj = 0
		DIM fb : fb = ""
		DIM fc : fc = ""
		IF TRIM(val) <> "" THEN
				FOR fi = 1 TO LEN(val)
						fc = LCase(MID(val, fi,  1))
						IF fb = "" THEN fb = fc
						IF fc = fb THEN
								fj = fj + 1
						ELSE
								fj = 1
						END IF
	
						IF fj >= CInt(vlen) THEN
								EXIT FOR
						END IF

						fb = fc

				NEXT

				IF fj >= CInt(vlen) THEN
						retVal = False
				ELSE
						retVal = True
				END IF
		END IF

		chk_SameChr = retVal
END FUNCTION




'-----------------------------------------------------------------------------------------'
'나이계산
'Date : 2018.10.16
'-----------------------------------------------------------------------------------------'
FUNCTION GetAge(ByVal val)
		DIM retVal : retVal = ""

		IF val <> "" THEN
			val = LEFT(val, 4)
			IF ISNUMERIC(val) THEN
					RetVal = YEAR(Date) - CDbl(val) + 1
			END IF
		END IF
		GetAge = retVal
END FUNCTION




'-----------------------------------------------------------------------------------------'
'쿠폰 사용처 구분
'Date : 2018.11.23
'-----------------------------------------------------------------------------------------'
FUNCTION GetCouponUseType(ByVal pcFlag, ByVal mobileFlag, ByVal appFlag, ByVal offFlag)

		DIM retVal : retVal = ""

		IF offFlag = "Y" THEN
				IF pcFlag = "Y" OR mobileFlag = "Y" OR appFlag = "Y" THEN
						retVal = "온/오프 통합"
				ELSE
						retVal = "오프라인 전용"
				END IF
		ELSE
				IF pcFlag = "Y" AND mobileFlag = "Y" AND appFlag = "Y" THEN
						retVal = "온라인 전용"
				ELSEIF pcFlag = "Y" AND mobileFlag = "Y" AND appFlag = "N" THEN
						retVal = "PC/모바일 전용"
				ELSEIF pcFlag = "Y" AND mobileFlag = "N" AND appFlag = "Y" THEN
						retVal = "PC/앱 전용"
				ELSEIF pcFlag = "Y" AND mobileFlag = "N" AND appFlag = "N" THEN
						retVal = "PC 전용"
				ELSEIF pcFlag = "N" AND mobileFlag = "Y" AND appFlag = "Y" THEN
						retVal = "모바일/앱 전용"
				ELSEIF pcFlag = "N" AND mobileFlag = "Y" AND appFlag = "N" THEN
						retVal = "모바일 전용"
				ELSEIF pcFlag = "N" AND mobileFlag = "N" AND appFlag = "Y" THEN
						retVal = "앱 전용"
				END IF
		END IF

		GetCouponUseType = retVal
END FUNCTION


'-----------------------------------------------------------------------------------------'
'결제수단
'Date : 2018.11.27
'-----------------------------------------------------------------------------------------'
FUNCTION GetPayType(ByVal payType)

		DIM retVal : retVal = ""


		SELECT CASE payType
				CASE "C" : retVal = "신용카드"
				CASE "B" : retVal = "계좌이체"
				CASE "V" : retVal = "가상계좌"
				CASE "M" : retVal = "휴대폰결제"
				CASE "S" : retVal = "슈마커페이"
				CASE "N" : retVal = "네이버페이"

				CASE ELSE : retVal = payType
		END SELECT

		GetPayType = retVal
END FUNCTION





'-----------------------------------------------------------------------------------------'
'취소/반품/교환 배송비 결제수단
'Date : 2018.12.18
'-----------------------------------------------------------------------------------------'
FUNCTION GetDelvFeeType(ByVal delvFeeType)

		DIM retVal : retVal = ""


		SELECT CASE delvFeeType
				CASE "0" : retVal = "면제"
				CASE "1" : retVal = "슈마커부담"
				CASE "2" : retVal = "동봉"
				CASE "3" : retVal = "계좌이체"
				CASE "4" : retVal = "슈즈상품권차감"
				CASE "5" : retVal = "환불금액차감"
				CASE "6" : retVal = "신용카드결제"
				CASE "7" : retVal = "무료배송쿠폰사용"
				CASE "9" : retVal = "기타"

				CASE ELSE : retVal = delvFeeType
		END SELECT

		GetDelvFeeType = retVal
END FUNCTION




'-----------------------------------------------------------------------------------------'
'매장취소사유
'Date : 2018.01.22
'-----------------------------------------------------------------------------------------'
FUNCTION GetCancelReason(ByVal cancelCode)

		DIM retVal : retVal = ""


		SELECT CASE cancelCode
				CASE "01" : retVal = "재고부족"
				CASE "02" : retVal = "주문취소"

				CASE "03" : retVal = "재고부족(실물없음)"
				CASE "04" : retVal = "제품오염"
				CASE "05" : retVal = "짝발(편족)"
				CASE "06" : retVal = "단체주문건"
				CASE "07" : retVal = "반품미확정(실물없음)"

				CASE ELSE : retVal = cancelCode
		END SELECT

		GetCancelReason = retVal
END FUNCTION




'-----------------------------------------------------------------------------------------'
'창고/매장 구분 가져오기
'Date : 2018.11.28
'shopCode		 : 매장코드
'-----------------------------------------------------------------------------------------'
FUNCTION GetWareHouseType(ByVal shopCode)

		DIM retVal : retVal = ""

		SELECT CASE shopCode
				CASE "006740"	: retVal = "W"
				CASE "006774"	: retVal = "W"
				CASE "006775"	: retVal = "W"
				CASE "009800"	: retVal = "W"
				CASE ELSE		: retVal = "S"
		END SELECT

		GetWareHouseType = retVal
END FUNCTION


'-----------------------------------------------------------------------------------------'
'창고명가져오기
'Date : 2018.01.08
'nameType	 : F : Full, S : Short
'whCode		 : 창고코드
'-----------------------------------------------------------------------------------------'
FUNCTION GetWareHouseName(ByVal nameType, ByVal whCode)

		DIM retVal : retVal = ""

		IF nameType = "F" THEN
				SELECT CASE whCode
						CASE "11" : retVal = "B2B"
						CASE "61" : retVal = "B2C"
						CASE "63" : retVal = "아울렛창고"
						CASE "66" : retVal = "B2C(자사몰)"
						CASE "67" : retVal = "B2C(외부몰)"
				END SELECT
		ELSEIF nameType = "R" THEN
				SELECT CASE whCode
						CASE "11" : retVal = "한솔정상창고(B2B)"
						CASE "61" : retVal = "한솔정상창고(B2C)"
						CASE "63" : retVal = "아울렛창고"
						CASE "66" : retVal = "정상창고(온라인)"
						CASE "67" : retVal = "정상창고(외부몰)"
						CASE "69" : retVal = "조정창고"
						CASE "70" : retVal = "불용창고"
				END SELECT
		ELSE
				SELECT CASE whCode
						CASE "11" : retVal = "B2B"
						CASE "61" : retVal = "B2C"
						CASE "63" : retVal = "아울렛"
						CASE "66" : retVal = "자사몰"
						CASE "67" : retVal = "외부몰"
				END SELECT
		END IF

		GetWareHouseName = retVal
END FUNCTION



'-----------------------------------------------------------------------------------------'
'Circle 가져오기
'Date : 2018.11.15
'grade		 : 점수
'-----------------------------------------------------------------------------------------'
FUNCTION GetCircleGrade(ByVal grade)

		DIM retVal : retVal = ""

		IF IsNull(grade) OR grade = "" THEN
				grade = 0
		ELSE
				grade = CDbl(grade)
		END IF

		IF grade < 0.5 THEN
				retVal = "progress-0"
		ELSEIF grade < 1.0 THEN
				retVal = "progress-05"
		ELSEIF grade < 1.5 THEN
				retVal = "progress-1"
		ELSEIF grade < 2.0 THEN
				retVal = "progress-15"
		ELSEIF grade < 2.5 THEN
				retVal = "progress-2"
		ELSEIF grade < 3.0 THEN
				retVal = "progress-25"
		ELSEIF grade < 3.5 THEN
				retVal = "progress-3"
		ELSEIF grade < 4.0 THEN
				retVal = "progress-35"
		ELSEIF grade < 4.5 THEN
				retVal = "progress-4"
		ELSEIF grade < 5.0 THEN
				retVal = "progress-45"
		ELSE
				retVal = "progress-5"
		END IF

		GetCircleGrade = retVal
END FUNCTION

'-----------------------------------------------------------------------------------------'
'별점 가져오기
'Date : 2019.01.06
'gradeType	 : 점수단위 0.5점
'grade		 : 점수
'-----------------------------------------------------------------------------------------'
FUNCTION GetStarGrade(ByVal grade)

		DIM retVal	: retVal	= "0"
		IF		CDbl(grade) > 4.5	THEN 
				retVal = "50"
		ELSEIF	CDbl(grade) > 4.0	THEN 
				retVal = "45"
		ELSEIF	CDbl(grade) > 3.5	THEN 
				retVal = "40"
		ELSEIF	CDbl(grade) > 3.0	THEN 
				retVal = "35"
		ELSEIF	CDbl(grade) > 2.5	THEN 
				retVal = "30"
		ELSEIF	CDbl(grade) > 2.0	THEN 
				retVal = "25"
		ELSEIF	CDbl(grade) > 1.5	THEN 
				retVal = "20"
		ELSEIF	CDbl(grade) > 1.0	THEN 
				retVal = "15"
		ELSEIF	CDbl(grade) > 0.5	THEN 
				retVal = "10"
		ELSEIF	CDbl(grade) > 0.0	THEN 
				retVal = "05"
		END IF

 		GetStarGrade = retVal
END FUNCTION

'-----------------------------------------------------------------------------------------'
'별점 가져오기
'Date : 2018.11.15
'gradeType	 : 점수단위 F : 1점, H : 0.5점
'grade		 : 점수
'-----------------------------------------------------------------------------------------'
'# FUNCTION GetStarGrade(ByVal gradeType, ByVal grade)
'# 
'# 		DIM retVal : retVal = ""
'# 		DIM fi
'# 
'# 		IF IsNull(grade) OR grade = "" THEN
'# 				grade = 0
'# 		ELSE
'# 				grade = CDbl(grade)
'# 		END IF
'# 
'# 		IF gradeType = "H" THEN
'# 				FOR fi = 0 TO 4
'# 						IF grade >  fi THEN
'# 								retVal = retVal & "<span class=""star-l on""></span>" & vbCrLf
'# 						ELSE
'# 								retVal = retVal & "<span class=""star-l""></span>" & vbCrLf
'# 						END IF
'# 						IF grade >= fi + 1 THEN
'# 								retVal = retVal & "<span class=""star-r on""></span>" & vbCrLf
'# 						ELSE
'# 								retVal = retVal & "<span class=""star-r""></span>" & vbCrLf
'# 						END IF
'# 				NEXT
'# 		ELSE
'# 				FOR fi = 0 TO 4
'# 						IF grade >  fi THEN
'# 								retVal = retVal & "<span class=""star-l on""></span>" & vbCrLf
'# 								retVal = retVal & "<span class=""star-r on""></span>" & vbCrLf
'# 						ELSE
'# 								retVal = retVal & "<span class=""star-l""></span>" & vbCrLf
'# 								retVal = retVal & "<span class=""star-r""></span>" & vbCrLf
'# 						END IF
'# 				NEXT
'# 		END IF
'# 
'# 		GetStarGrade = retVal
'# END FUNCTION


'-----------------------------------------------------------------------------------------'
'메일발송
'Date : 2010.02.22
'-----------------------------------------------------------------------------------------'
SUB MailSend(ByVal FromName, ByVal FromEmail, ByVal ToName, ByVal ToEmail, ByVal MailTitle, ByVal MailContents)

		DIM ObjMail
		DIM iConf
		DIM Flds

		CONST cdoSendUsingPickup = 1   
		SET ObjMail	 = CreateObject("CDO.Message")
		SET iConf	 = CreateObject("CDO.Configuration")
		SET Flds	 = iConf.Fields

		WITH Flds
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup" 
				.Update
		END WITH

		WITH ObjMail
				SET .Configuration = iConf
				.BodyPart.Charset		 = "utf-8"
				.To						 = ToName & "<" & ToEmail & ">"
				.From					 = FromName & "<" & FromEmail & ">"
				.Subject				 = MailTitle
				.HTMLBody				 = MailContents
				.HTMLBodyPart.Charset	 = "utf-8"
				.Send
		END WITH

		SET ObjMail = Nothing
		SET iConf = Nothing
		SET Flds = Nothing

END SUB



'-----------------------------------------------------------------------------------------'
'SMS 발송
'Date :  2018.12.07
'-----------------------------------------------------------------------------------------'
SUB SmsSend(ByVal oConn, ByVal Title, ByVal SendNum, ByVal ReceiveNum, ByVal Msg)

		DIM ContentsByte
		DIM SmsKinds

		ContentsByte		 = calByte(Msg)
		IF ContentsByte > 80 THEN
				SmsKinds	 = 6
		ELSE
				SmsKinds	 = 4
		END IF


		DIM oCmd

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_Msg_Tran_Insert"

				.Parameters.Append .CreateParameter("@Phone_No",	 adVarChar,	 adParamInput,   32,	 Replace(ReceiveNum, "-",""))
				.Parameters.Append .CreateParameter("@Callback_No",	 adVarChar,	 adParamInput,   32,	 Replace(SendNum, "-",""))
				.Parameters.Append .CreateParameter("@Msg_Type",	 adInteger,	 adParamInput,     ,	 SmsKinds)
				.Parameters.Append .CreateParameter("@@Subject",	 adVarChar,	 adParamInput,   40,	 Title)
				.Parameters.Append .CreateParameter("@Message",		 adVarChar,	 adParamInput, 2000,	 Msg)
		
				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

END SUB



'-----------------------------------------------------------------------------------------'
'바이트 수 계산.
'Date : 2017.12.06
'-----------------------------------------------------------------------------------------'
FUNCTION calByte(ByVal val)
		DIM wLen : wLen = 0
		DIM fi
		DIM charAt

		FOR fi = 1 TO LEN(val)
				charAt = ASC(MID(val, fi, 1))
				IF charAt > 0 AND charAt < 255 THEN
				'IF charAt < 0 THEN
						wLen = wLen + 1
				ELSE
						wLen = wLen + 2
				END IF
		NEXT

		calByte = wLen
END FUNCTION



'-----------------------------------------------------------------------------------------'
'스크립트 구문 체킹.
'Date : 2009.08.12
'-----------------------------------------------------------------------------------------'
SUB CheckScript(str)

		DIM checkstr
		DIM msg
		checkstr = str

		IF checkstr <> "" THEN
				DIM sstrarr
				DIM istrarr

				sstrarr = split(LCase(Replace(checkstr, " ", "")), "<script")
				IF UBound(sstrarr) > 0 THEN
						Call AlertMessage("\n script 태그는 입력 할 수 없습니다. \t\n", "history.back();")
						Response.End
				END IF

				istrarr = split(LCase(Replace(checkstr, " ", "")), "<iframe")
				IF UBound(istrarr) > 0 THEN
					Call AlertMessage("\n iframe 태그는 입력 할 수 없습니다. \t\n", "history.back();")
					Response.End
				END IF
		END IF

END SUB



'-----------------------------------------------------------------------------------------'
'입력값 검사
'Date :  2010.10.13
'-----------------------------------------------------------------------------------------'
FUNCTION StringCheck(Str)

		DIM TmpStr : TmpStr = ""

		IF (Str <> "" ) AND (NOT ISNULL(Str)) THEN

				TmpStr = TRIM(STR)
				TmpStr = REPLACE(TmpStr, "'", "''")
				TmpStr = REPLACE(TmpStr, "--", "__")
				TmpStr = REPLACE(TmpStr, ";", "|")
				TmpStr = REPLACE(TmpStr, "%", "")
				TmpStr = REPLACE(TmpStr, "<", "&lt;")
				TmpStr = REPLACE(TmpStr, ">", "&gt;")

		END IF
	
		StringCheck = TmpStr

END FUNCTION



'-----------------------------------------------------------------------------------------'
'날짜포멧 YYYY-MM-DD
'Date : 2011-07-18
'-----------------------------------------------------------------------------------------'
FUNCTION GetDateYMD(ByVal ODate)

		DIM retVal : retVal = ""
		IF LEN(ODate) = 8 THEN
			retVal = LEFT(ODate,4) & "-" & MID(ODate,5,2) & "-" & MID(ODate,7,2)
		END IF
		GetDateYMD = retVal

END FUNCTION
FUNCTION GetDateYMD2(ByVal ODate)

		DIM retVal : retVal = ""
		IF LEN(ODate) = 8 THEN
			retVal = LEFT(ODate,4) & "." & MID(ODate,5,2) & "." & MID(ODate,7,2)
		END IF
		GetDateYMD2 = retVal

END FUNCTION



'-----------------------------------------------------------------------------------------'
'시간포멧 HH-MM-SS
'Date : 2018-01-10
'-----------------------------------------------------------------------------------------'
FUNCTION GetTimeHMS(ByVal OTime)

		DIM retVal : retVal = ""
		IF LEN(OTime) = 6 THEN
			retVal = LEFT(OTime,2) & ":" & MID(OTime,3,2) & ":" & MID(OTime,5,2)
		ELSEIF LEN(OTime) = 4 THEN
			retVal = LEFT(OTime,2) & ":" & MID(OTime,3,2)
		END IF
		GetTimeHMS = retVal

END FUNCTION


'-----------------------------------------------------------------------------------------'
'날짜포멧 YYYY-MM-DD HH:MM
'Date : 2011-07-18
'-----------------------------------------------------------------------------------------'
SUB GetDateYMDHM(ByVal ODate, ByVal OTime, ByRef RDate, ByRef RTime)

	IF LEN(ODate) = 8 AND LEN(OTime) = 6 THEN
			RDate = LEFT(ODate,4) & "-" & MID(ODate,5,2) & "-" & MID(ODate,7,2)
			RTime = LEFT(OTime,2) & ":" & MID(OTime,3,2)
	END IF

END SUB



'-----------------------------------------------------------------------------------------'
'주문상태
'Date : 2011-07-18
'-----------------------------------------------------------------------------------------'
FUNCTION GetOrderState (ByVal oState, ByVal cState1, ByVal cState2)

		DIM retState
		SELECT CASE oState
				CASE "0" : retState = "임시주문"
				CASE "1" : retState = "입금대기"
				CASE "2" : retState = "입금전취소"

				CASE "3"
						IF cState1 = "0" THEN
								IF cState2 = "0" THEN
										retState = "결제완료"
								ELSEIF cState2 = "R" THEN
										retState = "취소신청"
								ELSEIF cState2 = "C" THEN
										retState = "취소접수"
								END IF
						ELSEIF cState1 = "9" THEN
								IF cState2 = "0" THEN
										retState = "교환접수"
								ELSEIF cState2 = "R" THEN
										retState = "교환취소신청"
								ELSEIF cState2 = "C" THEN
										retState = "교환취소접수"
								END IF
						END IF

				CASE "4"
						IF cState1 = "0" THEN
								IF cState2 = "0" THEN
										retState = "상품준비중"
								ELSEIF cState2 = "R" THEN
										retState = "취소신청"
								ELSEIF cState2 = "C" THEN
										retState = "취소접수"
								END IF
						ELSEIF cState1 = "9" THEN
								IF cState2 = "0" THEN
										retState = "교환상품준비중"
								ELSEIF cState2 = "R" THEN
										retState = "교환취소신청"
								ELSEIF cState2 = "C" THEN
										retState = "교환취소접수"
								END IF
						END IF

				CASE "5"
						IF cState1 = "0" THEN
								retState = "배송중"
						ELSEIF cState1 = "8" THEN
								SELECT CASE cState2
										CASE "1" : retState = "반품신청"
										CASE "2" : retState = "반품접수"
										CASE "3" : retState = "반품제품도착"
										CASE "4" : retState = "반품승인"
										CASE "5" : retState = "반품반려"
										CASE "6" : retState = "반품보류"
										CASE "7" : retState = "반품완료"
								END SELECT
						ELSEIF cState1 = "9" THEN
								SELECT CASE cState2
										CASE "0" : retState = "교환배송중"
										CASE "1" : retState = "교환신청"
										CASE "2" : retState = "교환접수"
										CASE "3" : retState = "교환제품도착"
										CASE "4" : retState = "교환승인"
										CASE "5" : retState = "교환반려"
										CASE "6" : retState = "교환보류"
										CASE "7" : retState = "교환반품완료"
								END SELECT
						END IF

				CASE "6"
						retState = "배송완료"

				CASE "7"
						retState = "구매확정"

				CASE "C"
						IF cState1 = "0" THEN
								retState = "취소완료"
						ELSEIF cState1 = "9" THEN
								retState = "교환취소완료"
						END IF
				CASE "S"
						IF cState1 = "0" THEN
								retState = "취소완료[S]"
						ELSEIF cState1 = "9" THEN
								retState = "교환취소완료[S]"
						END IF
				CASE "R"
						IF cState1 = "0" THEN
								retState = "취소완료[R]"
						ELSEIF cState1 = "9" THEN
								retState = "교환취소완료[R]"
						END IF
				CASE "M"
						IF cState1 = "0" THEN
								retState = "취소완료[M]"
						ELSEIF cState1 = "9" THEN
								retState = "교환취소완료[M]"
						END IF
		END SELECT

		GetOrderState = retState

END FUNCTION



'-----------------------------------------------------------------------------------------'
'태그 제거
'Date :  2010.03.07
'-----------------------------------------------------------------------------------------'
FUNCTION RemoveTag(ByVal val)

		DIM strlst, textlst, newcont
		DIM k
		IF val <> "" THEN
				strlst = Split(val,"<")
				newcont = ""
				For k = 0 to UBound(strlst)
						IF strlst(k) <> "" THEN
								textlst = Split(strlst(k),">")

								IF UBound(textlst) < 1 THEN
									newcont = newcont & textlst(0)
								ELSE
									newcont = newcont & textlst(1)
								END IF
						END IF
				NEXT
		END IF

		RemoveTag = newcont

END FUNCTION

	

'-----------------------------------------------------------------------------------------'
'숫자 문자로 변환("0" ++)
'Date :  2010.01.18
'-----------------------------------------------------------------------------------------'
Function MakeZeroChr(ByVal val, ByVal leng)

		DIM chrZero
		DIM k

		IF IsNull(val) OR val = "" THEN val = "0"
		chrZero = CDbl(val)
		FOR k = 1 TO leng
				IF LEN(chrZero) < leng THEN
						chrZero = "0" & chrZero
				END IF
		NEXT

		MakeZeroChr = chrZero

END FUNCTION



'-----------------------------------------------------------------------------------------'
'맨앞 Zero숫자 지우기
'Date :  2017.12.11
'-----------------------------------------------------------------------------------------'
Function FirstZeroDel(ByVal val)

		DIM chrRtn
		DIM str
		DIM k
		DIM j

		IF IsNull(val) THEN val = ""
		chrRtn = ""
		j = LEN(val)
		FOR k = 1 TO j
				IF MID(val,k,1) <> "0" THEN
						chrRtn = MID(val,k)
						EXIT FOR
				END IF
		NEXT

		FirstZeroDel = chrRtn

END FUNCTION


'/*-------------------------------------
'Date :  2018.11.13
'전화번호 마스킹 처리
'-------------------------------------*/
Function MaskTel(ByVal oTel)
	oTel = ChgTelFormat(oTel)

	DIM retVal
	IF oTel <> "" THEN
		DIM sOTel
		oTel = TRIM(oTel)
		sOTel = SPLIT(oTel, "-")
		
		IF UBound(sOTel) = 2 THEN
			retVal = sOTel(0) & "-" & LEFT("**********",LEN(sOTel(1))) & "-" & sOTel(2)
		ELSEIF UBound(sOTel) = 1 THEN
			retVal = LEFT("**********",LEN(sOTel(0))) & "-" & sOTel(1)
		ELSE
			retVal = oTel
		END IF
	END IF

	MaskTel = retVal
End Function


'/*-------------------------------------
'Date :  2018.11.13
'이름 마스킹 처리
'-------------------------------------*/
Function MaskName(ByVal oName)
	IF IsNull(oName) THEN oName = ""
	oName = Replace(oName, " ", "")

	DIM retVal
	IF oName <> "" THEN
		oName = TRIM(oName)
		
		IF LEN(oName) >= 2 THEN
			retVal = LEFT(oName,1) & "*" & MID(oName,3)
		ELSE
			retVal = oName
		END IF
	END IF

	MaskName = retVal
End Function


'/*-------------------------------------
'Date :  2018.11.13
'UserID 마스킹 처리
'-------------------------------------*/
Function MaskUserID(ByVal oUserID)
		IF IsNull(oUserID) THEN oUserID = ""
		oUserID = Replace(oUserID, " ", "")

		DIM retVal
		DIM ANum
		DIM Acc
		IF oUserID <> "" THEN
				oUserID = TRIM(oUserID)
		
				IF LEN(oUserID) >= 2 THEN
						ANum = INSTR(oUserID, "@")
						IF ANum > 0 THEN
								Acc = LEFT(oUserID, ANum - 1)
	
								IF LEN(Acc) = 1 THEN
										retVal = "*" & MID(oUserID, ANum)
								ELSEIF LEN(Acc) = 2 THEN
										retVal = "**" & MID(oUserID, ANum)
								ELSEIF LEN(Acc) > 2 THEN
										retVal = LEFT(Acc, LEN(Acc) - 2) & "**" & MID(oUserID, ANum)
								END IF
						ELSE
								retVal = LEFT(oUserID, LEN(oUserID)-2) & "**"
						END IF
				ELSE
						retVal = oUserID
				END IF
		END IF

		MaskUserID = retVal
End Function

'-----------------------------------------------------------------------------------------'
'QueryString 특수문자 공백으로 치환처리
'Date :  2010.01.13
'-----------------------------------------------------------------------------------------'
FUNCTION sqlFilter(BYVAL str)

		DIM strSearch(11)
		DIM strReplace(11)
		DIM cnt
		DIM retVal : retVal = ""

		IF str = "" OR IsNull(str) THEN 
				retVal = ""
		ELSE
				strSearch(0)	 = "\"
				strSearch(1)	 = "#"
				strSearch(2)	 = "--"
				strSearch(4)	 = ";"
				strSearch(5)	 = "select"
				strSearch(6)	 = "update"
				strSearch(7)	 = "delete"
				strSearch(8)	 = "drop"
				strSearch(9)	 = ""
				strSearch(10)	 = "exec"
				strSearch(11)	 = "dbcc"

				strReplace(0)	 = ""
				strReplace(1)	 = ""
				strReplace(2)	 = ""
				strReplace(4)	 = ""
				strReplace(5)	 = ""
				strReplace(6)	 = ""
				strReplace(7)	 = ""
				strReplace(8)	 = ""
				strReplace(9)	 = ""
				strReplace(10)	 = ""
				strReplace(11)	 = ""
	
				retVal = str
				FOR cnt = 0 TO 11
						retVal = Replace(retVal, LCASE(strSearch(cnt)), strReplace(cnt)) 
				NEXT

		END IF

		sqlFilter = retVal

END FUNCTION



'-----------------------------------------------------------------------------------------'
'자바스크립트 메세지 띄우기
'Date :  2010.01.13
'-----------------------------------------------------------------------------------------'
SUB AlertMessage(ByVal msg, ByVal lurl)

		Response.Write "<script type=""text/javascript"">"
		IF msg <> "" THEN
				Response.Write "alert(""" & msg & """);"
		END IF
		Response.Write lurl 
		Response.Write "</script>"

END SUB


'-----------------------------------------------------------------------------------------'
'자바스크립트 메세지 띄우기
'Date :  2018.12.28
'-----------------------------------------------------------------------------------------'
SUB AlertMessage2(ByVal msg, ByVal lurl)

		Response.Redirect ("/ASP/Error/ErrorPopup.asp?Title=SHOEMARKER&Msg=" & msg & "&Script=" + lurl)

END SUB

'-----------------------------------------------------------------------------------------'
'내용 줄 바꾸기
'Date :  2010.02.09
'-----------------------------------------------------------------------------------------'
FUNCTION ReplaceDetails(ByVal val)

		DIM retVal :retVal = ""
	
		val = trim(val)
		IF val = "" OR IsNull(val) THEN
				retVal = ""
		ELSE
				val		 = Replace (val, chr(13), "<br>")
				val		 = Replace (val, chr(10), "<br>")
				val		 = Replace (val, "><br>", ">")
				val		 = Replace (val, "</b>", "</b><br>")
				retVal	 = Replace (val, chr(96), chr(39))
		END IF

		ReplaceDetails = retVal

END FUNCTION



'/*--------------------------------------------
'전화번호 변경
'Date : 2014.01.03
'--------------------------------------------*/
FUNCTION ChgTel(ByVal val)

		DIM tel
		DIM cnt
		DIM z
		DIM schar
		DIM retVal : retVal = ""

		IF IsNull(val) THEN
				tel = ""
		ELSE
				IF val = "" THEN
						tel = ""
				ELSE
						FOR z = 1 TO LEN(val)
								schar = MID(val, z, 1)
								IF IsNumeric(schar) THEN
										tel = tel & schar
								END IF
						NEXT
				END IF
		END IF


		SELECT CASE LEN(tel)
				CASE 12
						retVal = LEFT(tel, 4) & "-" & MID(tel, 5, 4) & "-" & MID(tel, 9, 4)
				CASE 11
						retVal = LEFT(tel, 3) & "-" & MID(tel, 4, 4) & "-" & MID(tel, 8, 4)
				CASE 10
						IF LEFT(tel, 2) = "02" THEN
								retVal = LEFT(tel, 2) & "-" & MID(tel, 3, 4) & "-" & MID(tel, 7, 4)
						ELSE
								retVal = LEFT(tel, 3) & "-" & MID(tel, 4, 3) & "-" & MID(tel, 7, 4)
						END IF
				CASE 9
						IF LEFT(tel, 2) = "02" THEN
								retVal = LEFT(tel, 2) & "-" & MID(tel, 3, 3) & "-" & MID(tel, 6, 4)
						ELSE
								retVal = LEFT(tel, 3) & "-" & MID(tel, 4, 2) & "-" & MID(tel, 6, 4)
						END IF
				CASE 8
						IF LEFT(tel, 2) = "02" THEN
								retVal = LEFT(tel, 2) & "-" & MID(tel, 3, 2) & "-" & MID(tel, 5, 4)
						ELSE
								retVal = LEFT(tel, 3) & "-" & MID(tel, 4, 1) & "-" & MID(tel, 5, 4)
						END IF
				CASE 7
						retVal = LEFT(tel, 3) & "-" & MID(tel, 4, 4) 
				CASE 6
						retVal = LEFT(tel, 2) & "-" & MID(tel, 3, 4)
				CASE 5
						retVal = LEFT(tel, 1) & "-" & MID(tel, 2, 4)
				CASE ELSE
						retVal = tel
		END SELECT

		ChgTel = retVal

END FUNCTION
	


'-----------------------------------------------------------------------------------------'
'문자앞에 공백 추가 일정자리수 맞추기
'Date : 2017.11.08
'-----------------------------------------------------------------------------------------'
FUNCTION StringLength(ByVal str, ByVal lenNum)

		DIM retVal : retVal = ""
		FOR y = 1 TO lenNum
				retVal = retVal & " "
		NEXT
		retVal = retVal & str
		retVal = RIGHT(retVal, lenNum)
		retVal = REPLACE(retVal, " ", "&nbsp;")
		
		StringLength = retVal

END FUNCTION


'-----------------------------------------------------------------------------------------'
'상품 배지 보여주기 (0:할인률, 1:예약, 2:1+1, 3:사은품, 4:매장픽업)
' 위 우선순위 중 2개만 보여준다.
'Date : 2019.01.05
'-----------------------------------------------------------------------------------------'
FUNCTION ProductBadge(ByVal ProductCode, ByVal DiscountRate, ByVal ReserveFlag, ByVal OPOFlag, ByVal PickupFlag, ByVal GiftCnt)

		DIM retVal : retVal = ""
		Dim BadgeCount : BadgeCount = 0

		'0: 할인률 체크
		If Cint(DiscountRate) > 0 Then
			retVal = "<span class='badge' style='background-color:#282828;'>" & FormatNumber(Cint(DiscountRate), 0) & "%</span>"
			BadgeCount = BadgeCount + 1
		End If

		'1:예약 체크
		If ReserveFlag = "Y" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>예약</span>"
			BadgeCount = BadgeCount + 1
		End If

		'2:1+1 체크
		If BadgeCount < 2 AND OPOFlag = "Y" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>1+1</span>"
			BadgeCount = BadgeCount + 1
		End If

		'3:사은품 체크
		If BadgeCount < 2 AND Cint(GiftCnt) > 0  Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>사은품</span>"
			BadgeCount = BadgeCount + 1
		End If

		'4:매장픽업
		If BadgeCount < 2 AND PickupFlag = "Y" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>매장픽업</span>"
		End If
			
		ProductBadge = retVal

END FUNCTION

'-----------------------------------------------------------------------------------------'
'상품 배지 보여주기 (0:할인률, 1:쿠폰, 2:예약, 3:1+1, 4:사은품, 5:매장픽업)
' 위 우선순위 중 2개만 보여준다.
'Date : 2019.02.19
'-----------------------------------------------------------------------------------------'
FUNCTION ProductBadgeNew(ByVal ProductCode, ByVal DiscountRate, ByVal ReserveFlag, ByVal OPOFlag, ByVal PickupFlag, ByVal GiftCnt, ByVal CouponIdx)

		DIM retVal : retVal = ""
		Dim BadgeCount : BadgeCount = 0

		'0: 할인률 체크
		If Cint(DiscountRate) > 0 Then
			retVal = "<span class='badge' style='background-color:#282828;'>" & FormatNumber(Cint(DiscountRate), 0) & "%</span>"
			BadgeCount = BadgeCount + 1
		End If

		'1: 쿠폰 체크
		if CouponIdx <> "0" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>쿠폰</span>"
			BadgeCount = BadgeCount + 1
		End If

		'2:예약 체크
		If BadgeCount < 2 AND ReserveFlag = "Y" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>예약</span>"
			BadgeCount = BadgeCount + 1
		End If

		'3:1+1 체크
		If BadgeCount < 2 AND OPOFlag = "Y" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>1+1</span>"
			BadgeCount = BadgeCount + 1
		End If

		'4:사은품 체크
		If BadgeCount < 2 AND Cint(GiftCnt) > 0 Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>사은품</span>"
			BadgeCount = BadgeCount + 1
		End If

		'5:매장픽업
		If BadgeCount < 2 AND PickupFlag = "Y" Then
			retVal = retVal & "<span class='badge' style='background-color:#ff201b;'>매장픽업</span>"
		End If
			
		ProductBadgeNew = retVal

END FUNCTION


'-----------------------------------------------------------------------------------------'
'QueryInjection 구문 필터링.
'Date : 2009.08.12
'-----------------------------------------------------------------------------------------'
SUB SQLInjectFilter()

		DIM strTemp : strTemp = CStr(Trim(Request.querystring))
		DIM i
      
		'# URL SQL Inject 금지어를 배열로 만든다.
		'# Dim arrNA : arrNA = Array( "SELECT", "INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "UNION", "DECLARE", "@A", "INT-",  "IS_SRVROLEMEMBER", "IS_MEMBER", "DB_NAME()", "CHAR(", "ISNULL", "VARCHAR", "XP_CMDSHELL", "XP_STARTMAIL", "XP_SENDMAIL", "SP_MAKEWEBTASK", "XP_REGREAD", "XP_REGWRITE", "XP_DIRTREE", "SYSOBJECTS", "SYSDATABASES", "SYSCOLUMNS", "@@VERSION" , "ADDEXTENDEDPROC" , "XPLOG70.DLL" , "SP_DROPEXTENDEDPROC" , "SP_ADDSRVROLEMEMBER", "SP_ADDLOGIN", "D99_CMD", "D99_REG", "D99_TMP", "DIY_TEMPCOMMAND_TABLE", "T_JIAOZHU", "SIWEBTMP", "NB_COMMANDER_TMP", "COMD_LIST", "REG_ARRT",  "JIAOZHU", "XIAOPAN", "DIY_TEMPTABLE", "KILL_KK", "WSCRIPT.SHELL" )
		DIM arrNA : arrNA = Array( "SELECT", "INSERT", "UPDATE", "DELETE", "CREATE", "UNION", "DECLARE", "@A", "INT-",  "IS_SRVROLEMEMBER", "IS_MEMBER", "DB_NAME()", "CHAR(", "ISNULL", "VARCHAR", "XP_CMDSHELL", "XP_STARTMAIL", "XP_SENDMAIL", "SP_MAKEWEBTASK", "XP_REGREAD", "XP_REGWRITE", "XP_DIRTREE", "SYSOBJECTS", "SYSDATABASES", "SYSCOLUMNS", "@@VERSION" , "ADDEXTENDEDPROC" , "XPLOG70.DLL" , "SP_DROPEXTENDEDPROC" , "SP_ADDSRVROLEMEMBER", "SP_ADDLOGIN", "D99_CMD", "D99_REG", "D99_TMP", "DIY_TEMPCOMMAND_TABLE", "T_JIAOZHU", "SIWEBTMP", "NB_COMMANDER_TMP", "COMD_LIST", "REG_ARRT",  "JIAOZHU", "XIAOPAN", "DIY_TEMPTABLE", "KILL_KK", "WSCRIPT.SHELL" )
		
		'# Loop를 돌리면서 SQL Inject 금지어가 있는지 검사
		FOR i = LBOUND( arrNA ) TO UBOUND( arrNA )
				'# 금지어가 있으면 공백으로 치환
				IF INSTR( UCASE(strTemp), TRIM(UCase(arrNA(i))) ) > 0 THEN
						Response.Clear

						Call AlertMessage("불법적인 내용을 포함한 접근으로 인하여 해당 페이지의 접속을 차단합니다.", "history.back();")
						Response.End

						EXIT FOR
				End IF
		NEXT
   
END SUB

SUB SQLInjectSearch(StrValue)

		IF StrValue <> "" THEN
    
				DIM i
				'# URL SQL Inject 금지어를 배열로 만든다.
				DIM arrNA : arrNA = Array( "SELECT", "INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "UNION", "DECLARE", "@A", "INT-",  "IS_SRVROLEMEMBER", "IS_MEMBER", "DB_NAME()", "CHAR(", "ISNULL", "VARCHAR", "XP_CMDSHELL", "XP_STARTMAIL", "XP_SENDMAIL", "SP_MAKEWEBTASK", "XP_REGREAD", "XP_REGWRITE", "XP_DIRTREE", "SYSOBJECTS", "SYSDATABASES", "SYSCOLUMNS", "@@VERSION" , "ADDEXTENDEDPROC" , "XPLOG70.DLL" , "SP_DROPEXTENDEDPROC" , "SP_ADDSRVROLEMEMBER", "SP_ADDLOGIN", "D99_CMD", "D99_REG", "D99_TMP", "DIY_TEMPCOMMAND_TABLE", "T_JIAOZHU", "SIWEBTMP", "NB_COMMANDER_TMP", "COMD_LIST", "REG_ARRT",  "JIAOZHU", "XIAOPAN", "DIY_TEMPTABLE", "KILL_KK", "WSCRIPT.SHELL" )
		

				'#Loop를 돌리면서 SQL Inject 금지어가 있는지 검사
				FOR i = LBOUND( arrNA ) TO UBOUND( arrNA )
						'# 금지어가 있으면 해당 금지어 선택
						IF INSTR( UCASE(StrValue), TRIM(UCase(arrNA(i))) ) > 0 THEN
								Call AlertMessage("불법적인 내용을 포함한 접근으로 인하여 해당 페이지의 접속을 차단합니다. [" & arrNA(i) & "]", "history.back();")
								Response.End

								EXIT FOR
						END IF
				NEXT
   
		END IF

END SUB

FUNCTION SQLInjectReplace(StrValue)
		
		DIM i
		DIM returnStr  : returnStr = StrValue
		'# URL SQL Inject 금지어를 배열로 만든다.
		DIM arrNA : arrNA = Array( "SELECT", "INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "UNION", "DECLARE", "@A", "INT-",  "IS_SRVROLEMEMBER", "IS_MEMBER", "DB_NAME()", "CHAR(", "ISNULL", "VARCHAR", "XP_CMDSHELL", "XP_STARTMAIL", "XP_SENDMAIL", "SP_MAKEWEBTASK", "XP_REGREAD", "XP_REGWRITE", "XP_DIRTREE", "SYSOBJECTS", "SYSDATABASES", "SYSCOLUMNS", "@@VERSION" , "ADDEXTENDEDPROC" , "XPLOG70.DLL" , "SP_DROPEXTENDEDPROC" , "SP_ADDSRVROLEMEMBER", "SP_ADDLOGIN", "D99_CMD", "D99_REG", "D99_TMP", "DIY_TEMPCOMMAND_TABLE", "T_JIAOZHU", "SIWEBTMP", "NB_COMMANDER_TMP", "COMD_LIST", "REG_ARRT",  "JIAOZHU", "XIAOPAN", "DIY_TEMPTABLE", "KILL_KK", "WSCRIPT.SHELL" )
			
		'# Loop를 돌리면서 SQL Inject 금지어가 있는지 검사
		FOR i = LBOUND( arrNA ) TO UBOUND( arrNA )
				'# 금지어가 있으면 공백으로 치환
				IF INSTR( UCASE(returnStr), TRIM(UCase(arrNA(i))) ) > 0 THEN
						returnStr = REPLACE(UCASE(returnStr), TRIM(UCase(arrNA(i))), "" )
				END IF
		NEXT
   
		SQLInjectReplace = returnStr

END FUNCTION


'# 모든 페이지에서 SQL Inject 처리 함수 호출
Call SQLInjectFilter
%>