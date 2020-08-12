****************************************************************************************************
*	ERRTRAP.PRG			MRP - Error Trapping Program
*
*	17/11/97
*	12/17/2001			Backend Error from the Message_mst (e.g. C0001)		- PKY 
*	12/26/2001			Display Backend Error from the Message_mst			- VS
*						With Specified Variable	(e.g. SQLERROR C0001Product No.)					
*	03/06/2002			Multilanguage Changes								- NP
****************************************************************************************************
#Include 'language.h'
** Declare Local variables

*
STORE SPACE(1) TO mesg_chk,mesg_cd,mflag,sql_error,mess_no
STORE 0 TO ctr, mlen_err5

**Before setting 'Exact' on, save the status of 'Exact'. Reset it in the end

IF SET("exact")='ON'
	mflag='y'
ELSE
	mflag='n'
ENDIF

** Store the length of the Error no./Backend error text
mlen = LEN(mP_err(2))

IF TYPE("mP_err(5)") = "N"		&& Back end error
	mlen_err5 = LEN(ALLT(STR(mP_err(5))))
ENDIF

**Error no. defined in program
IF mlen_err5 =0
	sql_error = mP_err(2)
	DO Locate_error_msg IN Errtrap.prg	&& Reinitialise the array
	STORE .F. TO mP_err
	RETURN
ENDIF

**Finding Constraint no. in the error text
ctr= AT('constraint',mP_err(2))
IF ctr <> 0
	ctr = ctr + 12		&& Move the pointer to '
ENDIF

DO CASE
	CASE ctr <> 0
**Check for Constraint No.
		DO WHILE SUBSTR(mP_err(2),ctr,1) != "'"
			sql_error = sql_error + SUBSTR(mP_err(2),ctr,1)
			ctr = ctr +1
			IF ctr > mlen
				=MESSAGEB(defMsg0251,16,mP_login)
				EXIT
			ENDIF
		ENDDO

**If Error no. found, locate the error message in Message master
		IF LEN(sql_error) <> 0 .AND. ctr =< mlen
			DO Locate_error_msg IN Errtrap.prg
		ENDIF
	OTHERWISE
**Backend Error Trapping
**=MESSAGEB("BACKEND ERROR")
		IF mlen_err5 = 0
			=MESSAGEB(defMsg0252,16,mP_login)
		ELSE
*PKY*	DO Backend_error_msg in Errtrap.prg
*VS*    Display Backend Error from the Message_mst With Specified Variable
			sql_no=AT('SQLERROR',mP_err(2))		
			IF sql_no>0							 
				SQL_ERROR1	= SUBSTR(mP_err(2),sql_no)	
				sql_error	= SUBSTR(sql_error1,10,5)	
				mess_no		= SUBSTR(sql_error1,15)		
			ELSE
				sql_error=RIGHT(mP_err(2),5)			
			ENDI
			
			mRef=.F.
			IF ALLT(SUBSTR(sql_error,1,1))='C'
				mRef=.T.
				i=2
				DO WHILE i <= 5
					IF TYPE(ALLT(SUBSTR(sql_error,i,1)))<>'N'
						mRef=.F.
						EXIT
					ENDIF
					i=i+1
				ENDDO
			ENDIF	

			IF mRef=.T.
				DO Locate_error_msg IN Errtrap.prg
			ELSE
				DO Backend_error_msg IN Errtrap.prg
			ENDIF
		ENDIF
ENDCASE

**Resetting the 'Exact' status
IF mflag = 'y'
	SET EXACT ON
ENDIF

STORE .F. TO mP_err		&& Reinitialise array
RETURN
***************************************************************************************
FUNC Locate_error_msg
**Locating Error no. in Message master
	LOCAL lnOldArea
	lnOldArea = SELECT()				&& Save Current Alias
	sql_error	= ALLTRIM(sql_error)		&& Error No.
	mess_no		= ALLT(mess_no)				&& Specified Variable (e.g. Product)
	SELECT J
	USE messRvew
	GO TOP
	LOCATE FOR msg_c = sql_error
	IF FOUND()
		IF LEN(ALLT(mess_no))=0
			=MESSAGEB(sql_error + ' : ' + msg_1,ICON,mP_login)
		ELSE
			=MESSAGEB(sql_error + ' : ' + ALLT(msg_1)+' '+ mess_no,ICON,mP_login)	&& vivek
		ENDI
	ELSE
		DO Backend_error_msg IN Errtrap.prg
	ENDIF
	SELECT (lnOldArea)	 				&& Select Saved Alias
RETURN
***************************************************************************************
FUNC Backend_error_msg
	IF (LEN(ALLT(sql_error))!=0)	&& Error No. not found in Message Master
		=MESSAGEB(defError_Text+mP_err[2],48,mP_login)
	ELSE
*=MESSAGEB("Error Occured. Reason Unknown."+ CHR(13)+;
*"The Message Text was not found in Message Master. Check in SQL Server Messages "+CHR(13)+;
*"or Code Manual.",48,mP_login)
		=MESSAGEB(defError_Text+mP_err[2],48,mP_login)
	ENDIF
	RETURN
***************************************************************************************

