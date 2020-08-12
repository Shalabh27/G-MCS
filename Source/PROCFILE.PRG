*************************************************************************************************************************
*	Procfile.prg		Functions for date manipulation for DHS-MCS Standard Version
*	PKY & NP
*	Date				07/23/2002
*	Modified			10/09/2002		Fucntion to get default section code				- NP
*	Modified			05/12/2003		Change Fucntion to get default section code			- PKY
*	Modified			07/17/2003		Add Function SplitXLS to break up Excel File		- SA
*										output into multiple Files when rows > 16334
*	Modified			09/02/2003		Add messagebox after single file creation 			- SA
*										in SplitXls (Merging SEWT changes)
*						09/03/2003		Added function to check file name & path validation - NP
* 	Modified			10/07/2003		Error Handler Procedure included					- NP
*	Modified			12/15/2003		Added function FTPPut() and ConnectFTP() for 		- SK
*										transferring file to FTP Server
*	Modified			12/26/2003		Changed 'Put' command of FTP						- SK
*										(As discussed with Suzuki San)
*	Modified			12/31/2003		Added 'Wait Clear' before Messagebox in SplitXls	- SA
*************************************************************************************************************************
*Function for generating a blank date

#INCLUDE 'language.h'
FUNCTION blankCharDate
	LOCAL Cblankdate,cCurrSetting

	cCurrSetting=SUBSTR(mP_CurrDateFormat,1,1)

	DO CASE
		CASE cCurrSetting='A'		  				&&	mm/dd/yyyy
			Cblankdate = '  /  /    '

		CASE cCurrSetting='Y'						&&	yyyy/mm/dd
			Cblankdate = '    /  /  '

		CASE cCurrSetting='B'		  				&&	dd/mm/yyyy
			Cblankdate = '  /  /    '

		OTHERWISE									&&	Return ymd
			Cblankdate = '    /  /  '
	ENDCASE

	RETURN Cblankdate
ENDFUNC

*Set dateformats for VFP and SQL Server
*The function sets the dateformat as requested
*Pass nothing in the function to reset the original setting.
FUNCTION resetDateFormat
	LPARAMETERS cDateFormat

	cDateFormat=IIF(PARAMETERS()=0, mP_CurrDateFormat, cDateFormat)
	cDateFormat=SUBSTR(cDateFormat, 1, 1)		&&	Extract the first letter - somebody might type wrong spelllings

	mcom=''
	DO CASE										&&	Set date format for VFP and SQL Server
		CASE cDateFormat='A'
			SET DATE AMERICAN
			mP_sqlDateOption='101'
			mcom="SET DATEFORMAT mdy"

		CASE cDateFormat='Y'
			SET DATE YMD
			mP_sqlDateOption='111'
			mcom="SET DATEFORMAT mdy"

		CASE cDateFormat='B'					&& Set to British
			SET DATE BRITISH
			mP_sqlDateOption='103'
			mcom="SET DATEFORMAT mdy"

		OTHERWISE								&&	Return ymd
			SET DATE YMD
			mP_sqlDateOption='111'
			mcom="SET DATEFORMAT mdy"
	ENDCASE

	mP_ret=SQLEXEC(mP_handle, mcom)
	IF mP_ret<=0
		=AERROR(mP_err)
		DO Errtrap
		RETURN .F.
	ENDIF
	RETURN .T.
ENDFUNC


*Function for change British date to American

FUNCTION british_to_American
	LPARAMETERS cDate
	IF PARAMETERS()=0
		RETURN NULL
	ENDIF

	cCurrSetting=SUBSTR(mP_CurrDateFormat,1,1)
	LOCAL tempDate
	tempDate = '#'

	IF ISNULL(cDate)
		RETURN NULL
	ENDI

	IF TYPE('cDate')= 'D'		&& if param date type convert to char
		cDate=DTOC(cDate)
	ENDIF

	IF TYPE('cDate')<> 'C'		&& if not char type return blank char date in american
		tempDate=NULL
	ENDIF

	IF TYPE('Cdate') = 'C' AND LEN(ALLT(cDate)) < 10
		tempDate=NULL
	ENDI

	IF !ISNULL(tempDate)
		cDate=ALLT(cDate)

		IF cCurrSetting = 'B'						&&	dd/mm/yyyy
			tempDate = SUBSTR(cDate,4,2) + '/'+ SUBSTR(cDate,1,2) + '/' + SUBSTR(cDate,7,4)		&& Converted to dd/mm/yyyy
		ELSE
			tempDate = cDate							&& If not British Return without change
		ENDIF
	ENDI

	RETURN tempDate
ENDFUNC

FUNCTION def_section
	PARAMETERS lc_sectCur
&& Assumption :- Section cursor must have field 'section_c'
	LOCAL merror, mPos
	* Check for valid param
	IF EMPTY(lc_sectCur) OR !USED(lc_sectCur)  && No Param Passed
		merror = .T.
	ENDI

	IF !merror
		* Creating base section cursor
		mcom = "SELECT section_c FROM &mp_OwnerName defsection_mst WHERE default_section in ('1','4')"
		mRet = SQLEXEC(mP_handle,mcom,"base_SectCur")
		IF mRet <= 0
			=AERROR(mP_err)
			DO Errtrap
			merror = .T.
		ELSE
			IF RECCOUNT('base_sectCur') = 0
				merror = .T.
			ENDIF
		ENDI
	ENDIF

	IF !merror
		* moving pointer to base section
		SELECT (lc_sectCur)
		IF RECCOUNT() = 0
			merror = .T.
		ENDI
	ENDIF

	IF !merror
		IF !EOF()		&& if not end of file
			mPos = RECNO()
		ELSE
			mPos = -1
		ENDI
		GO TOP
		LOCATE FOR section_c = base_sectCur.section_c
		IF !FOUND()
			merror = .T.
			IF mPos = -1
				GO BOTTOM
			ELSE
				GOTO mPos
			ENDI
		ENDI
	ENDIF
	IF USED('base_SectCur')
		USE IN base_sectCur
	ENDI
	RETURN !merror
ENDFUNC

* - RD - 10 Feb 2003
*   This class is used in 'MRPFC720.SCX'
* - Purpose : Dynamically adding checkbox with intrinsic
* -     PEM (Property/Event/Method)
* -     Override the When/valid event.

DEFINE CLASS gCheck AS CHECKBOX

	VISIBLE = .T.

	PROCEDURE VALID
		RETURN THISFORM.RowValid( THIS )

	PROCEDURE WHEN
		RETURN THISFORM.RowWhen( THIS )

ENDDEFINE

* Function to break Excel File in Parts (when data is more than 16334 rows)
* Assumption : Parameters to the function will always be valid

FUNCTION  SplitXls
	LPARAMETERS lcFileName, lcCursorName

	IF PARAMETERS()=0
		RETURN
	ENDI

	lcFileName = ALLT(lcFileName)
	SELE (lcCursorName)

	lnRec = RECCOUNT()
	lnPage = 0
	lnctr = 1
	lnFileStr = ''

	** If rows more than 16334 break file in 16000 rows/file
	IF lnRec > 16334
		DO WHILE lnRec > lnPage

			lnPage = lnPage + 16000
			lcFileName1 = JUSTPATH(lcFileName) + '\' + JUSTSTEM(lcFileName) + '-' + PADL(ALLT(STR(lnctr)),2,'0') + '.xls'

			mHandleRet = CheckOpen(lcFileName1)
			IF mHandleRet < 0
				FCLOSE(mHandleRet)
				WAIT CLEAR
				DO Errtrap
				RETURN 0
			ELSE
				FCLOSE(mHandleRet)
			ENDIF
			COPY TO (lcFileName1) TYPE XLS FOR BETWEEN( RECN() , lnPage - 16000 + 1 , lnPage )
			lnctr = lnctr + 1

			lnFileStr = lnFileStr + JUSTFNAME(lcFileName1) + ", "
		ENDDO
		WAIT CLEAR
		MESSAGEBOX(defMsg0360 + " ( " + JUSTPATH(lcFileName) + " )." + CHR(13) +;
			SUBSTR(lnFileStr,1,LEN(lnFileStr)-2) ,64,mp_login)
	ELSE
		**	If Single File Creating...
		mHandleRet = CheckOpen(lcFileName)
		IF mHandleRet < 0
			WAIT CLEAR
			DO Errtrap
			RETURN 0
		ELSE
&& if file open then close
			FCLOSE(mHandleRet)
		ENDIF
		COPY TO (lcFileName) TYPE XLS
		WAIT CLEAR
		=MESSAGEBOX(defMsg0316,64,mp_login)
	ENDIF
	WAIT CLEAR
ENDFUNC

FUNCTION CheckOpen
	LPARAMETERS FName
	mFileHandle = -1
	mFound = .F.

	** locate specical characters in file name
	mFileName = ALLTRIM(JUSTFNAME(FName))
	FOR ctr = 1 TO LEN(mFileName )
		IF INLIST(SUBSTR(mFileName,ctr,1),'\','/','?',':','*','<','>','|','"')
			mFound = .T.
			mP_err[2]='E0399'
			EXIT
		ENDIF
	ENDFOR

	IF !mFound
		oldError = ON('ERROR')

		ON ERROR mP_err[2]='E0399'
		* Check whether file is open somewhere.
		IF FILE(FName)  						&& Does file exist?
			mFileHandle = FOPEN(FName,12)    	&& If so, open read-write
			mP_err[2]='E0362'					&& Cannot Access File
		ELSE
			mFileHandle = FCREATE(FName)  		&& If not create it
			mP_err[2]='E0399'					&& Cannot Create File
		ENDIF
		ON ERROR
	ENDIF
	RETURN mFileHandle
ENDFUNC


FUNCTION ValidatePath
	LPARAMETERS pathFName,checkPath,ConfirmOverwrite

	IF PARAMETERS() < 3
		MESSAGEBOX("Invalid no of Parameters",48,mp_login)
	ENDIF

	mFileHandle = -1
	
	IF checkpath = .t.
		IF LEN(ALLT((pathFName))) = 0 OR EMPTY(JUSTPATH(pathFName))
			mP_err[2]='E0233'
			DO Errtrap
			RETURN .F.
		ENDIF

		IF !DIRECTORY(JUSTPATH(pathFName))
			mP_err[2]='E0220'
			DO Errtrap
			RETURN .F.
		ENDIF
	ENDIF

	mFileName = ALLTRIM(JUSTFNAME(pathFName))
	FOR ctr = 1 TO LEN(mFileName )
		IF INLIST(SUBSTR(mFileName,ctr,1),'\','/','?',':','*','<','>','|','"')
			mP_err[2]='E0399'
			DO Errtrap
			RETURN .F.
		ENDIF
	ENDFOR

	IF ConfirmOverwrite = .t.
		IF FILE(pathFName) AND ;
				MESSAGEBOX(defMsg0051,4+32,mp_login)=7
			mP_err[2]='E0218'	&&  Operation is cancelled by the user
			DO Errtrap
			RETURN .F.
		ENDIF
	ENDIF
	
	IF FILE(pathFName)
		mFileHandle = FOPEN(pathFName,12)    	&& If so, open read-write
		mP_err[2]='E0362'						&& Cannot Access File
	ELSE
		mFileHandle = FCREATE(pathFName)  		&& If not create it
		mP_err[2]='E0399'						&& Cannot Create File
	ENDIF

	IF mFileHandle < 0
		WAIT CLEAR
		DO Errtrap
		RETURN .F.
	ELSE
		FCLOSE(mFileHandle)
		ERASE (pathFName)
	ENDIF

	RETURN .T.
ENDFUNC

* Error Handler Procedure called by any child form of Cleascre or Cleascre itself - NP - 10-07-2003
PROCEDURE errhand
	LPARAMETERS lErr
	mP_FormErr = lErr
	IF lErr=1 AND menucur->run_local_prg = '1'
		IF TYPE('mP_CalledForm') <> 'C'
			mP_CalledForm = ''
		ENDIF

		IF SUBSTR(mP_CalledForm,1,3) <> 'MRP'
			MESSAGEBOX(ALLTRIM(MESSAGE()) + CHR(13) + IIF(menucur->run_local_prg = '1', defMsg0367, ''), 64, mp_login)
		ELSE
			MESSAGEBOX(defError + ' ' + ALLTRIM(STR(ERROR())) + ' : ' +ALLTRIM(MESSAGE()), 16, mp_login)
		ENDIF
	ELSE
		MESSAGEBOX(defError + ' ' + ALLTRIM(STR(ERROR())) + ' : ' +ALLTRIM(MESSAGE()), 16, mp_login)
	ENDIF
	RETURN
ENDPROC

*...................................................................................................
FUNCTION FTPPut
PARAMETERS cHost, nPort_No, cUser, cPassword, cSource, cTrans_ID, cAprf, nXFerType
*...................................................................................................
*:  Usage : FTPPut ('FTP_IP_Address', 'userid', 'password', 'source_file', 'target_file', 1 | 2)
*:  Where : cHost     = Host computer IP address or name
*:          nPort_No  = Port No (If not specified, default (usually 21) will be used)
*:          cUser     = user name - anonymous may be used
*:          cPassword = password
*:          cSource   = source file name (Local)
*:          cTrans_ID = Transmission Place ID on Remote site (Path on FTP site)
*:          cAprf	  = Destination File on remote site (File name on FTP site)
*:          nXFerType = 1 (default) for ascii, 2 for binary
*...................................................................................................
DECLARE INTEGER InternetOpen IN wininet.DLL;
   STRING  sAgent, INTEGER lAccessType, STRING  sProxyName, STRING  sProxyBypass, STRING  lFlags

DECLARE INTEGER InternetCloseHandle IN wininet.DLL INTEGER hInet

DECLARE INTEGER InternetConnect IN wininet.DLL;
   INTEGER hInternetSession, STRING lcHost, INTEGER nServerPort, STRING lcUser, STRING lcPassword,;
   INTEGER lService, INTEGER lFlags, INTEGER lContext

DECLARE INTEGER FtpPutFile IN wininet.DLL;
   INTEGER hConnect, STRING lpszLocalFile, STRING lpszNewRemoteFile, INTEGER dwFlags, INTEGER dwContext

PUBLIC hOpen, hftpSession, cDestinationFile
STORE .F. to lTransfer_flg 

*--- Initialize Target File name on FTP Server
*--- Target file name = Transmission Place ID + "/" + Aprf

*!*	cDestinationFile = IIF( LEFT(ALLTRIM( cTrans_ID),1) <> '/', '/', '') +;
*!*		cTrans_ID +;
*!*		IIF( RIGHT(ALLTRIM( cTrans_ID),1) <> '/', '/', '') +;
*!*		ALLTRIM( cAprf )

*--- Changed concept of cDestinationFile after discussion with Suzuki San (24-Dec-2003)
cDestinationFile = '%' + ALLTRIM( cTrans_ID ) + '%' +ALLTRIM( cAprf ) + '%b'

IF ConnectFTP(cHost, nPort_No, cUser, cPassword)
	WAIT WINDOW defMsg0379 + cSource NOWAIT
	IF FtpPutFile(hftpSession, cSource, cDestinationFile, nXFerType, 0) = 1

		*...............................................................
		*--- After successful transfer, move file to 'Backup' folder
		mBackup_Path = JUSTPATH( mSource_File ) + "\Backup"
		mBackup_file = mBackUpPath +"\" + ALLTRIM( JUSTFNAME(mSource_File))
		*--- Check if file already exists in Backup folder
		IF FILE( mBackUp_File) 		&& File exists
			ERASE (mBackUp_File)	&& Remove existing file
		ENDIF
		*--- Move new file to the Backup folder
		RENAME (mSource_File) TO (mBackup_file)
		*...............................................................

		STORE .T. to lTransfer_flg
		=MESSAGEBOX(defMsg0382,64,mp_login)
	ELSE
		STORE .F. to lTransfer_flg 
		mP_err[2] = 'E0415'
		DO ErrTrap
	ENDIF
	= InternetCloseHandle (hftpSession)
	= InternetCloseHandle (hOpen)
ENDIF
RELEASE hOpen, hftpSession, cDestinationFile
RETURN lTransfer_flg

*............................................................
FUNCTION  ConnectFTP (cHost, nPort_No, cUser, cPassword)
*............................................................
hOpen = InternetOpen ("vfp", 1, 0, 0, 0)

IF hOpen = 0
	mP_err[2]='E0411'	
	DO Errtrap
    RETURN .F.
ENDIF

hftpSession = InternetConnect (hOpen, cHost, nPort_No, cUser, cPassword, 1, 0, 0) &&... 1 = ftp protocol

IF hftpSession = 0
   * close access to Inet functions and exit
	= InternetCloseHandle (hOpen)
	mP_err[2]='E0412'	
	DO Errtrap
	RETURN .F.
ENDIF
WAIT WINDOW defMsg0380 + cHost NOWAIT 
RETURN
