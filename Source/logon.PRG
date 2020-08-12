*!*	****************************************************************************************************
*	MRP  	      	LOGON.PRG	(DHS-Asia)
*	
*	11/01/2000	
*	11/10/2001		Option for Safety Inventory Calculation in Ratio/Days  	HPS
*	Modified		Multilanguage Changes									NP
*	Modified		03/23/2002				Multilocks set to 'ON'			VB
*	Modified		06/19/2002				Multiple Dateformat Support		RSN
*	Modified		10/26/2002				Add variable for BOM IBIS/EDI 	VS
*	Modified		05/24/2003				Add variable for CTS DB 		NP
*	Modified		06/18/2003				Remove variable for BOM Upload 			NP
*	Modified		08/22/2003				Pick Database Name from connection 		NP
****************************************************************************************************
PARAMETERS mDsn,mSite

*!*	* 
*mDsn='SS2'
*!*	mDsn='mrpdb_iwsp_toyota'
*!*	*!*	mDsn='mrpdb_iwsp_plnt2'

*mDsn='ANYNAME'

mSite='IWSP Test'


#INCLUDE 'language.h'

*	Change VFP Screen Title
APPLICATION.CAPTION= mSite+ defApplication_Caption
* 	Settings
CLEAR
CLOSE ALL
SET SAFETY OFF
SET TALK OFF
SET SYSM OFF
SET MULTILOCKS ON
SET DELETED on
RESTORE MACROS FROM DEFAULT.FKY					&& Macro settings

*	Declaration of public Variables
PUBLIC 	mP_login,mP_handle,mP_trmlnm,mP_xcor,mP_ycor,mP_xcor1,mP_ycor1,mP_axsstp,;
		mP_ownername

PUBLIC mP_totalpages,mP_output,mP_rangehigh,mP_rangelow	

*Safety Inventory Calculation in Ratio/Days HPS 11/10/2001
PUBLIC mP_safetyinvt_base
mP_safetyinvt_base='R'			&& 'R'=Ratio,'D'='Days'


********************************************************
** Specify Common Transfer System Database Name
PUBLIC mp_CTSDB
mp_CTSDB = 'Common_Transfer'
********************************************************
*	Public Variable for Date Format Setting
* 	Value can be- 	'YMD','AMERICAN',British
* 	YMD 		- 	yyyy/mm/dd
* 	AMERICAN 	- 	mm/dd/yyyy
*	British		-	dd/mm/yyyy
*	Uncomment the setting which you want for the system
********************************************************
PUBLIC mP_CurrDateFormat
*mP_CurrDateFormat = 'YMD'
*mP_CurrDateFormat = 'AMERICAN'
mP_CurrDateFormat = 'BRITISH'

	   		
*	Public Variables
PUBLIC ARRAY mP_err[7]	&& Array for trapping ODBC errors

*	Initialization of Public Variables
STORE SPACE(1) TO mP_axsstp						&& User access type
STORE SPACE(30) TO mP_trmlnm					&& Terminal Name
STORE SPACE(8) TO mP_login,m_pwd				&& User Login,passwd
STORE 0 TO mP_xcor,mP_ycor,mP_xcor1,mP_ycor1	&& For trapping mouse move
mP_totalpages=1
mP_safetyinvt_base=UPPER(mP_safetyinvt_base)

mcom=' '

DO FORM LOG							&& Background Form 

*!*	mP_handle=SQLCONNECT(mDsn,'mrpsa','admin')			&&	Displays SQL Server Login Prompt
mP_handle=SQLCONNECT(mDsn)			&&	Displays SQL Server Login Prompt

IF mP_handle <= 0					&&	Error if connection Invalid
	=MESSAGEBOX(defMsg0248,16,defLogon_Message)
	RELEASE mP_handle
	RELE WINDOW
	CLOSE ALL
	RETU
ELSE	
	
	** Obtain database name from connection 
	mcom="SELECT DBNAME=convert(varchar,db_name())"
	mP_ret=SQLEXEC(mP_handle,mcom,"cur")
	
	IF mP_ret <=0		&& Error
		=MESSAGEBOX(defMsg0249,48,mP_login)
		=SQLDISCONN(mP_handle)
	   	RELE WINDOW
		RETURN
	ELSE
		mDatabase = ALLTRIM(cur->dbname)
		mP_ownername = mDatabase + '..'
		USE IN cur
	ENDI

	mSource=SQLGETPROP(mP_Handle, "ConnectString")
	slen=len(msource)		&& length of msource
	**	Find 'UID' in variable mSource
	ctr= AT('UID',mSource)
	IF ctr <> 0
		ctr = ctr + 4		&& Move the pointer in front of =
		DO WHILE SUBSTR(mSource,ctr,1) != ";" 
			mP_login = mP_login + SUBSTR(mSource,ctr,1)
			ctr = ctr +1
		ENDDO
		mP_login=ALLT(mP_login)
	ENDIF
	
	**	Find 'PWD' in variable mSource
	ctr= AT('PWD',mSource)
	IF ctr <> 0
		ctr = ctr + 4		&& Move the pointer in front of =
		DO WHILE SUBSTR(mSource,ctr,1) != ";" 
			m_pwd = m_pwd + SUBSTR(mSource,ctr,1)
			ctr = ctr +1
		ENDDO
		m_pwd=ALLT(m_pwd)
	ENDIF

	**	Find 'WSID' in variable mSource
	ctr= AT('WSID',mSource)
	IF ctr <> 0
		ctr = ctr + 5		&& Move the pointer in front of =
		DO WHILE SUBSTR(mSource,ctr,1) != ";" AND ctr <= slen
			mP_trmlnm = mP_trmlnm + SUBSTR(mSource,ctr,1)
			ctr = ctr +1
		ENDDO
		mP_trmlnm=ALLT(mP_trmlnm)	&& Terminal Name
	ENDIF
	
	mcom="Select * From &mP_ownername.auth_ctrl where userid= '" +mP_login+ "'"
	mP_ret=SQLEXEC(mP_handle,mcom,"cur")
	
	IF mP_ret <=0		&& Error
		=MESSAGEBOX(defMsg0249,48,mP_login)
		=SQLDISCONN(mP_handle)
	   	RELE WINDOW
		RETU
	ENDI
	
	SELECT Cur
	IF RECCOUNT() > 0
		mP_axsstp=access_tp		&& User access type
		CLOSE DATABASES
		OPEN DATABASE MRP SHARE

		* Record Fetch Settings for Views
		=CURSORSETPROP("MaxRecords", -1, 0)
		=CURSORSETPROP("FetchSize", -1, 0)

		*	Creating Connections
		CREATE CONNECTION Connect1 ;
		DATASOURCE &mDsn ;
		USERID &mP_login ;
		PASSWORD (m_pwd) ;
		DATABASE &mDatabase
		
		CREATE SQL VIEW messRvew ;
		CONNECTION Connect1 SHARED ;
		AS ;
		SELECT * FROM Message_Mst 
			
		CREATE SQL VIEW sectRvew ;
		CONNECTION Connect1 SHARED  ;
		AS ;
		SELECT Section_c,Section_nm FROM Section_Mst 
		
		SELECT I
		USE sectRvew		&& from MRP Database 
		
		SELECT J
		USE messRvew		&& from MRP Database 
		WAIT CLEA
		RELE WINDOW
		
		*	Setting the date format for SQL Server and VFP according to current format
		SET CENTU ON
		PUBLIC mP_sqlDateOption
		SET PROCEDURE TO procfile.prg ADDITIVE
		IF (mP_CurrDateFormat<>'AMERICAN') AND (mP_CurrDateFormat<>'YMD') AND (mP_CurrDateFormat<>'BRITISH')
			mP_CurrDateFormat = 'YMD'			
		ENDIF
		ResetDateFormat(mP_CurrDateFormat)
		=SQLEXEC(mp_handle,'select * from prg_mst','menucur')
		SET EXACT on
*		DO FORM MRPFS010	&& Call MRP Menu		
	ELSE
		=MESSAGEBOX(defMsg0250,16,mP_login)
		=SQLDISCONN(mP_handle)
		RELE WINDOW
	ENDI
ENDI
