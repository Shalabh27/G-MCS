  �                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 VERSION =   3.00      
language.h      dataenvironment      dataenvironment      Dataenvironment      Name = "Dataenvironment"
      1      1      form      form      frmBOMVarRep     Height = 335
Width = 767
DoCreate = .T.
ShowTips = .T.
AutoCenter = .T.
BorderStyle = 2
Caption = "[MRPFC140] Material Requirement Planning"
ControlBox = .F.
MaxButton = .F.
MinButton = .F.
Movable = .F.
WindowType = 1
mode = 
Name = "frmBOMVarRep"
     B�PROCEDURE clicked
KEYBOARD CHR(255)
INKEY()

WITH THISFORM

	DO CASE
		CASE .cmdgoperations.VALUE = 1		&&	'Add','Save' & 'Confirm' for Delete
			DO CASE
				CASE EMPTY(.mode)				&&	'Add'
					.mode = 'a'
					.txtdmode.VALUE = defmodeAdd
					.objenable(1)
					.setmode(1)
					.optgroup.opta.SETFOCUS
				CASE .mode = 'a'			&&	'Save'
					rtn_val = .ADD()
					IF rtn_val = 0
						RETU
					ENDIF
					.setmode(0)
				CASE .mode = 'd'			&&	'Confirm' for Delete
					rtn_val = .delcheck()
					IF rtn_val = 0
						RETU
					ENDIF
					.DELETE
					.setmode(0)
			ENDCASE

		CASE .cmdgoperations.VALUE = 2		&&	'Delete' & 'Revert'
			DO CASE
				CASE EMPTY(.mode)				&&	'Delete'
					.mode = 'd'
					.txtdmode.VALUE = defmodeDelete
					.optgroup.opta.ENABLED = .T.
					.optgroup.optb.ENABLED = .T.
					.lstproduct_a.ENABLED = .T.
					.lstproduct_b.ENABLED = .T.
					.setmode(1)
					.optgroup.opta.SETFOCUS
				OTHERWISE
					.setmode(0)				&&	'Revert'
			ENDCASE

		CASE .cmdgoperations.VALUE = 3		&& Print
			err = 'N'
			cntdel = 0
			ctr = 1
			actual_cur = ''

			**	If either of the two group cursors are blank, return...
			DO WHILE ctr <= 2
				err = 'N'
				IF ctr = 1
					actual_cur = 'grpcur_A'
				ELSE
					actual_cur = 'grpcur_B'
				ENDIF

				IF USED('&actual_cur')
					SELE &actual_cur
					COUNT FOR DELETED() = .F. TO cntdel
					IF cntdel = 0
						err= 'Y'
					ENDIF
				ELSE
					err= 'Y'
				ENDIF

				IF err = 'Y'
					mp_err = 'E0206'		&& No Results
					DO errtrap
					RETU
				ENDIF
				ctr = ctr +1
			ENDDO

			mp_ret = sqlexec(mp_handle,"Begin Transaction")		&& Start Transaction
			IF mp_ret <=0
				=AERROR(mp_err)
				DO errtrap
				RETU
			ENDIF

			CNT = 1
			mflag = ''
			actual_cur = ''
			DO WHILE CNT <= 2
				DO CASE
					CASE CNT = 1
						mflag = 'A'
						actual_cur = 'grpcur_A'
						WAIT WIND defmsg0106 NOWAIT
					CASE CNT = 2
						mflag = 'B'
						actual_cur = 'grpcur_B'
						WAIT WIND defmsg0107 NOWAIT
				ENDCASE

				SELE &actual_cur
				GO TOP

				DO WHILE .NOT. EOF()
					mprod = SUBSTR(PRODUCT,1,18)
					mcus1 = SUBSTR(PRODUCT,20,4)
					mcus2 = SUBSTR(PRODUCT,25,2)
					mint  = SUBSTR(PRODUCT,28,2)
					mqty  = qty

					mcom = "EXEC &mP_Ownername bomvar_proc '&mprod','&mcus1','&mcus2','&mint',&mqty,'&mflag'"
					mp_ret = sqlexec(mp_handle,mcom)		&& Execute Backend Proc.

					IF mp_ret <=0
						=AERROR(mp_err)
						DO errtrap
						= sqlexec(mp_handle,"Rollback Transaction")	&& On error - Rollback
						RETU
					ENDIF

					SELE &actual_cur
					SKIP
					IF EOF()
						CNT = CNT + 1
					ENDIF
				ENDDO
			ENDDO

			= sqlexec(mp_handle,"Commit Transaction")		&& Commit & end Transaction
			= SQLCOMMIT(mp_handle)

			**	Make query for the Printer cursor
			WAIT WIND defmsg0108 NOWAIT
			mcom=""
			mcom = "SELECT DISTINCT a.material_no,d.material_nm,a.bom_unit,;
grpQty_A = ISNULL((SELECT CONVERT(NUMERIC(12,3),SUM(b.reqd_qty)) ;
FROM &mP_Ownername.bomvar_temp b ;
WHERE a.material_no = b.material_no and b.flag = 'A' ;
GROUP BY b.material_no),0),"

			mcom = mcom + "grpQty_B = ISNULL((SELECT CONVERT(NUMERIC(12,3),sum(c.reqd_qty)) ;
FROM &mP_Ownername.bomvar_temp c ;
WHERE a.material_no = c.material_no and c.flag = 'B' ;
GROUP BY c.material_no),0),"

			mcom = mcom + "variance = SPACE(11) "
			mcom = mcom + " FROM &mP_Ownername.bomvar_temp a,&mP_Ownername.material_mst d ;
WHERE a.material_no = d.material_no ;
ORDER BY a.material_no"

			SET STAT OFF

			cur='repCur'		&& Unique Name
			SELECT * FROM sectrvew WHERE 1=2 INTO CURSOR &cur		&& Get unique name of prn cursor

			mp_ret=sqlexec(mp_handle,mcom,cur)	&& Execute query

			IF mp_ret<0
				=AERROR(mp_err)
				DO errtrap
				RETU
			ENDIF

			SELE &cur

			**	Calculate the difference of the qty of both group
			WAIT WIND defmsg0109 NOWAIT

			REPLACE ALL variance WITH ;
				IIF((grpqty_b > 0 .AND. grpqty_a = 0),UPPER(defnew),;
				IIF((grpqty_b = 0 .AND. grpqty_a > 0),UPPER(defobsolete),STR((grpqty_b - grpqty_a),11,3)))

			**	Creating variable for products entered in both Groups

			STORE '' TO mp_prodb,mp_proda

			total_prod = 0
			SELE grpcur_a
			GO TOP

			DO WHILE !EOF()
				total_prod = total_prod + 1
				mp_proda =mp_proda + grpcur_a.PRODUCT + grpcur_a.qty
				SKIP
			ENDDO

			SELE grpcur_b
			GO TOP

			DO WHILE !EOF()
				total_prod = total_prod + 1
				mp_prodb =mp_prodb + grpcur_b.PRODUCT + grpcur_b.qty
				SKIP
			ENDDO

			WAIT CLEAR

			*	Calculating total pages for the report
			SELE &cur
			mp_totalpages=1
			_PAGENO=1

			REPO FORM mrpfc140 NOCONSOLE
			mp_totalpages=_PAGENO

			rep="MRPFC140.FRX"
			frm='[MRPFC140]'

			**	Call Print form
			lcmode = .txtdmode.VALUE
			.txtdmode.VALUE=defmodePrint
			DO FORM mrpfc140_print WITH rep,cur,frm		&& Call Print Form
			.txtdmode.VALUE=lcmode
			SET STAT BAR ON
			IF USED('&CUR')
				USE IN &cur
			ENDIF

			**	Delete records from bomvar_temp File
			mcom=	" DELETE FROM "+mp_ownername+"bomvar_temp "
			mp_ret=sqlexec(mp_handle,mcom)

			IF mp_ret <= 0
				=AERROR(mp_err)
				DO errtrap
				RETU
			ENDI
			=SQLCOMMIT(mp_handle)

			**	Close group cursors
			IF USED('grpcur_A')
				USE IN grpcur_a
			ENDIF

			IF USED('grpcur_B')
				USE IN grpcur_b
			ENDIF

			.createvew(1)		&& Make blank Cursors for both Groups

		CASE .cmdgoperations.VALUE = 4
			.RELEASE

	ENDCASE

ENDWITH

ENDPROC
PROCEDURE objenable
LPARAMETER mode

**	Enable / Disable Objects
WITH THISFORM
	DO CASE
		CASE mode = 0
			.optgroup.value = 1
			.optgroup.opta.ENABLED = .F.
			.optgroup.optb.ENABLED = .F.
			.cboproduct_no.ENABLED = .F.
			.txtcusdesch_c1.ENABLED = .F.
			.txtcusdesch_c2.ENABLED = .F.
			.txtintdesch_c.ENABLED = .F.
			.txtqty.ENABLED = .F.
			.lstproduct_a.ENABLED = .F.
			.lstproduct_b.ENABLED = .F.

		CASE mode = 1
			.optgroup.opta.ENABLED = .T.
			.optgroup.optb.ENABLED = .T.
			.cboproduct_no.ENABLED = .T.
			.txtcusdesch_c1.ENABLED = .T.
			.txtcusdesch_c2.ENABLED = .T.
			.txtintdesch_c.ENABLED = .T.
			.txtqty.ENABLED = .T.
			.lstproduct_a.ENABLED = .T.
			.lstproduct_b.ENABLED = .T.
	ENDCASE

ENDWITH

ENDPROC
PROCEDURE createvew
LPARAMETER mode

**Create blank cursor for both groups
WITH THISFORM
	mcom = "select product = product_no+'-'+cusdesch_c1+'-'+cusdesch_c2+'-'+intdesch_c,;
qty = CONVERT(CHAR,bom_qty) ;
from &mP_Ownername.bom_mst where 1 = 2"


	mp_ret = sqlexec(mp_handle,mcom,'grpcur_a')
	mp_ret = sqlexec(mp_handle,mcom,'grpcur_b')

	SELE grpcur_a
	.lstproduct_a.ROWSOURCE = 'grpcur_a'

	SELE grpcur_b
	.lstproduct_b.ROWSOURCE = 'grpcur_b'

	IF mode = 0		&& This proc called in Init of Program
		** Product no. view from Product mst
		mcom = " SELECT product_no,cusdesch_c1,cusdesch_c2,intdesch_c "+;
			"  FROM &mP_Ownername product_mst "+;
			"  where convert(char(10),ed_prd_dt,111) >= CONVERT(CHAR(10),getdate(),111)"+;
			"  ORDER BY product_no,cusdesch_c1,cusdesch_c2,intdesch_c "

			
		mp_ret = sqlexec(mp_handle,mcom,'prodcur')


*PKY 07/18/2002
		IF mP_ret<=0
			=AERROR(mP_err)
			DO ErrTrap
			RETURN 0
		ENDIF
*PKY 07/18/2002

		.cboproduct_no.ROWSOURCE = 'prodcur'
		*.Sele_cur = 'prodcur'

		** Read Terminal Name from Prg_mst
		mcom=" SELECT trml_nm "+;
			" from "+mp_ownername+"Prg_Mst"+;
			" where prg_file='mrpfc140.scx'"

		mp_ret=sqlexec(mp_handle,mcom,'TrmlCur')

		IF mp_ret <= 0
			=AERROR(mp_err)
			DO errtrap
			RETU 0
		ENDI

		**	If anyone else is using this Program -- Error & Return
		SELECT trmlcur
		IF LEN(ALLT(trml_nm))!=0 .AND. ALLT(trml_nm) <> mp_trmlnm
			=MESSAGEBOX(defMsg0026 + " ' " + ALLT(trml_nm) + " '. " + defMsg0027 ,64,mp_login)
			RETU 0
		ELSE
			**	Update terminal name in Prg master with current terminal name
			mcom=	" UPDATE "+mp_ownername+"Prg_Mst "+;
				" set trml_nm='"+mp_trmlnm+"'"+;
				" where prg_file='mrpfc140.scx'"

			mp_ret=sqlexec(mp_handle,mcom)

			IF mp_ret <= 0
				=AERROR(mp_err)
				DO errtrap
				RETU 0
			ENDI
			=SQLCOMMIT(mp_handle)

			**	Update terminal name in cursor
			SELE trmlcur
			REPLACE trml_nm WITH mp_trmlnm
			=TABLEUPDATE(.T.)
		ENDIF

		**	Delete old records from bomvar_temp File
		mcom=" DELETE FROM "+mp_ownername+"bomvar_temp "
		mp_ret=sqlexec(mp_handle,mcom)

		IF mp_ret <= 0
			=AERROR(mp_err)
			DO errtrap
			RETU 0
		ENDI
		=SQLCOMMIT(mp_handle)
	ENDIF

ENDWITH

RETU 1

ENDPROC
PROCEDURE setmode
LPARAMETER mode

**	Set mode for the current operation
WITH THISFORM
	DO CASE
		CASE mode = 0
			ON KEY LABEL esc APPLICATION.ACTIVEFORM.RELEASE
			.cmdgoperations.cmdaddsave.CAPTION = defcmdadd_a
			.cmdgoperations.cmddelcon.CAPTION = defcmddelete_d
			.txtdmode.VALUE = ''
			.cmdgoperations.cmdprint.ENABLED = .T.
			.cmdgoperations.cmdclose.ENABLED = .T.
			.cboproduct_no.DISPLAYVALUE = ''
			.txtcusdesch_c1.VALUE = ''
			.txtcusdesch_c2.VALUE = ''
			.txtintdesch_c.VALUE = ''
			.txtqty.VALUE = 0.000
			.mode = ''
			.objenable(0)

		CASE mode = 1
			ON KEY LABEL esc APPLICATION.ACTIVEFORM.setmode(0)
			DO CASE
				CASE .mode = 'a'
					.cmdgoperations.cmdaddsave.CAPTION = defcmdsave_s
				CASE .mode = 'd'
					.cmdgoperations.cmdaddsave.CAPTION = defcmdconfirm_c
			ENDCASE

			.cmdgoperations.cmddelcon.CAPTION = defcmdrevert_v
			.cmdgoperations.cmdprint.ENABLED = .F.
			.cmdgoperations.cmdclose.ENABLED = .F.
	ENDCASE

ENDWITH

ENDPROC
PROCEDURE add
**	Save pressed
err = 'Y'
actual_cur = ''
WITH THISFORM

	DO CASE
		CASE LEN(ALLTR(.cboproduct_no.DISPLAYVALUE)) = 0
			mp_err = 'E0001'
			obj = 'cboProduct_no.'
		CASE LEN(ALLTR(.txtcusdesch_c1.VALUE)) = 0
			mp_err = 'E0001'
			obj =  'txtcusdesch_c1.'
		CASE LEN(ALLTR(.txtcusdesch_c2.VALUE)) = 0
			mp_err = 'E0001'
			obj = 'txtcusdesch_c2.'
		CASE LEN(ALLTR(.txtintdesch_c.VALUE)) = 0
			mp_err = 'E0001'
			obj = 'txtintdesch_c.'
		CASE .txtqty.VALUE <= 0
			mp_err = 'E0004'
			obj = 'txtqty.'
		OTHERWISE
			err = 'N'
	ENDCASE

	** On error -- Retu
	IF err = 'Y'
		DO errtrap
		.&obj.SETFOCUS
		RETU 0
	ENDIF

	DO CASE
		CASE .optgroup.VALUE = 1
			actual_cur = 'grpcur_a'
		CASE .optgroup.VALUE = 2
			actual_cur = 'grpcur_b'
	ENDCASE

	** Check for duplicate record
	SELE &actual_cur
	GO TOP

	LOCATE FOR PRODUCT = .cboproduct_no.DISPLAYVALUE + '-' +;
		.txtcusdesch_c1.VALUE + '-' +;
		.txtcusdesch_c2.VALUE + '-' +;
		.txtintdesch_c.VALUE

	IF FOUND()
		mp_err = 'E0015'		&& Record already exist
		DO errtrap
		.cboproduct_no.SETFOCUS
		RETU 0
	ENDIF

	** Insert record
	SELE &actual_cur
	APPEN BLANK
	REPLACE PRODUCT WITH .cboproduct_no.DISPLAYVALUE + '-' +;
		.txtcusdesch_c1.VALUE + '-' +;
		.txtcusdesch_c2.VALUE + '-' +;
		.txtintdesch_c.VALUE,;
		qty WITH SPACE(4) + STR(.txtqty.VALUE,5)
	=TABLEUPDATE (.T.)
	SELE grpcur_a
	.lstproduct_a.ROWSOURCE = 'grpcur_a'
	
	SELE grpcur_b
	.lstproduct_b.ROWSOURCE = 'grpcur_b'

ENDWITH
RETU 1

ENDPROC
PROCEDURE delete
**	Confirm for delete pressed

WITH THISFORM
	
	DO CASE
		CASE .optgroup.VALUE = 1
			SELE grpcur_a
		CASE .optgroup.VALUE = 2
			SELE grpcur_b
	ENDCASE
	DELETE
	=TABLEUPDATE (.T.)	
	SELE grpcur_a
	GO TOP
	.lstproduct_a.ROWSOURCE = 'grpcur_a'

	SELE grpcur_b
	GO TOP
	.lstproduct_b.ROWSOURCE = 'grpcur_b'
	.lstproduct_a.REFRESH

ENDWITH
RETU 1

ENDPROC
PROCEDURE delcheck
**	Check before delete

err = 'N'
cntdel = 0
current_rec = 0

WITH THISFORM
	DO CASE
		CASE .optgroup.VALUE = 1
			IF USED('grpcur_A')
				SELE grpcur_a
				*current_rec = RECNO()
				*COUNT FOR DELETED() = .F. TO cntdel			&& if no undeleted records -- Retu
				*IF cntdel = 0
				IF .LSTPRODUCT_A.LISTCOUNT < 1
					RETU 0
				ENDIF
				*GOTO current_rec
			ELSE
				RETU 0
			ENDIF
		CASE .optgroup.VALUE = 2
			IF USED('grpcur_B')
				SELE grpcur_b
				*current_rec = RECNO()
				*COUNT FOR DELETED() = .F. TO cntdel			&& if no undeleted records -- Retu
				*IF cntdel = 0
				IF .LSTPRODUCT_B.LISTCOUNT < 1	
					RETU 0
				ENDIF
				*GOTO current_rec
			ELSE
				RETU 0
			ENDIF
	ENDCASE

ENDWITH
RETU 1

ENDPROC
PROCEDURE lang_change
PARAMETER  mref

WITH THISFORM

	DO CASE
		CASE mref = 0
			.label2.CAPTION = defc140heading
			.label1.CAPTION = defselect_group
			.label6.CAPTION = defproduct_number
			.label3.CAPTION = defgroup_a
			.label4.CAPTION = defgroup_b

			.optgroup.opta.CAPTION = defgroup_a
			.optgroup.optb.CAPTION =	defgroup_b
			.label5.CAPTION = defquantity

			.cmdgoperations.cmdaddsave.CAPTION 		= defcmdadd_a
			.cmdgoperations.cmddelcon.CAPTION 		= defcmddelete_d
			.cmdgoperations.cmdprint.CAPTION 	= defcmdprint_p
			.cmdgoperations.cmdclose.CAPTION = defcmdclose_c
			.command2.TOOLTIPTEXT 	= defhelp
		CASE mref = 1
			PUBLIC mpr_report,mpr_program,mpr_page,mpr_date,mpr_time,mpr_sno,mpr_materialno,mpr_name,mpr_bomunit,mpr_requiredqty,;
				mpr_a_braces,mpr_b_braces,mpr_variance,mpr_groupa,mpr_groupb,mpr_qty,mpr_option,mpr_eor,mpr_b_a_braces

			mpr_report	= defc140heading
			mpr_program	= defprogram
			mpr_page	= defpage
			mpr_date	= defdate
			mpr_time	= deftime
			mpr_sno    = defs_no
			mpr_materialno = defMaterial_Number
			mpr_name  = defname
			mpr_bomunit = defbom_unit
			mpr_requiredqty = defRequired_Quantity
			mpr_a_braces = defa_braces
			mpr_b_braces = defb_braces
			mpr_variance = defvariance
			mpr_groupa = defgroup_a
			mpr_groupb = defgroup_b
			mpr_qty       =defquantity
			mpr_option = defoption
			mpr_eor		= defend_of_report
			mpr_b_a_braces = defb_a_braces
		CASE mref = 2
			RELEASE mpr_report,mpr_program,mpr_page,mpr_date,mpr_time,mpr_sno,mpr_materialno,mpr_name,mpr_bomunit,mpr_requiredqty,;
				mpr_a_braces,mpr_b_braces,mpr_variance,mpr_groupa,mpr_groupb,mpr_qty,mpr_option,mpr_eor,rel_frm,mp_option,;
				mpr_b_a_braces,mp_oldcodecusd1,mp_oldcodecusd2,mp_oldcodeintd,mp_oldcodeprod,mp_qryflg
	ENDC
ENDW

ENDPROC
PROCEDURE GotFocus
**	On error in Form.Init -- Release form
IF rel_frm = 'Y'
	THISFORM.RELEASE
ENDIF


ENDPROC
PROCEDURE Unload
**	Close cursors & release public variables
RELEASE rel_frm,mp_option,mp_oldcodecusd1,mp_oldcodecusd2,mp_oldcodeintd,mp_oldcodeprod,mp_qryflg,;
mp_proda, mp_prodb

THISFORM.lang_change(2)		&& For Releasing Multilanguage Report Public Variables

IF USED('prodcur')
	USE IN prodcur
ENDIF

IF USED('grpcur_A')
	USE IN grpcur_a
ENDIF

IF USED('grpcur_B')
	USE IN grpcur_b
ENDIF

IF USED('ChkCur')
	USE IN chkcur
ENDIF

ON KEY LABEL esc

IF USED('TrmlCur')			&& PKY	07/18/2002
	* Update Terminal Name in Prg_mst
	SELECT trmlcur
	IF ALLT(trml_nm) = mp_trmlnm
		mcom=	" UPDATE "+mp_ownername+"Prg_Mst "+;
			" set trml_nm=''"+;
			" where prg_file='mrpfc140.scx'"

		mp_ret=sqlexec(mp_handle,mcom)

		IF mp_ret <= 0
			=AERROR(mp_err)
			DO errtrap
			RETU
		ENDI
		=SQLCOMMIT(mp_handle)
	ENDI

	USE IN trmlcur
ENDIF

APPLICATION.STATUSBAR=' '

ENDPROC
PROCEDURE Init
***********************************************************************************************
*	MRPFC140				BOM Variance Report
*	AM
*	Date					05/23/1998
*	Modified				01/01/2002		Changes for Multilanguage						MJ
*	Modified On				07/18/2002		DHS-MCS Standardization							PKY
* 	Modification			07/24/2002		Support to multiple date format					PKY
* 	Modification			08/06/2002		Internal Error									PKY
*	Modification			10/10/2002		Added Comas in Product Quantity					SS2
* 	Modification			10/17/2002		Specification No. D2-036						SS2	
*											(Add Search Combo)
*	Modification			04/24/2003		1. Revert Button click selects Grp A			SA
*											2. Report - separate controls for Grp A & B
*											3. Combo Width changed
***********************************************************************************************

PUBLIC rel_frm,mp_option,mp_oldcodecusd1,mp_oldcodecusd2,mp_oldcodeintd,mp_oldcodeprod,mp_qryflg,;
mp_proda, mp_prodb

STORE .F. TO mp_qryflg

STORE SPACE(1) TO mp_oldcodecusd1,mp_oldcodecusd2,mp_oldcodeintd,;
	mp_oldcodeprod

rel_frm = 'N'
STORE '' TO mp_proda, mp_prodb
**	Initialize Print variables
mp_output = 'P'
STORE 1 TO mp_totalpages

WITH THISFORM
	.mode = ''
	.lang_change(0)
	.lang_change(1)
ENDWITH
**	Make cursors & check if this Program is being executed at another Terminal
rtn_val = THISFORM.createvew(0)

WAIT CLEAR
** On error
IF rtn_val = 0
	rel_frm = 'Y'		&&	Release form in Form.Gotfocus
ENDIF

ON KEY LABEL esc APPLICATION.ACTIVEFORM.RELEASE

ENDPROC
PROCEDURE MouseMove
** Storing mouse co-ordinates 
LPARAMETERS nButton, nShift, nXCoord, nYCoord
mP_xcor		=	nXCoord
mP_ycor		=	nYCoord
mP_xcor1	=	nXCoord
mP_ycor1	=	nYCoord

ENDPROC
PROCEDURE Load
WAIT WINDOW defMsg0105 NOWAIT
ENDPROC
     ����    �  �                        c�   %   �      5  /   �          �  U  #  T�  �� � �� T� � ��  �� U  MP_OLDCODEINTD THIS VALUE STATUSBARTEXT� %�C|�� C|�/��  � B� � ���  ����% %�C�� � � �	 � � 	��Z � B� � %�CC� � �>� ��� � J��  �(� � T� ����� E0001�� �	 �	 B�� �� �A %��
 �� � � � �� � 	� � �� � 	� � � � 	��� � B� �� T� ��  SELECT * FROM � � Bom_Mst �  WHERE Product_no='�� � � '�  AND Cusdesch_c1 = '�� � � '�  AND Cusdesch_c2 = '�� � � ' AND Intdesch_c = '� � � '�� T� �C� � � ChkCur�i�� %�� � ���� ��C�� �z�� �	 � B� � F� � %�CN� ��P� J��  �(�
 � � � T� ����� E0106�� �	 �	 B�� �� �> %�CC�� � �>� � CC�� � �>� 	� CC� � �>� 	����> -�� �� � � � �� � 	� � �� � 	� � � � 	�� %�C4���� B� � � �� U  THISFORM MP_XCOR MP_XCOR1 MP_YCOR MP_YCOR1 THIS VALUE MP_OLDCODEINTD MP_ERR ERRTRAP MP_OLDCODEPROD CBOPRODUCT_NO DISPLAYVALUE MP_OLDCODECUSD1 TXTCUSDESCH_C1 MP_OLDCODECUSD2 TXTCUSDESCH_C2 MCOM MP_OWNERNAME MP_RET	 MP_HANDLE CHKCUR
 PRODUCT_NO CUSDESCH_C1 CUSDESCH_C2
 INTDESCH_C�  %�CC�  � � �>� ��" � B�-�� �� T� � ��{ Enter max (2) Character Internal Design Change Code                                                                        � Press <Esc> to Revert�� U  THISFORM TXTCUSDESCH_C2 VALUE THIS STATUSBARTEXT	 LostFocus,     �� Valid     �� When�    ��1 3 �A A � RA A �� �q � A A A ��"� q A A r Q�q � A ��� A A A B 3 �q A B
2                       G         c   �     *     �  C    )   �                       ����    �  �                        �Z   %   �      <  1   �          �  U  #  T�  �� � �� T� � ��  �� U  MP_OLDCODECUSD2 THIS VALUE STATUSBARTEXT� %�C|�� C|�/��  � B� � ���  ����% %�C�� � � �	 � � 	��Z � B� � %�CC� � �>� ��� � T�� � ��  �� J��  �(� �	 �
 � T� ����� E0001�� � �	 B�� �� �A %��
 �� � � � �� � 	� � � � 	� �	 �� � 	��� B� �� T� ��  SELECT * FROM � � Bom_Mst �  WHERE Product_no='�� � � '�  AND Cusdesch_c1 = '�� � � '�  AND Cusdesch_c2 = '� � � '�� T� �C� � � ChkCur�i�� %�� � ���� ��C�� �z�� � � B� � F� � %�CN� ��Z� T�� � ��  �� T� ����� E0106�� J��  �(� �	 �
 � � �	 B�� �� �> %�CC�� � �>� � CC� � �>� 	� CC�� � �>� 	����> -�� �� � � � �� � 	� � � � 	� � �� � 	�� %�C4���� B� � � �� U  THISFORM MP_XCOR MP_XCOR1 MP_YCOR MP_YCOR1 THIS VALUE TXTINTDESCH_C MP_OLDCODECUSD1 MP_OLDCODEINTD MP_OLDCODEPROD MP_ERR ERRTRAP CBOPRODUCT_NO DISPLAYVALUE TXTCUSDESCH_C1 MP_OLDCODECUSD2 MCOM MP_OWNERNAME MP_RET	 MP_HANDLE CHKCUR
 PRODUCT_NO CUSDESCH_C1 CUSDESCH_C2
 INTDESCH_C�  %�CC�  � � �>� ��" � B�-�� �� T� � ��t Enter max (2) character Customer Design Change Code Two                                                             � Press <Esc> to Revert�� U  THISFORM TXTCUSDESCH_C1 VALUE THIS STATUSBARTEXT	 LostFocus,     �� Valid�     �� When�    ��1 3 �A A � RA A �Q�q � A A A �	�"� q A A r �Qq � A ��� A A A B 2 �q A �	2                       H         d   !     ,   <  �  G    )   �                       ����    �  �                        �q   %   �      r  5             �  U  #  T�  � ��  �� T� ��  � �� U  THIS STATUSBARTEXT MP_OLDCODECUSD1 VALUE�  %�CC�  � � �>� ��" � B�-�� �� T� � ��u Enter max (4) character Customer Design Change Code One                                                              � Press <Esc> to Revert�� U  THISFORM CBOPRODUCT_NO DISPLAYVALUE THIS STATUSBARTEXT %�C|�� C|�/��  � B� � ���  �� �% %�C�� � � �	 � � 	��Z � B� � %�CC� � �>� ��� � T�� � ��  �� T�� � ��  �� J��  �(�	 �
 � � T� ����� E0001�� � �	 B�� �� �C %��	 �� � � � � � 	� �
 �  � � 	� � �� � 	�� � B� �t T� ��  SELECT * FROM � � Bom_Mst �  WHERE Product_no='�� � � '�  AND Cusdesch_c1 = '� � � '�� T� �C� � � ChkCur�i�� %�� � ���� ��C�� �z�� � � B� � F� � %�CN� ��Y� T�� � ��  �� T�� � ��  �� J��  �(�	 �
 � � T� ����� E0106�� � �	 B�� �� � F� � #)�@ %�CC� � �>� � CC�  � � �>� 	� CC�� � �>� 	����? -�� �� � � � �� � 	� � �� � 	� � �� � 	�� %�C4���� B� � � �� U  THISFORM MP_XCOR MP_XCOR1 MP_YCOR MP_YCOR1 THIS VALUE TXTCUSDESCH_C2 TXTINTDESCH_C MP_OLDCODEPROD MP_OLDCODECUSD2 MP_OLDCODEINTD MP_ERR ERRTRAP CBOPRODUCT_NO DISPLAYVALUE MP_OLDCODECUSD1 MCOM MP_OWNERNAME MP_RET	 MP_HANDLE CHKCUR PRODCUR
 PRODUCT_NO CUSDESCH_C1 TXTCUSDESCH_C1 CUSDESCH_C2
 INTDESCH_C	 LostFocus,     �� When�     �� Valid�    ��1 3 �q A �	3 �A A � RA A �Q�q � A 5A A D�"� q A A r Q�q � A r Q �� A A A B 1                       H         c   �      	   �   �      )   �                       ����    }  }                        f   %   �	      �
  ^   8
          �  U  _  F�  � ��� ��X � T�� � �� �� T�� � �� �� T�� � ��	 �� T��
 � �� �� �� U  PRODCUR THISFORM CBOPRODUCT_NO DISPLAYVALUE
 PRODUCT_NO TXTCUSDESCH_C1 VALUE CUSDESCH_C1 TXTCUSDESCH_C2 CUSDESCH_C2 TXTINTDESCH_C
 INTDESCH_C#  T�  � ��  �� T� ��  � �� U  THIS STATUSBARTEXT MP_OLDCODEPROD DISPLAYVALUE( ���  ��!�$ %�C|�� C|�� C|���G � \�� {13}��	 B�� �� � %�C|�����f �	 B�� �� � %�C|�� C|�/��� � B� �& %�C�� � � �	 � � 	��� � B� � %�CC� � �>� ��E� T�� � ��  �� T��	 � ��  �� T��
 � ��  �� J�� #�(� � � � T� ����� E0001�� � �	 B�� �� �A %�� � � � � �� � 	� � ��	 � 	� � ��
 � 	���� T� �-�� B� ��� T� �� � �� T� ��� � �� T� ���	 � �� T� ���
 � �� � F� � #)�? %�CC�� � �>� � CC��	 � �>� 	� CC��
 � �>� 	����> -�� � � � � �� � 	� � ��	 � 	� � ��
 � 	�� %�C4���� B� � ��� -�� � � �� � %�C4
��� T�� � ��  �� T��	 � ��  �� T��
 � ��  �� J��  �(� � � � � T� �� E0106�� � �	 B�� �� � �� U  THISFORM MP_XCOR MP_XCOR1 MP_YCOR MP_YCOR1 THIS DISPLAYVALUE TXTCUSDESCH_C1 VALUE TXTCUSDESCH_C2 TXTINTDESCH_C MP_OLDCODECUSD1 MP_OLDCODECUSD2 MP_OLDCODEINTD MP_ERR ERRTRAP MP_OLDCODEPROD	 MP_QRYFLG PRODCUR
 PRODUCT_NO CUSDESCH_C1 CUSDESCH_C2
 INTDESCH_C� � T�  � ��~ Enter / Select max (26) character Product Number                                                                              � Press <Esc> to Revert�� U  THIS STATUSBARTEXTI ��  � � �� � � � � � T� ��  �� T� �� � �	 �� T� ��E�� T� ��	 190,50,30�� T� �a�� %��  �����B�# %�C�
 � �
� C�
 � N� 	��'�/ �
 SEARCH.SCXJ� (� ��
 � � � � �� T� � �a�� ��C� � �� %�CC� �>� ����( T�
 � �CC� �C� ;� ���\���G T� � � �CC� C� ;� ���C� ;� ��C� ;� ���\���G T� � � �CC� C� ;� ���C� ;� ��C� ;� ���\���( T� � � �CC� C� ;� ���\��� ��C�
 � �� � <� �, 12� ESC� _SCREEN.ACTIVEFORM.SetMode(0)� � �� \�C�� �� ��C7�� � U  NKEYCODE NSHIFTALTCTRL LCRETURNVALUE	 LCCAPTION LNLISTBOXWIDTH LCFIRSTCOLUMNWIDTH LLVALUESFROMMULTIPLECOLUMN THISFORM LABEL6 CAPTION THIS	 ROWSOURCE SEARCH SCX OSEARCH
 AUTOCENTER SHOW DISPLAYVALUE TXTCUSDESCH_C1 VALUE TXTCUSDESCH_C2 TXTINTDESCH_C VALID ESC Click,     ��	 LostFocus'    �� Valid�    �� When�    �� KeyPress�    ��1 r � B 4 2 � B� � A #� A �A A cA A �a�q � A � A � A r Q ��� A A � A � �!q � A B 2 q
2 � q� 1� a� !1�� � Q�qq�� A q �A A � � A 2                            	   !  X        t  A     A   \  �  c   C   �  h  f    )   }                       '���                              c[   %   �      u     O          �  U  J  ��  � � � � T� �� �� T� �� �� T� �� �� T� �� �� U  NBUTTON NSHIFT NXCOORD NYCOORD MP_XCOR MP_YCOR MP_XCOR1 MP_YCOR1 
 ��  � � U  THISFORM CLICKED 
 ��  � � U  THISFORM CLICKED0  ��  � � � � T� �� �� T� �� �� U  NBUTTON NSHIFT NXCOORD NYCOORD MP_XCOR MP_YCOR 
 ��  � � U  THISFORM CLICKED 
 ��  � � U  THISFORM CLICKED	 MouseMove,     �� cmdAddSave.Click�     �� cmdDelCon.Click�     �� cmdDelCon.MouseMove    �� cmdPrint.Clickt    �� cmdClose.Click�    ��1 2� � � � 3 � 2 � 2 2� � 3 � 2 � 1                       �         �   �   
   	     !        K  �        �  �          ,      )                          ����    �  �                        �p   %         `     <          �  U  �  %�C|�� C|�/��  � B� �& %�C�� �  � �	 � � 	��N � B� � %�� � � ��� � T� ����� E0004�� � �	 B�� �� � U  MP_XCOR MP_XCOR1 MP_YCOR MP_YCOR1 THIS VALUE MP_ERR ERRTRAP�  %�CC�  � � �>� ��" � B�-�� �� T� � �ً Enter max (5) digit Quantity                                                                                                               � Press <Esc> to Revert�� U  THISFORM TXTINTDESCH_C VALUE THIS STATUSBARTEXT Valid,     �� When	    ��1 �A A cA A B�q � A 3 �q A B2                       +        F  �      )   �                       Y���    @  @                        �{   %   �      �     �          �  U  " T�  �� �� T� �� �� %�C� grpcur_B���W � F� � %�� � ���S � B�-�� � �f � B�-�� �� T� � �ٌ Select Product Number                                                                                                                       � Press <Esc> to Revert�� U  CNTDEL CURRENT_REC GRPCUR_B THIS	 LISTCOUNT STATUSBARTEXT
  F�  � U  GRPCUR_B When,     �� GotFocus�    ��1 � � rq Dq A � q A R2 r 3                       `          �      )   @                       Y���    @  @                        �{   %   �      �     �          �  U  " T�  �� �� T� �� �� %�C� grpcur_A���W � F� � %�� � ���S � B�-�� � �f � B�-�� �� T� � �ٌ Select Product Number                                                                                                                       � Press <Esc> to Revert�� U  CNTDEL CURRENT_REC GRPCUR_A THIS	 LISTCOUNT STATUSBARTEXT
  F�  � U  GRPCUR_A When,     �� GotFocus�    ��1 � � qq Cr A � q A R4 r 3                       ]        |  �      )   @                       ~���    e  e                        �t   %   �                     �  U  � � T�  � �ٓ Select Group                                                                                                                                       � Press <Esc> to Revert�� U  THIS STATUSBARTEXT� � T�  � �ٓ Select Group                                                                                                                                       � Press <Esc> to Revert�� U  THIS STATUSBARTEXT	 OptA.When,     ��	 OptB.When    ��1 �3 �2                       C         c   �       )   e                        ����    �   �                         ^�   %   :       O      I           �  U  
  $�  � U   Click,     ��1 q 1                       "       )   �                          
 language.h^�r��0      �Arial, 0, 9, 5, 15, 12, 32, 3, 0
Arial, 1, 8, 5, 14, 11, 29, 3, 0
Arial, 0, 10, 6, 16, 13, 35, 3, 0
Arial, 0, 8, 5, 14, 11, 30, 3, 0
      �FontSize = 8
Enabled = .F.
Format = "!"
Height = 24
InputMask = "XX"
Left = 553
StatusBarText = ""
TabIndex = 5
Top = 73
Width = 48
ForeColor = 255,255,255
BackColor = 64,128,128
Name = "txtIntdesch_c"
      frmBOMVarRep      txtIntdesch_c      textbox      textbox      �FontSize = 8
Enabled = .F.
Format = "!"
Height = 24
InputMask = "XX"
Left = 481
StatusBarText = ""
TabIndex = 4
Top = 73
Width = 48
ForeColor = 255,255,255
BackColor = 64,128,128
Name = "txtCusdesch_c2"
      frmBOMVarRep      txtCusdesch_c2      textbox      textbox      �FontSize = 8
Enabled = .F.
Format = "!"
Height = 24
InputMask = "XXXX"
Left = 385
MaxLength = 4
StatusBarText = ""
TabIndex = 3
Top = 73
Width = 72
ForeColor = 255,255,255
BackColor = 64,128,128
Name = "txtCusdesch_c1"
      frmBOMVarRep      txtCusdesch_c1      textbox      textbox     ?FontSize = 8
ColumnCount = 4
ColumnWidths = "250,100,58,29"
RowSourceType = 2
Enabled = .F.
Height = 24
ColumnLines = .F.
Left = 145
StatusBarText = ""
TabIndex = 2
Top = 73
Width = 216
ForeColor = 255,255,255
BackColor = 64,128,128
Format = "!"
InputMask = "XXXXXXXXXXXXXXXXXX"
Name = "cboProduct_no"
      frmBOMVarRep      cboProduct_no      combobox      combobox      frmBOMVarRep      Label6      omode
*clicked 
*objenable 
*createvew 
*setmode 
*add 
*delete Delete Event.
*delcheck 
*lang_change 
      line      line      Line1      frmBOMVarRep      SHeight = 0
Left = 25
Top = 24
Width = 719
BorderColor = 0,0,0
Name = "Line1"
      label      label      Label2      frmBOMVarRep      �AutoSize = .F.
FontBold = .T.
FontName = "Times New Roman"
FontSize = 18
BackStyle = 0
Caption = "BOM Comparison Report - Multi"
Height = 29
Left = 24
Top = 2
Width = 634
TabIndex = 11
ForeColor = 0,0,0
Name = "Label2"
      textbox      textbox      txtdDate      frmBOMVarRep     <FontBold = .T.
FontName = "Arial"
FontSize = 8
Alignment = 2
Value = (date())
ControlSource = ""
Enabled = .F.
Height = 24
InputMask = ""
Left = 673
ReadOnly = .T.
SpecialEffect = 0
TabIndex = 12
TabStop = .F.
Top = 0
Width = 72
ForeColor = 255,255,255
DisabledForeColor = 0,0,0
Name = "txtdDate"
      commandbutton      commandbutton      Command2      frmBOMVarRep      �AutoSize = .F.
Top = 1
Left = 747
Height = 24
Width = 20
FontSize = 10
Picture = help.bmp
Caption = ""
TabIndex = 13
TabStop = .F.
ToolTipText = "Help"
Name = "Command2"
      -PROCEDURE Click
HELP	&& Call Help
ENDPROC
      �FontBold = .T.
FontSize = 8
BackStyle = 0
Caption = "Product Number"
Height = 13
Left = 24
Top = 84
Width = 117
TabIndex = 18
Name = "Label6"
      label      label      label      label      Label1      frmBOMVarRep      �AutoSize = .F.
FontBold = .T.
FontSize = 8
BackStyle = 0
Caption = "Select Group"
Height = 15
Left = 24
Top = 60
Width = 117
TabIndex = 14
Name = "Label1"
      optiongroup      optiongroup      Optgroup      frmBOMVarRep     JButtonCount = 2
BackStyle = 0
Value = 1
Height = 25
Left = 145
Top = 48
Width = 216
TabIndex = 1
Name = "Optgroup"
Option1.FontBold = .T.
Option1.FontSize = 8
Option1.BackStyle = 0
Option1.Caption = "Group A"
Option1.Value = 1
Option1.Enabled = .F.
Option1.Height = 17
Option1.Left = 5
Option1.Top = 5
Option1.Width = 102
Option1.Name = "OptA"
Option2.FontBold = .T.
Option2.FontSize = 8
Option2.BackStyle = 0
Option2.Caption = "Group B"
Option2.Enabled = .F.
Option2.Height = 17
Option2.Left = 119
Option2.Top = 4
Option2.Width = 96
Option2.Name = "OptB"
      �PROCEDURE OptA.When
THIS.STATUSBARTEXT = defStb0314 + defStb0027

ENDPROC
PROCEDURE OptB.When
THIS.STATUSBARTEXT = defStb0314 + defStb0027

ENDPROC
      �FontBold = .T.
FontSize = 8
Alignment = 2
Enabled = .F.
Height = 24
Left = 553
TabIndex = 10
Top = 289
Width = 72
DisabledForeColor = 0,0,0
Name = "txtdMode"
      frmBOMVarRep      txtdMode      textbox      textbox      ?Height = 0
Left = 24
Top = 264
Width = 721
Name = "Line2"
      frmBOMVarRep      Line2      line      line      frmBOMVarRep      CmdgOperations      commandgroup      commandgroup      frmBOMVarRep      txtQty      textbox      textbox      frmBOMVarRep      Label5      label      label      frmBOMVarRep      Label4      label      label      listbox      listbox      listbox      lstProduct_A      frmBOMVarRep     FontSize = 8
ColumnCount = 2
ColumnWidths = "195,60"
RowSourceType = 2
Enabled = .F.
Height = 120
ColumnLines = .F.
Left = 145
TabIndex = 7
Top = 121
Width = 264
DisabledBackColor = 192,192,192
DisabledForeColor = 64,0,64
Name = "lstProduct_A"
     �PROCEDURE When
**	If no records in cursor 'grpcur_A' -- return
cntdel = 0
current_rec = 0
IF USED('grpcur_A')
	SELE grpcur_a
	*current_rec = RECNO()
	*COUNT FOR DELETED() = .F. TO cntdel
	IF THIS.LISTCOUNT < 1
	*IF cntdel = 0
		RETU .F.
	ENDIF
	*GOTO current_rec
ELSE
	RETU .F.
ENDIF

THIS.STATUSBARTEXT= defstb0315 + defstb0027


ENDPROC
PROCEDURE GotFocus
**	Make current record selected
SELE grpcur_A
*THIS.SELECTED(RECCOUNT()) = .T.

ENDPROC
     7PROCEDURE MouseMove
** Storing mouse co-ordinates 
LPARAMETERS nButton, nShift, nXCoord, nYCoord
mP_xcor		=	nXCoord
mP_ycor		=	nYCoord
mP_xcor1	=	nXCoord
mP_ycor1	=	nYCoord

ENDPROC
PROCEDURE cmdAddSave.Click
THISFORM.Clicked
ENDPROC
PROCEDURE cmdDelCon.Click
THISFORM.Clicked
ENDPROC
PROCEDURE cmdDelCon.MouseMove
** Storing mouse co-ordinates 
LPARAMETERS nButton, nShift, nXCoord, nYCoord
mP_xcor		=	nXCoord
mP_ycor		=	nYCoord

ENDPROC
PROCEDURE cmdPrint.Click
THISFORM.Clicked
ENDPROC
PROCEDURE cmdClose.Click
THISFORM.Clicked
ENDPROC
      frmBOMVarRep      lstProduct_B      label      label      Label3      frmBOMVarRep      �FontBold = .T.
FontSize = 8
BackStyle = 0
Caption = "Group A"
Height = 16
Left = 24
Top = 132
Width = 104
TabIndex = 15
Name = "Label3"
      listbox     FontSize = 8
ColumnCount = 2
ColumnWidths = "195,40"
RowSourceType = 2
Enabled = .F.
Height = 120
ColumnLines = .F.
Left = 481
TabIndex = 8
Top = 121
Width = 264
DisabledBackColor = 192,192,192
DisabledForeColor = 64,0,64
Name = "lstProduct_B"
     �PROCEDURE When
**	If no records in cursor 'grpcur_B' -- return

cntdel = 0
current_rec = 0

IF USED('grpcur_B')
	SELE grpcur_B
	*current_rec = RECNO()
	*COUNT FOR DELETED() = .F. TO cntdel	
	*IF cntdel = 0
	IF THIS.LISTCOUNT < 1
		Retu .F.	
	ENDIF
	*GOTO current_rec 
ELSE
	Retu .F.	
ENDIF

THIS.STATUSBARTEXT=defStb0315 + defstb0027
ENDPROC
PROCEDURE GotFocus
**	Make current record selected
SELE grpcur_B
*THIS.SELECTED(RECCOUNT()) = .T.

ENDPROC
     �PROCEDURE Valid
*	InCase ESCAPE Pressed  OR ALT + l
IF (LASTKEY()=27) OR (LASTKEY()=47)
	RETU
ENDI

*	Mouse movement(coordinates) Compared
IF (MDOWN() AND ((mp_xcor <> mp_xcor1) OR (mp_ycor <> mp_ycor1)))
	RETU
ENDI

IF THIS.VALUE <= 0
	mp_err[2] = 'E0004'
	DO errtrap
	RETU 0
ENDIF

ENDPROC
PROCEDURE When
IF LEN(ALLTRIM(THISFORM.txtIntdesch_c.VALUE))=0
	RETURN .F.
ENDIF

THIS.STATUSBARTEXT=defStb0316 + defstb0027

ENDPROC
      �FontSize = 8
Alignment = 3
Value = 0
Enabled = .F.
Height = 24
InputMask = "99999"
Left = 697
SelectOnEntry = .T.
TabIndex = 6
Top = 73
Width = 48
DisabledForeColor = 64,0,64
Name = "txtQty"
      �FontBold = .T.
FontSize = 8
BackStyle = 0
Caption = "Group B"
Height = 16
Left = 414
Top = 132
Width = 64
TabIndex = 16
Name = "Label4"
      �FontBold = .T.
FontSize = 8
BackStyle = 0
Caption = "Quantity"
Height = 16
Left = 626
Top = 84
Width = 71
TabIndex = 17
Name = "Label5"
     )ButtonCount = 4
BackStyle = 0
Value = 1
Height = 48
Left = 24
Top = 276
Width = 722
TabIndex = 9
Name = "CmdgOperations"
Command1.Top = 13
Command1.Left = 25
Command1.Height = 24
Command1.Width = 55
Command1.FontSize = 8
Command1.Caption = "\<Add"
Command1.Name = "cmdAddSave"
Command2.Top = 13
Command2.Left = 80
Command2.Height = 24
Command2.Width = 55
Command2.FontSize = 8
Command2.Caption = "\<Delete"
Command2.Name = "cmdDelCon"
Command3.Top = 13
Command3.Left = 135
Command3.Height = 24
Command3.Width = 55
Command3.FontSize = 8
Command3.Caption = "\<Print"
Command3.Name = "cmdPrint"
Command4.Top = 13
Command4.Left = 641
Command4.Height = 24
Command4.Width = 55
Command4.FontSize = 8
Command4.Cancel = .T.
Command4.Caption = "\<Close"
Command4.Name = "cmdClose"
     sPROCEDURE Click
*	If Product_no is selected from combo
SELECT ProdCur
WITH THISFORM

	.cboProduct_no.DISPLAYVALUE	=Product_no
	.txtCusdesch_c1.VALUE		=Cusdesch_c1
	.txtCusdesch_c2.VALUE		=Cusdesch_c2
	.txtIntdesch_c.VALUE		=Intdesch_c

ENDWITH


ENDPROC
PROCEDURE LostFocus
THIS.StatusBarText=''
mP_OldcodeProd=THIS.DisplayValue
ENDPROC
PROCEDURE Valid
*	If Last key pressed was Esc or Alt + V then don't do validations
WITH THISFORM

IF LASTKEY() =5 .OR. LASTKEY() =19 .OR. LASTKEY() = 127
	KEYBOARD '{13}'
	RETU 0
ENDI

*	F5 key press
IF (LASTKEY()=-4) 
	RETU 0
ENDI


*	InCase ESCAPE Pressed  OR ALT + l
IF (LASTKEY()=27) OR (LASTKEY()=47)
	RETU
ENDI

*	Mouse movement(coordinates) Compared
IF (MDOWN() AND ((mP_xcor <> mP_xcor1) OR (mP_ycor <> mP_ycor1)))
	RETU
ENDI

*	If Blank Cannot Leave Field Unless ESC Pressed
IF LEN(ALLT(THIS.DISPLAYVALUE))=0
	.txtCusdesch_c1.VALUE=''
	.txtCusdesch_c2.VALUE=''
	.txtIntdesch_c.VALUE=''
	STORE '#' TO mP_OldcodeCusd1,mP_OldcodeCusd2,mP_OldcodeIntd
	mP_err[2]='E0001'
	DO Errtrap
	RETU 0
ENDI

IF mP_OldcodeProd=THIS.DISPLAYVALUE AND ;
	mP_OldcodeCusd1=.txtCusdesch_c1.VALUE AND ;
	mP_OldcodeCusd2=.txtCusdesch_c2.VALUE AND ;
	mP_OldcodeIntd=.txtIntdesch_c.VALUE
	mp_qryFlg = .F.
	RETU
ELSE
*	.objref(1)
	mP_OldcodeProd=THIS.DISPLAYVALUE 
	mP_OldcodeCusd1=.txtCusdesch_c1.VALUE 
	mP_OldcodeCusd2=.txtCusdesch_c2.VALUE 
	mP_OldcodeIntd=.txtIntdesch_c.VALUE
ENDI

SELECT ProdCur
GO TOP

IF LEN(ALLT(.txtCusdesch_c1.VALUE)) !=0 AND ;
	LEN(ALLT(.txtCusdesch_c2.VALUE)) !=0 AND ;
	LEN(ALLT(.txtIntdesch_c.VALUE)) !=0
	LOCATE FOR	Product_no	=	THIS.DISPLAYVALUE AND ;
	Cusdesch_c1	=	.txtCusdesch_c1.VALUE AND ;
	Cusdesch_c2	=	.txtCusdesch_c2.VALUE AND ;
	Intdesch_c	=	.txtIntdesch_c.VALUE

	IF FOUND()
		RETU
	ENDI
ELSE
	LOCATE FOR 	Product_no	=	THIS.DISPLAYVALUE
ENDI

IF !FOUND()
	.txtCusdesch_c1.VALUE=''
	.txtCusdesch_c2.VALUE=''
	.txtIntdesch_c.VALUE=''
	STORE '' TO mP_OldcodeProd,mP_OldcodeCusd1,mP_OldcodeCusd2,mP_OldcodeIntd
	mP_err='E0106'
	DO Errtrap
	RETU 0
ENDI

ENDW
ENDPROC
PROCEDURE When
THIS.StatusBarText= defStb0002 + defStb0027
ENDPROC
PROCEDURE KeyPress
LPARAMETERS nKeyCode, nShiftAltCtrl
LOCAL lcReturnValue , lcCaption , lnListBoxWidth , lcFirstColumnWidth , llValuesFromMultipleColumn
lcReturnValue	= ""
lcCaption		= THISFORM.Label6.CAPTION
lnListBoxWidth	= 325
lcFirstColumnWidth = '190,50,30'
llValuesFromMultipleColumn = .T.
IF nKeyCode = -4
	IF !EMPTY(THIS.ROWSOURCE) AND RECCOUNT(THIS.ROWSOURCE) > 0
		DO FORM SEARCH.SCX WITH THIS, lcCaption, lnListBoxWidth, lcFirstColumnWidth, llValuesFromMultipleColumn TO lcReturnValue NAME oSearch NOSHOW
		oSearch.AUTOCENTER = .T.
		oSearch.SHOW()
		IF LEN(ALLTRIM(lcReturnValue))!=0
			THIS.DISPLAYVALUE				= ALLT(SUBSTR(lcReturnValue,1,(ATC(';',lcReturnValue,1)-1)))
			THISFORM.txtCusdesch_c1.VALUE= ALLT(SUBSTR(lcReturnValue,(ATC(';',lcReturnValue,1)+ 1),(ATC(';',lcReturnValue,2)- (ATC(';',lcReturnValue,1)+1))))
			THISFORM.txtCusdesch_c2.VALUE= ALLT(SUBSTR(lcReturnValue,(ATC(';',lcReturnValue,2)+ 1),(ATC(';',lcReturnValue,3)-(ATC(';',lcReturnValue,2)+1))))
			THISFORM.txtIntdesch_c.VALUE	= ALLT(SUBSTR(lcReturnValue,(ATC(';',lcReturnValue,3)+ 1)))
			THIS.VALID()
		ENDIF
		RELE oSearch
		ON KEY LABEL ESC _SCREEN.ACTIVEFORM.SetMode(0)
	ENDIF
	NODEFAULT
	KEYBOARD CHR(255)
	INKEY()
ENDIF

ENDPROC
     �PROCEDURE LostFocus
mP_OldcodeIntd=THIS.VALUE
THIS.STATUSBARTEXT=''

ENDPROC
PROCEDURE Valid
*	InCase ESCAPE Pressed  OR ALT + l
IF (LASTKEY()=27) OR (LASTKEY()=47)
	RETU
ENDI

WITH THISFORM
*	Mouse movement(coordinates) Compared
IF MDOWN() AND ((mP_xcor <> mP_xcor1) OR (mP_ycor <> mP_ycor1))
	RETU
ENDI

*	If Blank Cannot Leave Field Unless ESC Pressed
IF LEN(ALLT(THIS.VALUE))=0
	STORE '' TO mP_OldcodeIntd
	mP_err[2]='E0001'
	DO Errtrap
	RETU 0
ENDI

IF mP_OldcodeProd=.cboProduct_no.DISPLAYVALUE AND ;
	mP_OldcodeCusd1=.txtCusdesch_c1.VALUE AND ;
	mP_OldcodeCusd2=.txtCusdesch_c2.VALUE AND mP_OldcodeIntd=THIS.VALUE
	RETU
ENDI

mcom=	" SELECT * FROM "+mP_Ownername+"Bom_Mst "+;
" WHERE Product_no='"+.cboProduct_no.DISPLAYVALUE+"'"+;
" AND Cusdesch_c1 = '"+.txtCusdesch_c1.VALUE+"'"+;
" AND Cusdesch_c2 = '"+.txtCusdesch_c2.VALUE+"' AND Intdesch_c = '"+THIS.VALUE+"'"

mP_ret=SQLEXEC(mP_handle,mcom,"ChkCur")

IF (mP_ret <= 0)
	=AERROR(mP_err)
	DO Errtrap    && Checks Backend Error
	RETU
ENDI

SELECT ChkCur
IF RECCOUNT()=0
	STORE '' TO mP_OldcodeProd,mP_OldcodeCusd2,mP_OldcodeCusd1
	mP_err[2]='E0106'
	DO Errtrap
	RETU 0
ENDI

IF LEN(ALLT(.txtCusdesch_c1.VALUE)) !=0 AND ;
	LEN(ALLT(.txtCusdesch_c2.VALUE)) !=0 AND LEN(ALLT(THIS.VALUE)) !=0
	LOCATE FOR	Product_no	=	.cboProduct_no.DISPLAYVALUE AND ;
	Cusdesch_c1	=	.txtCusdesch_c1.VALUE AND ;
	Cusdesch_c2	=	.txtCusdesch_c2.VALUE AND Intdesch_c	=	THIS.VALUE

	IF FOUND()
		RETU
	ENDI
ENDI

ENDWITH

ENDPROC
PROCEDURE When
IF LEN(ALLTRIM(THISFORM.txtCusdesch_c2.VALUE))=0
	RETURN .F.
ENDIF

THIS.STATUSBARTEXT = defStb0005 + defstb0027

ENDPROC
     �PROCEDURE LostFocus
mP_OldcodeCusd2=THIS.VALUE
THIS.STATUSBARTEXT=''

ENDPROC
PROCEDURE Valid
*	InCase ESCAPE Pressed  OR ALT + l
IF (LASTKEY()=27) OR (LASTKEY()=47)
	RETU
ENDI

WITH THISFORM
*	Mouse movement(coordinates) Compared
IF MDOWN() AND ((mP_xcor <> mP_xcor1) OR (mP_ycor <> mP_ycor1))
	RETU
ENDI

*	If Blank Cannot Leave Field Unless ESC Pressed
IF LEN(ALLT(THIS.VALUE))=0
	.txtIntdesch_c.VALUE=''
	STORE '' TO mP_OldcodeCusd1,mP_OldcodeIntd,mP_OldcodeProd
	mP_err[2]='E0001'
	DO Errtrap
	RETU 0
ENDI

IF mP_OldcodeProd	=	.cboProduct_no.DISPLAYVALUE AND ;
	mP_OldcodeCusd1	=	.txtCusdesch_c1.VALUE AND ;
	mP_OldcodeCusd2	=	THIS.VALUE AND ;
	mP_OldcodeIntd	=	.txtIntdesch_c.VALUE
	RETU
ENDI

mcom=	" SELECT * FROM "+mP_Ownername+"Bom_Mst "+;
" WHERE Product_no='"+.cboProduct_no.DISPLAYVALUE+"'"+;
" AND Cusdesch_c1 = '"+.txtCusdesch_c1.VALUE+"'"+;
" AND Cusdesch_c2 = '"+THIS.VALUE+"'"

mP_ret=SQLEXEC(mP_handle,mcom,"ChkCur")

IF (mP_ret <= 0)
	=AERROR(mP_err)
	DO Errtrap    && Checks Backend Error
	RETU
ENDI

SELECT ChkCur

IF RECCOUNT()=0
	.txtIntdesch_c.VALUE=''
	mP_err[2]='E0106'
	STORE '' TO mP_OldcodeCusd1,mP_OldcodeIntd,mP_OldcodeProd
	DO Errtrap
	RETU 0
ENDI


IF LEN(ALLT(.txtCusdesch_c1.VALUE)) !=0 AND LEN(ALLT(THIS.VALUE)) !=0 AND ;
	LEN(ALLT(.txtIntdesch_c.VALUE)) !=0
	LOCATE FOR	Product_no	=	.cboProduct_no.DISPLAYVALUE AND ;
	Cusdesch_c1	=	.txtCusdesch_c1.VALUE AND ;
	Cusdesch_c2	=	THIS.VALUE AND Intdesch_c	=	.txtIntdesch_c.VALUE

	IF FOUND()
		RETU
	ENDI
ENDI

ENDWITH
ENDPROC
PROCEDURE When
IF LEN(ALLTRIM(THISFORM.txtCusdesch_c1.VALUE))=0
	RETURN .F.
ENDIF

THIS.STATUSBARTEXT = defStb0112 + defstb0027

ENDPROC
     �PROCEDURE LostFocus
THIS.STATUSBARTEXT=''
mP_OldcodeCusd1=THIS.VALUE

ENDPROC
PROCEDURE When
IF LEN(ALLTRIM(THISFORM.cboProduct_no.DISPLAYVALUE))=0
	RETURN .F.
ENDIF

THIS.STATUSBARTEXT = defStb0111 + defstb0027

ENDPROC
PROCEDURE Valid
*	InCase ESCAPE Pressed  OR ALT + l
IF (LASTKEY()=27) OR (LASTKEY()=47)
	RETU
ENDI

WITH THISFORM
*	Mouse movement(coordinates) Compared
IF MDOWN() AND ((mP_xcor <> mP_xcor1) OR (mP_ycor <> mP_ycor1))
	RETU
ENDI

*	If Blank Cannot Leave Field Unless ESC Pressed
IF LEN(ALLT(THIS.VALUE))=0
	.txtCusdesch_c2.VALUE=''
	.txtIntdesch_c.VALUE=''
	STORE '' TO mP_OldCodeProd,mP_OldcodeCusd2,mP_OldcodeIntd
	mP_err[2]='E0001'
	DO Errtrap
	RETU 0
ENDI

IF 	mP_OldCodeProd=.cboProduct_no.DISPLAYVALUE AND ;
	mP_OldcodeCusd1=THIS.VALUE AND ;
	mP_OldcodeCusd2=THISFORM.txtCusdesch_c2.VALUE AND ;
	mP_OldcodeIntd=.txtIntdesch_c.VALUE
	RETU
ENDI

mcom=	" SELECT * FROM "+mP_Ownername+"Bom_Mst "+;
" WHERE Product_no='"+.cboProduct_no.DISPLAYVALUE+"'"+;
" AND Cusdesch_c1 = '"+THIS.VALUE+"'"

mP_ret=SQLEXEC(mP_handle,mcom,"ChkCur")

IF (mP_ret <= 0)
	=AERROR(mP_err)
	DO Errtrap    && Checks Backend Error
	RETU
ENDI

SELECT ChkCur

IF RECCOUNT()=0
	.txtCusdesch_c2.VALUE=''
	.txtIntdesch_c.VALUE=''
	STORE '' TO mP_OldCodeProd,mP_OldcodeCusd2,mP_OldcodeIntd
	mP_err[2]='E0106'
	DO Errtrap
	RETU 0
ENDI

SELECT ProdCur
GO TOP

IF LEN(ALLT(THIS.VALUE)) !=0 AND LEN(ALLT(THISFORM.txtCusdesch_c2.VALUE)) !=0 AND ;
	LEN(ALLT(.txtIntdesch_c.VALUE)) !=0
	LOCATE FOR	Product_no	=	.cboProduct_no.DISPLAYVALUE AND ;
	Cusdesch_c1	=	.txtCusdesch_c1.VALUE AND ;
	Cusdesch_c2	=	.txtCusdesch_c2.VALUE AND ;
	Intdesch_c	=	.txtIntdesch_c.VALUE

	IF FOUND()
		RETU
	ENDI
ENDI

ENDWITH
ENDPROC
     3����    �3  �3                        �   %   �-      �2  �  �.          �  U  ` \�C�� �� ��C7�� ���  ��Y� H�- �U� ��� � ���J� H�O �F� �C�� ���� � T�� �� a�� T�� � �� Add�� ��C��� �� ��C��� �� ��� � �	 � ��� � a��� � T�
 �C�� �� %��
 � ��� � B� � ��C� �� �� ��� � d��F� T�
 �C�� �� %��
 � ��,� B� � ��� � ��C� �� �� � ��� � ���� H�l�� �C�� ��� � T�� �� d�� T�� � �� Delete�� T�� � � �a�� T�� � � �a�� T�� � �a�� T�� � �a�� ��C��� �� ��� � �	 � 2�� ��C� �� �� � ��� � ���8� T� �� N�� T� �� �� T� ���� T� ��  �� +�� ����� T� �� N�� %�� ����� T� �� grpcur_A�� ��� T� �� grpcur_B�� � IF USED('&actual_cur')�/� SELE &actual_cur
 �C'-�(� � %�� � ��+� T� �� Y�� � �E� T� �� Y�� � %�� � Y��x� T� �� E0206�� � � B� � T� �� ��� �$ T� �C� � Begin Transaction�i�� %�� � ���� ��C�� �z�� � � B� � T� ���� T� ��  �� T� ��  �� +�� ����� H�)�� �� ����� T� �� A�� T� �� grpcur_A��> R,:��3 Performing BOM Explosion For Group A... Please wait�� �� ���� T� �� B�� T� �� grpcur_B��> R,:��3 Performing BOM Explosion For Group B... Please wait�� � SELE &actual_cur
 #)� +�C+
���� T� �C� ��\�� T� �C� ��\�� T� �C� ��\�� T�  �C� ��\�� T�! ��" ��] mcom = "EXEC &mP_Ownername bomvar_proc '&mprod','&mcus1','&mcus2','&mint',&mqty,'&mflag'"
 T� �C� �# �i�� %�� � ��V� ��C�� �z�� � �# ��C� � Rollback Transaction�i�� B� � SELE &actual_cur
 H� %�C+���� T� �� ��� � � �! ��C� � Commit Transaction�i�� ��C� �{��3 R,:��( Reading BOM Variance Data... Please wait�� T�# ��  ��� mcom = "SELECT DISTINCT a.material_no,d.material_nm,a.bom_unit,grpQty_A = ISNULL((SELECT CONVERT(NUMERIC(12,3),SUM(b.reqd_qty)) FROM &mP_Ownername.bomvar_temp b WHERE a.material_no = b.material_no and b.flag = 'A' GROUP BY b.material_no),0),"
� mcom = mcom + "grpQty_B = ISNULL((SELECT CONVERT(NUMERIC(12,3),sum(c.reqd_qty)) FROM &mP_Ownername.bomvar_temp c WHERE a.material_no = c.material_no and c.flag = 'B' GROUP BY c.material_no),0),"
& T�# ��# � variance = SPACE(11) ��� mcom = mcom + " FROM &mP_Ownername.bomvar_temp a,&mP_Ownername.material_mst d WHERE a.material_no = d.material_no ORDER BY a.material_no"
 G0� T�$ �� repCur��7 SELECT * FROM sectrvew WHERE 1=2 INTO CURSOR &cur		
 T� �C� �# �$ �i�� %�� � ��
� ��C�� �z�� � � B� � SELE &cur
. R,:��# Calculating Variance... Please wait��f >�& ��C�' � � �( � 	� C� Newf�8 C�' � � �( � 	� C� Obsoletef� C�' �( ��Z66�� J��  �(�) �* � T�+ �� �� F�, � #)� +�C+
��� T�+ ��+ ��� T�* ��* �, � �, �" �� H� � F�- � #)� +�C+
��b� T�+ ��+ ��� T�) ��) �- � �- �" �� H� � R� SELE &cur
 T�. ���� T����� ?� mrpfc1409� T�. ���� T�0 �� MRPFC140.FRX�� T�1 ��
 [MRPFC140]�� T�2 ��� � �� T�� � �� Print��" � mrpfc140_print��0 �$ �1 � T�� � ��2 �� G0 � IF USED('&CUR')�e� USE IN &cur
 �. T�# ��  DELETE FROM �4 � bomvar_temp �� T� �C� �# �i�� %�� � ���� ��C�� �z�� � � B� � ��C� �{�� %�C� grpcur_A���� Q�, � � %�C� grpcur_B���&� Q�- � � ��C���5 �� ��� � ���U� ���6 � � �� U7  THISFORM CMDGOPERATIONS VALUE MODE TXTDMODE	 OBJENABLE SETMODE OPTGROUP OPTA SETFOCUS RTN_VAL ADD DELCHECK DELETE ENABLED OPTB LSTPRODUCT_A LSTPRODUCT_B ERR CNTDEL CTR
 ACTUAL_CUR MP_ERR ERRTRAP MP_RET	 MP_HANDLE CNT MFLAG MPROD PRODUCT MCUS1 MCUS2 MINT MQTY QTY MCOM CUR ALL VARIANCE GRPQTY_B GRPQTY_A MP_PRODB MP_PRODA
 TOTAL_PROD GRPCUR_A GRPCUR_B MP_TOTALPAGES MRPFC140 REP FRM LCMODE MRPFC140_PRINT MP_OWNERNAME	 CREATEVEW RELEASEy ��  � ��� ��r� H�! �n� ��  � ��� � T�� � ���� T�� � � �-�� T�� � � �-�� T�� � �-�� T�� � �-�� T��	 � �-�� T��
 � �-�� T�� � �-�� T�� � �-�� T�� � �-�� ��  ���n� T�� � � �a�� T�� � � �a�� T�� � �a�� T�� � �a�� T��	 � �a�� T��
 � �a�� T�� � �a�� T�� � �a�� T�� � �a�� � �� U  MODE THISFORM OPTGROUP VALUE OPTA ENABLED OPTB CBOPRODUCT_NO TXTCUSDESCH_C1 TXTCUSDESCH_C2 TXTINTDESCH_C TXTQTY LSTPRODUCT_A LSTPRODUCT_BH ��  � ��� ��8�� mcom = "select product = product_no+'-'+cusdesch_c1+'-'+cusdesch_c2+'-'+intdesch_c,qty = CONVERT(CHAR,bom_qty) from &mP_Ownername.bom_mst where 1 = 2"
 T� �C� � � grpcur_a�i�� T� �C� � � grpcur_b�i�� F� � T�� � �� grpcur_a�� F� � T��	 � �� grpcur_b�� %��  � ��4�� mcom = " SELECT product_no,cusdesch_c1,cusdesch_c2,intdesch_c "+ "  FROM &mP_Ownername product_mst "+ "  where convert(char(10),ed_prd_dt,111) >= CONVERT(CHAR(10),getdate(),111)"+ "  ORDER BY product_no,cusdesch_c1,cusdesch_c2,intdesch_c "
 T� �C� � � prodcur�i�� %�� � ��z� ��C��
 �z�� � �	 B�� �� � T�� � �� prodcur��X T� ��  SELECT trml_nm �  from � � Prg_Mst�  where prg_file='mrpfc140.scx'�� T� �C� � � TrmlCur�i�� %�� � ��:� ��C��
 �z�� � �	 B�� �� � F� �" %�CC� �>� �
 C� �� 	����f ��C�! This program is being executed at�  ' C� ��  '. � Execute permission Denied.�@� �x��	 B�� �� ���b T� ��  UPDATE � � Prg_Mst �  set trml_nm='� � '�  where prg_file='mrpfc140.scx'�� T� �C� � �i�� %�� � ���� ��C��
 �z�� � �	 B�� �� � ��C� �{�� F� � >� ��� ��
 ��Ca��� �. T� ��  DELETE FROM � � bomvar_temp �� T� �C� � �i�� %�� � ��$� ��C��
 �z�� � �	 B�� �� � ��C� �{�� � ��	 B���� U  MODE THISFORM MP_RET	 MP_HANDLE MCOM GRPCUR_A LSTPRODUCT_A	 ROWSOURCE GRPCUR_B LSTPRODUCT_B MP_ERR ERRTRAP CBOPRODUCT_NO MP_OWNERNAME TRMLCUR TRML_NM	 MP_TRMLNM MP_LOGIN< ��  � ��� ��5� H�! �1� ��  � ��B�- 12� esc� APPLICATION.ACTIVEFORM.RELEASE� T�� � � �� \<Add�� T�� � � �� \<Delete�� T�� � ��  �� T�� �	 �
 �a�� T�� � �
 �a�� T�� � ��  �� T�� � ��  �� T�� � ��  �� T�� � ��  �� T�� � ��        �� T��  ��  �� ��C� �� �� ��  ���1�0 12� esc�! APPLICATION.ACTIVEFORM.setmode(0)� H����� ���  � a���� T�� � � �� \<Save�� ���  � d���� T�� � � ��	 \<Confirm�� � T�� � � �� Re\<vert�� T�� �	 �
 �-�� T�� � �
 �-�� � �� U  MODE THISFORM ESC CMDGOPERATIONS
 CMDADDSAVE CAPTION	 CMDDELCON TXTDMODE VALUE CMDPRINT ENABLED CMDCLOSE CBOPRODUCT_NO DISPLAYVALUE TXTCUSDESCH_C1 TXTCUSDESCH_C2 TXTINTDESCH_C TXTQTY	 OBJENABLE� T�  �� Y�� T� ��  �� ��� ���� H�5 ��� �CC�� � �>� ��{ � T� �� E0001�� T� �� cboProduct_no.�� �CC�� � �>� ��� � T� �� E0001�� T� �� txtcusdesch_c1.�� �CC��	 � �>� ��	� T� �� E0001�� T� �� txtcusdesch_c2.�� �CC��
 � �>� ��O� T� �� E0001�� T� �� txtintdesch_c.�� ��� � � ���� T� �� E0004�� T� �� txtqty.�� 2��� T�  �� N�� � %��  � Y���� � � .&obj.SETFOCUS
	 B�� �� � H���=� ��� � ���� T� �� grpcur_a�� ��� � ���=� T� �� grpcur_b�� � SELE &actual_cur
 #)�9 -�� �� � � -�� � � -��	 � � -��
 � �� %�C4���� T� �� E0015�� � � ��� � �	 B�� �� � SELE &actual_cur
 �Q >� ���� � � -�� � � -��	 � � -��
 � �� ��C�XC�� � �Z��
 ��Ca��� F� � T�� � �� grpcur_a�� F� � T�� � �� grpcur_b�� ��	 B���� U  ERR
 ACTUAL_CUR THISFORM CBOPRODUCT_NO DISPLAYVALUE MP_ERR OBJ TXTCUSDESCH_C1 VALUE TXTCUSDESCH_C2 TXTINTDESCH_C TXTQTY ERRTRAP OPTGROUP PRODUCT SETFOCUS QTY GRPCUR_A LSTPRODUCT_A	 ROWSOURCE GRPCUR_B LSTPRODUCT_B�  ���  ��� � H� �R � ��� � ���6 � F� � ��� � ���R � F� � � �
 ��Ca��� F� � #)� T�� � �� grpcur_a�� F� � #)� T�� � �� grpcur_b�� ��� � � ��	 B���� U	  THISFORM OPTGROUP VALUE GRPCUR_A GRPCUR_B LSTPRODUCT_A	 ROWSOURCE LSTPRODUCT_B REFRESH* T�  �� N�� T� �� �� T� �� �� ��� ��� H�B �� ��� � ���� � %�C� grpcur_A���� � F� � %��� � ���� �	 B�� �� � �� �	 B�� �� � ��� � ���� %�C� grpcur_B���� F�	 � %���
 � ���� �	 B�� �� � ��	 B�� �� � � ��	 B���� U  ERR CNTDEL CURRENT_REC THISFORM OPTGROUP VALUE GRPCUR_A LSTPRODUCT_A	 LISTCOUNT GRPCUR_B LSTPRODUCT_B 4�  � ��� ��� H�! �� ��  � ����. T�� � �� BOM Comparison Report - Multi�� T�� � �� Select Group�� T�� � �� Product Number�� T�� � �� Group A�� T�� � �� Group B�� T�� �	 � �� Group A�� T�� �
 � �� Group B�� T�� � �� Quantity�� T�� � � �� \<Add�� T�� � � �� \<Delete�� T�� � � �� \<Print�� T�� � � �� \<Close�� T�� � �� Help�� ��  �����O 7� � � � � � � � � � � � � �  �! �" �# �$ �% �* T� �� BOM Comparison Report - Multi�� T� �� Program�� T� �� Page�� T� �� Date�� T� �� Time�� T� �� S.No.�� T� �� Material Number�� T� �� Name�� T� �� BOM Unit�� T� �� Required Quantity�� T� �� (A)�� T� �� (B)�� T� �� Variance�� T�  �� Group A�� T�! �� Group B�� T�" �� Quantity�� T�# �� Option�� T�$ �� END OF REPORT�� T�% �� (B-A)�� ��  ����k <� � � � � � � � � � � � � �  �! �" �# �$ �& �' �% �( �) �* �+ �, � � �� U-  MREF THISFORM LABEL2 CAPTION LABEL1 LABEL6 LABEL3 LABEL4 OPTGROUP OPTA OPTB LABEL5 CMDGOPERATIONS
 CMDADDSAVE	 CMDDELCON CMDPRINT CMDCLOSE COMMAND2 TOOLTIPTEXT
 MPR_REPORT MPR_PROGRAM MPR_PAGE MPR_DATE MPR_TIME MPR_SNO MPR_MATERIALNO MPR_NAME MPR_BOMUNIT MPR_REQUIREDQTY MPR_A_BRACES MPR_B_BRACES MPR_VARIANCE
 MPR_GROUPA
 MPR_GROUPB MPR_QTY
 MPR_OPTION MPR_EOR MPR_B_A_BRACES REL_FRM	 MP_OPTION MP_OLDCODECUSD1 MP_OLDCODECUSD2 MP_OLDCODEINTD MP_OLDCODEPROD	 MP_QRYFLG#  %��  � Y�� �
 �� � � � U  REL_FRM THISFORM RELEASE�' <�  � � � � � � � � � ��C��	 �
 �� %�C� prodcur���U � Q� � � %�C� grpcur_A���x � Q� � � %�C� grpcur_B���� � Q� � � %�C� ChkCur���� � Q� � � {2� esc�  � %�C� TrmlCur����� F� � %�C� �� ����Z T� ��  UPDATE � � Prg_Mst �  set trml_nm=''�  where prg_file='mrpfc140.scx'�� T� �C� � �i�� %�� � ���� ��C�� �z�� � � B� � ��C� �{�� � Q� � � T� � ��  �� U  REL_FRM	 MP_OPTION MP_OLDCODECUSD1 MP_OLDCODECUSD2 MP_OLDCODEINTD MP_OLDCODEPROD	 MP_QRYFLG MP_PRODA MP_PRODB THISFORM LANG_CHANGE PRODCUR GRPCUR_A GRPCUR_B CHKCUR ESC TRMLCUR TRML_NM	 MP_TRMLNM MCOM MP_OWNERNAME MP_RET	 MP_HANDLE MP_ERR ERRTRAP APPLICATION	 STATUSBAR.' 7�  � � � � � � � � � J�-�(� � J�C�X�(� � � � � T�  �� N�� J��  �(� � � T�	 �� P�� J���(�
 � ��� ��� � T�� ��  �� ��C� �� �� ��C��� �� �� T� �C� � � �� R� %�� � ��� � T�  �� Y�� �- 12� esc� APPLICATION.ACTIVEFORM.RELEASE� U  REL_FRM	 MP_OPTION MP_OLDCODECUSD1 MP_OLDCODECUSD2 MP_OLDCODEINTD MP_OLDCODEPROD	 MP_QRYFLG MP_PRODA MP_PRODB	 MP_OUTPUT MP_TOTALPAGES THISFORM MODE LANG_CHANGE RTN_VAL	 CREATEVEW ESCJ  ��  � � � � T� �� �� T� �� �� T� �� �� T� �� �� U  NBUTTON NSHIFT NXCOORD NYCOORD MP_XCOR MP_YCOR MP_XCOR1 MP_YCOR1) & R,:�� Loading Form... Please wait�� U   clicked,     ��	 objenablev    ��	 createvew�    �� setmode�    �� add�    �� delete#    �� delcheckO     �� lang_change�!    �� GotFocus(    �� UnloadD(    �� Init5+    ��	 MouseMove,-    �� Load�-    ��1 � � � � Q� � A� � � 1� A A � 1� A A � � A R� � q!!� � � � � � A R� � � � #� Q� QA �A� � A � � A "!q A A A B� q A A � � � !� � Q�� Q�A BQ � QQQQ� �1� q 1A A BA � A A A � 3� eeb�b 2qb� q A A � �d� q Q � �A A r Q � �A A R � � � � �qa!q q� A �1� q A A � s� A r� A � R� B B 3 q � � !!� � � � � � � !!� � � � � � � A B 3 q � �	��r �r �5�� q � A ���� q � A s !a� � $2� q � A � s � � A �1� q � A � A B � 3 q � � ���!!�� � � 1�1�A �!!A B 3 � � � � �!��!��!��!�Q!A� � A #q !� A � QQQQA CQ �� !q � � A CQ � q �r �B � 3 � � Qq Qq A A � q Q �r Q �� B � 3 � � � � � Qqq T� A � � A Qqq T� A � � A A B � 3 q � � ������������Q��A!�Q�QAAQ1�!�A A 3 "� A 4 sb� A r� A r� A R� A � br 1�2� q A A � A � A 3 q� �� � � � � � � A BR � A �3 2� � � � 3 a1                       �     �   �  �  �   �   �  �!    	  �!  �%  {  (  �%  �+  �  ]  �+  E-  �  p  d-  O0    �  q0  r7  /  �  �7  �7  d  �  8  y;  l  �  �;  �A  �  �  �A  XB  �  �  sB  �B  �   )   �3                  