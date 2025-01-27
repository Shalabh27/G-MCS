****************************************************************************************************
*	INV_POGENTXT	Generate PO Text Data
*	VS
*	Date 		01/17/2004
*				02/09/2004		Do not separate record for local vendor text file		TT
****************************************************************************************************
LPARAMETERS mFile_nm,mSection_c,mValue,mSupply_by		&& Parameters passed

PRIVATE MROW,mPo_hdr,mPo_dtl,mCond			&& Local variables


WAIT WINDOW "Reading "+' ' + IIF(LEN(mValue)=4,mValue+' ','') +"PO Data... Please wait" NOWAIT

** Wrhse Date = Due Date **
mcom="SELECT a.section_c,"
mcom=mcom + "a.po_no,a.supplier_c,a.po_date,due_date=a.wrhse_date,b.material_no,"
mcom=mcom + "b.po_qty,b.po_unit,b.po_price as po_amt,d.material_nm,d.pack_form,c.po_price,"
mcom=mcom + "a.bill_to_sec,b_t_sa1=x.section_add1,b_t_sa2=x.section_add2,b_t_sa3=x.section_add3,"
mcom=mcom + "a.ship_to_sec,s_t_sa1=y.section_add1,s_t_sa2=y.section_add2,s_t_sa3=y.section_add3,"
mcom=mcom + "b_t_snm=x.section_nm,s_t_snm=y.section_nm,c.supp_mat_no,location_c"
mcom=mcom + " FROM &mP_ownername supplier_mst e,&mP_ownername material_mst d,&mP_ownername"
mcom=mcom + " matsupp_mst c,&mP_ownername invPo_dtl_temp b,&mP_ownername invPo_hdr_temp a,&mP_ownername"
mcom=mcom + " section_mst x,&mP_ownername section_mst y,&mP_ownername matloc_mst l"
mcom=mcom + " WHERE a.section_c='&mSection_c' AND b.section_c=a.section_c AND a.supplier_c='&mValue' AND"
mcom=mcom + " b.po_no=a.po_no AND d.material_no=b.material_no AND d.material_no=c.material_no"
mcom=mcom + " AND c.material_no=b.material_no AND c.supplier_c=a.supplier_c AND"
mcom=mcom + " e.supplier_c=a.supplier_c AND a.bill_to_sec=x.section_c AND"
mcom=mcom + " a.ship_to_sec=y.section_c AND b.section_c*=l.section_c AND b.material_no*=l.material_no"
mcom=mcom + " ORDER BY a.po_no,a.supplier_c,b.material_no "

mP_ret=SQLEXEC(mP_handle,mcom,"podataCur")

IF (mP_ret <= 0)
	=AERROR(mP_err)		&& Checks Backend Error
	RETU -1
ENDI

WAIT CLEAR

IF RECCOUNT() = 0
	IF USED('podataCur')
		USE IN podataCur
	ENDIF
	mP_err[2]='E0206'	&&  No Results
	RETU -1
ENDIF

**  Get Data Type  **
mSupplier_c=supplier_c

mcom="SELECT data_no from &mP_ownername datatp_mst"
mcom=mcom + " where section_c='&mSection_c' and data_tp1='PO' and data_tp2='&mSupplier_c'"

mP_ret=SQLEXEC(mP_handle,mcom,"datatpCur")

IF (mP_ret <= 0)
	IF USED('podataCur')
		USE IN podataCur
	ENDIF
	=AERROR(mP_err)		&& Checks Backend Error
	RETU -1
ENDI

IF RECCOUNT() = 0
	IF USED('podataCur')
		USE IN podataCur
	ENDIF
	IF USED('datatpCur')
		USE IN datatpCur
	ENDIF
	mP_err[2]='E0150'	&&  Record Not found in Data Type Master
	RETU -1
ENDIF

mDataTp=datatpCur->data_no
***************

IF FILE(mFile_nm) AND MESSAGEBOX("File Already Exists. Overwrite It ?",4+32,mP_login)=7
	IF USED('podataCur')
		USE IN podataCur
	ENDIF
	IF USED('datatpCur')
		USE IN datatpCur
	ENDIF
	mP_err[2]='E0218'	&&  Operation is cancelled by the user
	RETU -1
ENDIF

mFileHandle = FCREATE(mFile_nm)

IF mFileHandle < 0
	IF USED('podataCur')
		USE IN podataCur
	ENDIF
	IF USED('datatpCur')
		USE IN datatpCur
	ENDIF
	mP_err[2]='E0217'	&& PO Text File creation Error
	RETU -1
ENDIF

mBuffsize=0

SELECT podataCur

WAIT WINDOW "Creating" +' '+ IIF(LEN(mValue)=4,mValue+' ','') + "PO Text Data... Please wait" NOWAIT


IF mSupply_by = '1'			&&
	mRow1 = '$$$$100' + SPACE(42) + mDataTp + '000160' + PADL(ALLT(STR(RECCOUNT()*2+1)),5,'0') + ;
		SUBSTR(TTOC(DATETIME()),9,2) + SUBSTR(TTOC(DATETIME()),1,2) + SUBSTR(TTOC(DATETIME()),4,2) + ;
		SUBSTR(TTOC(DATETIME()),12,2) + SUBSTR(TTOC(DATETIME()),15,2) + SPACE(3) + '$$$$'
	FPUTS(mFileHandle,mRow1)
ENDIF

DO WHILE !EOF()
	mPo_date=DTOC(po_date)
	mDue_date=DTOC(due_date)

	mRow1 = 'R'+ section_c + '1' + po_no + supplier_c + section_c + ;
		SUBSTR(mPo_date,1,2) + SUBSTR(mPo_date,4,2) + SUBSTR(mPo_date,7,4) + ;
		SUBSTR(mDue_date,1,2) + SUBSTR(mDue_date,4,2) + SUBSTR(mDue_date,7,4) + ;
		material_no + SUBSTR(material_nm,1,14)

	mPo_price= STR(po_price,12,4)
	mPo_amt  = STR(po_amt,15,4)
	mRow2 = SUBSTR(material_nm,15,1) + IIF (supplier_c='0001','0','1') + ;
		PADL(ALLT(STR(po_qty)),9,'0') + po_unit + '0' + NVL(pack_form,' ') + ;
		PADL(ALLT(SUBSTR(mPo_price,1,7)),7,'0') + SUBSTR(mPo_price,9,4) + ;
		PADL(ALLT(SUBSTR(mPo_amt,1,10)),10,'0') + SUBSTR(mPo_amt,12,2) + ;
		'0' + LEFT(NVL(location_c,'ZZZZZ'),4) + SPACE(36)  

	IF mSupply_by='1'
		mOutBytes = FPUTS(mFileHandle,mRow1) + FPUTS(mFileHandle,mRow2)
	ELSE
		mOutBytes = FPUTS(mFileHandle,mRow1+mRow2)
	ENDIF
	mBuffsize=mBuffsize + mOutBytes


* Flush Buffer
	IF 	mBuffsize > 10000
		FFLUSH(mFileHandle)
		mBuffsize=0
	ENDIF
	SKIP
ENDDO

* Close the File
FCLOSE(mFileHandle)

IF USED('podataCur')
	USE IN podataCur
ENDIF

IF USED('datatpCur')
	USE IN datatpCur
ENDIF

WAIT CLEAR
RETU 0
