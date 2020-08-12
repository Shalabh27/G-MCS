SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MRP_Weekly_Tran_36weeks_proc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MRP_Weekly_Tran_36weeks_proc]
GO


/*
	Procedure to generate The MRP Transaction Analysis Report (Weekly)- DHS-MCS Version 4.2.2
	Mar/18/2004		SS2
	Modifications
	03/01/2004		TT 	-	Safety Inv Ratio should be selected from invt_mst,safetyinvt_mst,safe_inv_ratio
	04/04/2004		SA 	-	IWSP consolidation merged with procedure
	04/23/2003		SS	-	Removed Hardcoded collation and fixed bug of safe_inv_ratio 
	05/26/2004		VS	-	Generate the List for 36 weeks and Calling MRP_Weekly_Tran_excel_proc for making 
						36 weeks excel file (Specific to IWSP)
	05/26/2005		ST	-	Bug fix :Consideration of the safety inventory ratio(safetyinvt_mst)  have per supplier improvement is insufficient.
*/
CREATE PROCEDURE [MRP_Weekly_Tran_36weeks_proc]

@seccode char(4),
@fromdate varchar(10),
@material_no varchar(26),
@supplier_c char(4), 
@No_of_column1 int, 
@sql_date_option int,
@multilang_str nvarchar(4000) 
AS

if (@seccode is null)
begin
	print 'Error : Must specify Section Code.'
	return -1
end
if (@fromdate is null)
begin
	print 'Error : Must specify From Date.'
	return -1
end

declare @No_of_column int

set @No_of_column=@No_of_column1+1

declare @StartDatetemp datetime
declare @MStartDatetemp datetime
declare @EndDatetemp datetime
declare @mvar varchar(8000)
declare @mvar0 varchar(8000)
declare @mvar1 varchar(8000)
declare @mvar2 varchar(8000)
declare @mvar3 varchar(8000)
declare @strsql varchar(8000)

declare @po_qt1 varchar(8000)
declare @iss1 varchar(8000)
declare @mrp1 varchar(8000)
declare @sumpo1 varchar(8000)
declare @dqty1 varchar(8000)

declare @matno varchar(26)
declare @Mmatno varchar(26)
declare @i integer
declare @dt varchar(4)
declare @mon varchar(2)
declare @dy varchar(2)
declare @cond int
declare @rowcnt int
declare @Supp_c char(4)
declare @Supp_c2 char(4)
declare @Supp_c3 char(4)
declare @order_ltime int
declare @order_ltime2 int
declare @order_ltime3 int
declare @trans_ltime int
declare @trans_ltime2 int
declare @trans_ltime3 int
declare @Po_minord_qty2 int
declare @Po_minord_qty3 int
declare @lot_size2 int
declare @lot_size3 int
declare @cond1 int
declare @rowcnt1 int
declare @pono varchar(10)
declare @poqty integer
declare @potype varchar(3)
declare @iss int
declare @invt int
declare @diff1 int
declare @diff2 int
declare @mrp int
declare @sumpo int
declare @resqty int
declare @resqty1 int
declare @Mdiff1 varchar(10)
declare @Mdiff2 varchar(10)
declare @Minvt int
declare @Mmrp varchar(10)
declare @Msumpo varchar(10)
declare @Mpoqty varchar(10)
declare @dtdate varchar(5)
declare @fromdateval char(8)
declare @day int
declare @month int
declare @dt1 varchar(4)
-- declare @date1 varchar(4)
-- declare @date2 varchar(4)
-- declare @date3 varchar(4)
-- declare @date4 varchar(4)
-- declare @date5 varchar(4)
-- declare @date6 varchar(4)
-- declare @date7 varchar(4)
-- declare @date8 varchar(4)
-- declare @date9 varchar(4)
-- declare @date10 varchar(4)
-- declare @date11 varchar(4)
-- declare @date12 varchar(4)
-- declare @date13 varchar(4)
-- declare @date14 varchar(4)
-- declare @date15 varchar(4)
-- declare @date16 varchar(4)
declare @date varchar(4)
begin
	set nocount on
	set quoted_identifier off

	set @StartDatetemp=@fromdate
	set @EndDatetemp=dateadd(day,7*@No_of_column,@StartDatetemp)
	if len(datepart(month,@startdatetemp))=1
	begin
		set @mon = '0' + convert(char(1),datepart(month,@startdatetemp))
	end
	else
	begin
		set @mon = convert(char(2),datepart(month,@startdatetemp))
	end
	
	if len(datepart(day,@startdatetemp))=1
	begin
		set @dy = '0' + convert(char(1),datepart(day,@startdatetemp))
	end
	else
	begin
		set @dy = convert(char(2),datepart(day,@startdatetemp))
	end
	set @fromdateval=convert(char(4),datepart(yyyy,@fromdate))+@mon+@dy

	if @material_no=''
	begin
		set @material_no=null
	end
	if @supplier_c=''
	begin
		set @supplier_c=null
	end

	set rowcount 0
	
	SET DATEFIRST 1


--  Query changed for Safety Inventory Ratio (IWSP Consolidation - SA)
-- May.26.2005 SWICS
	SELECT distinct bm.material_no,bm.material_nm,loc_c =ISNULL(bm.location_c,'ZZZZZ'),bm.bom_unit,
	lot_size1=ISNULL(ms.lot_size,ms.lot_size),
	lot_size2=0,lot_size3=0, 
	supp_c=ms.supplier_c, 
	po_ratio=ms.po_ratio,
	po_minord_qty1= ms.po_minord_qty, 0 po_minord_qty2, 0 po_minord_qty3,
	order_ltime=CONVERT(char(3),ms.order_ltime), 
	trans_ltime=CONVERT(char(3),ms.trans_ltime),
	safetyinvt_qty=ISNULL(inv.safetyinvt_qty,0),
--	safetyinvt_ratio= CASE 	WHEN ISNULL(inv.safetyinvt_flg,4)=1 THEN ISNULL(inv.safetyinvt_ratio,0)
--							WHEN ISNULL(inv.safetyinvt_flg,4)=2 THEN ISNULL(sim.safetyinvt_ratio,0)
--							WHEN ISNULL(inv.safetyinvt_flg,4)=3 THEN ISNULL(sir.calc_safetyinvt_ratio,0) 
--							ELSE 0 END,
	safetyinvt_ratio= CASE 	WHEN ISNULL(inv.safetyinvt_flg,4)=1 THEN ISNULL(inv.safetyinvt_ratio,0)
							WHEN ISNULL(inv.safetyinvt_flg,4)=2 THEN ISNULL((SELECT safetyinvt_ratio
								FROM safetyinvt_mst WHERE Section_c = @seccode and supplier_c = (SELECT 
								Supplier_c FROM mrp_matsupp_temp  Where Section_c = @seccode and Material_no =
								bm.material_no and supp_priority = 1)),0)--ISNULL(sim.safetyinvt_ratio,0)
							WHEN ISNULL(inv.safetyinvt_flg,4)=3 THEN ISNULL(sir.calc_safetyinvt_ratio,0) 
							ELSE 0 END,
	bm.unit_confac_po,fixed_ord_qty=ISNULL(inv1.fixed_ord_qty,0)
	,po_method_c=ISNULL(inv1.po_method_c,''),supp_priority=ISNULL(ms.supp_priority,0), 
	Plant_lt= tot_lt,inv1.daily_order,trgfr=isnull(convert(char(10),trgfr,111),''),
	trgto=isnull(convert(char(10),trgto,111),''),'Y' Flag into #tempcur
	FROM  mrp_bom_temp bm 
	LEFT OUTER JOIN invt_mst inv ON bm.section_c=inv.section_c and bm.material_no=inv.material_no 
	LEFT OUTER JOIN  mrp_invt_temp inv1 ON bm.section_c=inv1.section_c and bm.material_no=inv1.material_no 
	LEFT OUTER JOIN  mrp_matsupp_temp ms  ON bm.material_no=ms.material_no and bm.section_c=ms.section_c 
	LEFT OUTER JOIN  target_temp tr ON bm.section_c=tr.section_c and bm.material_no=tr.material_no 
--	LEFT OUTER JOIN safetyinvt_mst sim ON bm.section_c=sim.section_c
	LEFT OUTER JOIN safe_inv_ratio sir ON bm.section_c=sir.section_c and bm.material_no=sir.material_no 
	WHERE bm.section_c=@seccode and bm.material_no = isnull(@material_no, bm.material_no)
	and ms.po_ratio>0 
 
	--Select the materials for the single supplier, If supplied
	select * into #matcur from #tempcur where supp_c = isnull(@supplier_c, supp_c) 
	order by material_no,supp_priority ASC

	--select * from #matcur
	--Backlog po qty only 7 day back open po use and insert in (matCur1)
	SELECT DISTINCT bm.material_no,bck_qty=SUM((isnull(convert(int,(convert(float,po.remain_plan))),0)))
	into #matcur1
	FROM (select distinct section_c,material_no from  mrp_bom_temp where section_c=@seccode) bm 
	LEFT  OUTER JOIN mrp_podata_temp po ON bm.section_c=po.section_c and 
	bm.material_no = po.material_no 
	and  convert(datetime,po.order_date) >= DATEADD(DD,-7,@StartDatetemp) AND 
	convert(datetime,po.order_date) < @fromdate
	WHERE bm.section_c=@seccode and bm.MATERIAL_NO = isnull(@material_no, bm.MATERIAL_NO)
	GROUP BY bm.material_no 
	HAVING SUM((isnull(convert(int,(convert(float,po.remain_plan))),0)))>0
	ORDER BY bm.material_no
	--select * from #matcur1

	set @StartDatetemp=@fromdate
	set @EndDatetemp=dateadd(day,7*@No_of_column,@StartDatetemp)
	set @i = 1
	set @MVAR0='material_no char(26) COLLATE database_default NOT NULL , material_nm char(15) COLLATE database_default NOT NULL, 
		   loc_c char(5) COLLATE database_default NOT NULL, bom_unit char(3) COLLATE database_default NOT NULL ,
	           lot_size1 Numeric(6), lot_size2 Numeric(6), 
                   lot_size3 Numeric(6), supp_c char(12) COLLATE database_default NOT NULL,
	           po_ratio Numeric(3), po_minord_qty1 Numeric,
                   po_minord_qty2 Numeric, po_minord_qty3 Numeric,
                   order_ltime char(9) COLLATE database_default NOT NULL, trans_ltime char(9) COLLATE database_default NOT NULL,
	           safetyinvt_qty Numeric, safetyinvt_ratio Numeric,
                   unit_confac_po Numeric, fixed_ord_qty Numeric,
                   po_method_c char(1) COLLATE database_default NOT NULL, supp_priority Numeric,
                   plant_lt Numeric, daily_order Char(1) COLLATE database_default NOT NULL,
                   bckqty Numeric default 0, reserved_qty Numeric default 0,'
	
        select @mvar='',@mvar1='',@mvar2='',@mvar3=''

	--loop for i=1 to @No_of_column to insert @No_of_column columns
	While datediff(day,@StartDatetemp,@EndDatetemp)>=0
	Begin
                

		set @mon = replicate('0',2-datalength(rtrim(datepart(month,@startdatetemp))))+ ltrim(str(datepart(month,@startdatetemp))) 

		set @dy =  replicate('0',2-datalength(rtrim(datepart(day,@startdatetemp))))+ ltrim(str(datepart(day,@startdatetemp))) 

	

		set @dt= rtrim(@mon) + rtrim(@dy)


		IF @i=@No_of_column 
		begin 	
	  	   set @MVAR='dt' + convert(varchar(2),@i) + ' datetime default null, 
			iss' + @dt + ' Numeric default 0,
			po_no' + @dt + ' char(9) COLLATE database_default default '''',
	      		po_qt' + @dt + ' Numeric  default 0,
			sumpo' + @dt + ' Numeric default 0,
			invt' + @dt + ' Numeric default 0,
			mrp' + @dt + ' Numeric default 0,
			diff2'  + convert(varchar(2),@i) + ' Numeric default 0,
			potp' + @dt + ' char(6) COLLATE database_default default '''',
			diff1' + @dt + ' Numeric  default 0, 
			dqty' + @dt + ' Numeric default 0, 
			trg_fr datetime default null, trg_to datetime default null'
			
		end	
                else	
                begin 
			set @MVAR= 'dt' + convert(varchar(2),@i) + ' datetime default null, 
			iss' + @dt + ' Numeric default 0,
			po_no' + @dt + ' char(9) COLLATE database_default default '''',
			po_qt' + @dt + ' Numeric default 0,
			sumpo' + @dt + ' Numeric default 0,
			invt' + @dt + ' Numeric default 0,
			mrp' + @dt + ' Numeric default 0,
			diff2' + convert(varchar(2),@i) + ' Numeric default 0,
			potp' + @dt + ' char(6) COLLATE database_default default '''',
			diff1' + @dt + ' Numeric default 0,
			dqty' + @dt + ' Numeric default 0,'
                end
                
                IF @i<=16
                BEGIN
                        set @mvar1=@mvar1+@mvar
                END 

                IF @i>16 and @i<=32 
                BEGIN
                        set @mvar2=@mvar2+@mvar
                END 

                IF @i>32 and @i<=48
                BEGIN
                        set @mvar3=@mvar3+@mvar
                END 
                
                IF @i=@No_of_column
                    break
                    

		set @i=@i+1
		Set @StartDatetemp = dateadd(day, 7, @StartDatetemp)		
	End

        set @mvar3 =rtrim(@mvar3) +')'
	set @strsql = 'create table ##finalCur (' + rtrim(@mvar0)
        
 	if exists(select * from tempdb..sysobjects where name='##finalcur' and type='U')
 	drop table ##finalCur
	execute(@strsql+@mvar1+@mvar2+@mvar3)


--	drop table ##finalCur

	insert into ##finalcur (material_no,material_nm,loc_c,bom_unit,lot_size1,lot_size2,lot_size3,supp_c,
        po_ratio,po_minord_qty1,po_minord_qty2,po_minord_qty3,order_ltime,trans_ltime,
        safetyinvt_qty,safetyinvt_ratio,unit_confac_po,fixed_ord_qty,po_method_c,
        supp_priority,plant_lt,daily_order,bckqty,reserved_qty,trg_fr,trg_to)
        select material_no,material_nm,loc_c,bom_unit,
	lot_size1,lot_size2,lot_size3, supp_c, po_ratio, po_minord_qty1,po_minord_qty2,po_minord_qty3,
	order_ltime, trans_ltime, safetyinvt_qty, safetyinvt_ratio,unit_confac_po, fixed_ord_qty,
	po_method_c,supp_priority, Plant_lt,daily_order,0,0,trgfr,trgto 
        from #matcur where po_ratio=100

	delete from #matcur where po_ratio=100

	--Pick only top 3 Suppliers for all those materials whose po_ratio<>100
	select @rowcnt = count(*) from #matcur
	
	Set @cond = 1
	update #matcur set flag='Y'
	while @cond <= @rowcnt
	begin
		set rowcount 1
		--get first material no from #matcur
		select @matno=material_no from #matcur where flag='Y'
		--insert the material no in #finalcur
		insert into ##finalcur (material_no,material_nm,loc_c,bom_unit,lot_size1,lot_size2,lot_size3,supp_c,
                po_ratio,po_minord_qty1,po_minord_qty2,po_minord_qty3,order_ltime,trans_ltime,
                safetyinvt_qty,safetyinvt_ratio,unit_confac_po,fixed_ord_qty,po_method_c,
                supp_priority,plant_lt,daily_order,bckqty,reserved_qty,trg_fr,trg_to)
		select material_no,material_nm,loc_c,bom_unit,
		lot_size1,lot_size2,lot_size3, supp_c, po_ratio, po_minord_qty1,po_minord_qty2,po_minord_qty3,
		order_ltime, trans_ltime, safetyinvt_qty,safetyinvt_ratio, unit_confac_po, fixed_ord_qty,
		po_method_c,supp_priority, Plant_lt,daily_order,0,0,
		trgfr,trgto from #matcur where po_ratio<>100 and material_no=@matno and flag='Y'
		
		--SELECT * FROM #FINALCUR
		set rowcount 0
		--pick top 3 suppliers from #matcur for the material no selected
		--if exists(select * from sysobjects where name='#top3supp' and type='U') drop table #top3supp
		select top 3 *,'Y' FLag1 into #top3supp from #matcur where material_no=@matno
		--SELECT  * FROM #top3supp
		--select @rowcnt1 = count(*) from #top3supp
		--print @rowcnt1
		Set @cond1 = 1
		while @cond1 <= 3
		begin
			set rowcount 1
			if @COND1=1
			begin
				select @supp_c=isnull(supp_c,''),@Order_ltime=isnull(order_ltime,0),
				@Trans_ltime=isnull(trans_ltime,0) from #top3supp where flag1='Y'	
			end
			if @COND1=2
			begin
				select @supp_c2=isnull(supp_c,''),@Order_ltime2=isnull(order_ltime,0),
				@Trans_ltime2=isnull(trans_ltime,0),@Po_minord_qty2=isnull(po_minord_qty1,0), 
				@Lot_size2=isnull(lot_size1,0) from #top3supp where flag1='Y'	
			end
			if @COND1=3
			begin
				select @supp_c3=isnull(supp_c,''),@Order_ltime3=isnull(order_ltime,0),
				@Trans_ltime3=isnull(trans_ltime,0),@Po_minord_qty3=isnull(po_minord_qty1,0), 
				@Lot_size3=isnull(lot_size1,0) from #top3supp where flag1='Y'	
			end
			update #top3supp set flag1='N' where flag1='Y'
			Set @COND1=@COND1+1
		end
		--update order lt, trans lt, po monord qty and lot size for the top 3 suppliers
		update #finalcur set supp_c = isnull(@supp_c,'')+isnull(@Supp_c2,'')+isnull(@Supp_c3,''),
		order_ltime		= isnull(convert(char(3),@order_ltime),'   ') + isnull(convert(char(3),@Order_ltime2),'   ') + isnull(convert(char(3),@Order_ltime3),'   '),
		trans_ltime		= isnull(convert(char(3),@trans_ltime),'   ') + isnull(convert(char(3),@Trans_ltime2),'   ') + isnull(convert(char(3),@Trans_ltime3),'   '),
		po_minord_qty2	= isnull(@Po_minord_qty2,0),
		po_minord_qty3	= isnull(@Po_minord_qty3,0),
		lot_size2		= isnull(@lot_size2,0),
		lot_size3		= isnull(@lot_size3,0) where material_no=@matno
		drop table #top3supp
		set rowcount 0
		update #matcur set flag='N' where material_no=@matno
		Set @COND=@COND+1
	end

	set rowcount 0
        
        select * into #finalCur from ##finalCur 
	drop table ##finalCur

--Create data for Daily Order Qty. SS2
	SELECT date=dateadd(dd,-(datepart(dw,a.date)-1),a.date), a.material_no, sum(po_qty) 
	as po_qty into #dcur from matratio_temp_po a, 
	(select distinct b.material_no from mrp_bom_temp b where b.section_c=@seccode) ib   
	 where a.section_c=@seccode and a.material_no=ib.material_no  
         and a.material_no=isnull(@material_no,a.material_no)
	and date >= @fromdate and  date < DATEADD(week,15,@fromdate) and po_type='d'
	group by dateadd(dd,-(datepart(dw,a.date)-1),a.date), a.material_no

	--create table for @No_of_column week data and insert week dates in #finalcur
	set @StartDatetemp=@fromdate
	set @EndDatetemp=dateadd(day,7*@No_of_column,@StartDatetemp)
	set @i=1

	--drop table #dtcur 
	create table #dtcur (date1 datetime)

	While datediff(day,@StartDatetemp,@EndDatetemp)>0
	Begin
		insert into #dtcur values(@startdatetemp)
		--insert full date in dt column

 		set @strsql = 'update #finalcur set dt' + convert(varchar(2),@i) + '=''' +convert(char(11),@startdatetemp) + ''''
 	 	execute(@strsql)
		set @i=@i+1
		Set @StartDatetemp = dateadd(day, 7, @StartDatetemp)
	End
	--select * from #dtcur
	--Replacing Back order Qty In #finalCur
	
	update #finalcur set bckqty = (select isnull(bck_qty,0) from #matcur1 where 
	#finalcur.material_no=#matcur1.material_no COLLATE database_default)

	--Pick up PO qty week wise from 'po_dtl' for every material_no in 'posource_temp' file

	SELECT distinct bm.material_no,
		date=ISNULL(CASE 
		WHEN dateadd(dd,-(datepart(dw,po.order_date)-1),po.order_date)<@fromdate
		THEN @fromdate 
		ELSE convert(char(10),dateadd(dd,-(datepart(dw,po.order_date)-1),po.order_date),111)
		END, @fromdate) , 
	due_Date1=ISNULL(po.order_date,''),
	po_no=ISNULL(po.po_no,SPACE(10)),po_qty=ISNULL(convert(int,convert(float,remain_plan)),0)
	,'Y' flag into #poqtyCur FROM mrp_bom_temp bm LEFT OUTER JOIN 
	mrp_podata_temp po on po.section_c = bm.section_c AND bm.material_no = po.material_no 
	WHERE bm.section_c=@seccode and bm.MATERIAL_NO = isnull(@material_no, bm.MATERIAL_NO)
	ORDER BY bm.material_no,due_date1,po_no asc
	--select * from #poqtyCur

	SELECT material_no,convert(char(10),dateadd(day,7*(@No_of_column-1),@fromdate),111) AS DATE,' ' AS due_Date1,' ' as po_no,
	SUM(po_qty) AS po_qty INTO #poqty1Cur FROM #poqtycur 
	WHERE DATE >= dateadd(day,7*(@No_of_column-1),@fromdate) GROUP BY material_no
	--select * from #poqty1Cur

	delete from #poqtyCur where DATE >= dateadd(day,7*(@No_of_column-1),@fromdate)
	
	if (select count(*) from #poqty1Cur) <> 0
	begin
		insert into #poqtyCur 
		select material_no,DATE,due_Date1,po_no,po_qty,'Y' from #poqty1Cur
	end
	--select * from #poqtyCur
	--select * from #finalCur
	
	--Put b/o qty,Po_no & sumPo_qty in FinalCur
	--Selecting the first PO for Every Material

	select * into #FirstPOqtyCur from #poqtycur x 
	where material_no+date+po_no = (select top 1 material_no+date+po_no from 
	#poqtycur y where y.material_no = x.material_no and y.date = x.date order by material_no,due_date1,po_no asc) 

	select material_no, date, sum(po_qty) as po_qty into #RestPOqtyCur from #poqtycur x 
	where material_no+date+po_no not in (select material_no+date+po_no from #FirstPOqtyCur)
	group by material_no, date

	set @StartDatetemp=@fromdate
        
        select @day='',@mon='',@date=''
	set @i = 1
	while @i <= @No_of_column
	begin

	        
                set @day = datepart(day, @StartDatetemp)
                set @mon = datepart(month,@StartDatetemp)
                set @date = replicate('0',2-datalength(rtrim(datepart(month,@startdatetemp))))+ ltrim(str(datepart(month,@startdatetemp))) +
                            replicate('0',2-datalength(rtrim(datepart(day,@startdatetemp))))+ ltrim(str(datepart(day,@startdatetemp)))    

-- here
                set @strsql = 'update #finalcur set #finalcur.po_no' + @date + ' = isnull(b.po_no,''''),po_qt' + @date + '= isnull(b.po_qty,0) '
                set @strsql = @strsql + ' from #FirstPOqtyCur b where #finalcur.material_no = b.material_no COLLATE database_default and '
                set @strsql = @strsql + ' datepart(month,b.date) = '+ convert(char(2),@mon) +'and datepart(day,b.date) = '+ convert(char(2),@day) + ''

 	 	execute(@strsql)
	
		set @strsql = 'update #finalcur set #finalcur.sumpo' + @date + ' =  isnull(isnull(#finalcur.sumpo' + @date + ',0)+ b.po_qty,0) '
		set @strsql = @strsql + ' from #restPOqtyCur b where #finalcur.material_no = b.material_no COLLATE database_default and '
		set @strsql = @strsql + ' datepart(month,b.date) = '+ convert(char(2),@mon) +'and datepart(day,b.date) = '+ convert(char(2),@day) + ''
	 	execute(@strsql)
		-- Update Daily qty

		set @strsql = 'update #finalcur set #finalcur.dqty' + @date + ' = b.po_qty '
		set @strsql = @strsql + ' from #dCur b where #finalcur.material_no = b.material_no COLLATE database_default and '
		set @strsql = @strsql + ' datepart(month,b.date) = '+ convert(char(2),@mon) +'and datepart(day,b.date) = '+ convert(char(2),@day) + ''

		execute(@strsql) 
                
                set @StartDatetemp=dateadd(day, 7, @StartDatetemp)
        	set @i = @i+1
	end
        set @StartDatetemp=@fromdate
	set @EndDatetemp=dateadd(day,7*@No_of_column,@StartDatetemp)

	-- PICKED UP ALL PO'S FROM MATRATIO_TEMP_PO
	-- WEEKLY CASE SOME PROB(U,R,F) 
	-- USE Also (U/R,R/F)

	 select date, t.material_no, po_type=(case po_type when 1 then 'U' when 2 then 'R' when 3 then 'D' when 4 then 'F' end), 
		isnull(sum(po_qty_po * unit_confac_po ),0) as po_qty ,'Y' FLag into #mrpcur
		from material_mst a join
		 (  
		 Select date=dateadd(dd,-(datepart(dw,a.date)-1),a.date), a.material_no, a.po_qty_po, 
		 po_type=case po_type when 'u' then 1 when 't' then 2 when 'd' then '3' when 'f' then 4 end 
		 from  matratio_temp_po a 
		 where a.section_c=@seccode and a.MATERIAL_NO = isnull(@material_no, a.MATERIAL_NO)
		 and a.po_type<>'d'
		 UNION ALL
	Select date=dateadd(dd,-(datepart(dw,a.date)-1),a.date), a.material_no, a.po_qty_po,
		po_type = case po_type when 'u' then 1 when 't' then 2 when 'd' then '3' when 'f' then 4 end 
		from matratio_temp_po a
		where a.section_c=@seccode and a.MATERIAL_NO = isnull(@material_no, a.MATERIAL_NO)
	 	and a.po_type='d'
		group by a.date,a.material_no,id,a.po_qty_po,po_type
		 ) t on 
		a.material_no =t.material_no group by date, t.material_no,po_type 
		order by date, t.material_no,case po_type when 1 then '1' when 2 then '2' when 3 then '3' when 4 then '4' end

--	select * from #mrpCur

	--Pick up MRP Qty
	SELECT dateadd(day,7*(@No_of_column-1),@fromdate) AS DATE1, material_no,
	' ' AS po_type, isnull(SUM(po_qty),0) AS po_qty INTO #mrp1Cur FROM #mrpCur
	WHERE DATE >=  dateadd(day,7*(@No_of_column-1),@fromdate) GROUP BY material_no 
	--select * from #mrp1cur

	delete from #mrpCur where [DATE] >= dateadd(day,7*(@No_of_column-1),@fromdate)
	
	if (select count(*) from #mrp1Cur) <> 0
	begin
		insert into #mrpCur
		select DATE1,material_no,po_type,po_qty,'Y' from #mrp1Cur
	end
	-- Select the First PO Type of all the materials
	select * into #FirstMrpCur from #mrpCur x 
	where material_no+convert(char(10),date,111)+po_type = (select top 1 material_no+convert(char(10),date,111)+po_type from 
	#mrpcur y where y.material_no = x.material_no and y.date = x.date ) 

	-- Select the Rest PO Type of all the materials
	select * into #restMrpCur from #mrpCur x 
	where material_no+convert(char(10),date,111)+po_type not in (select material_no+convert(char(10),date,111)+po_type from #FirstMrpCur)

	--Put MRP Qty & PO Type & Mutiple PO Types in finalcur
	set @i = 1
	while @i <= @No_of_column
	begin

                set @day = datepart(day, @StartDatetemp)
                set @mon = datepart(month,@StartDatetemp)
                set @date = replicate('0',2-datalength(rtrim(datepart(month,@startdatetemp))))+ ltrim(str(datepart(month,@startdatetemp))) +
                            replicate('0',2-datalength(rtrim(datepart(day,@startdatetemp))))+ ltrim(str(datepart(day,@startdatetemp)))    

		set @strsql = 'update #finalcur set #finalcur.potp' + @date + ' = isnull(b.po_type,''''), mrp' + @date + '= isnull(b.po_qty,0) '
		set @strsql = @strsql + ' from #FirstMrpCur b where #finalcur.material_no = b.material_no COLLATE database_default and '
		set @strsql = @strsql + ' datepart(month,b.date) = '+ convert(char(2),@mon) +'and datepart(day,b.date) = '+ convert(char(2),@day) + ''

	 	execute(@strsql)
                set @StartDatetemp=dateadd(day, 7, @StartDatetemp)	
	        set @i = @i+1
	end

-- Updating the rest of the PO of the material
------------------------------------------------------
--Put MRP Qty & PO Type & Mutiple PO Types in finalcur

	select @rowcnt = count(*) from #restmrpCur
	Set @cond = 1
	set @poqty= 0
	update #restMrpCur set flag='Y'
	while @cond <= @rowcnt
	begin
		set rowcount 1
		select @potype=po_type,@poqty=po_qty,@startdatetemp=date,
		@matno=material_no from #restmrpCur where flag='Y'
		if len(datepart(month,@startdatetemp))=1
		begin
			set @mon = '0' + convert(char(1),datepart(month,@startdatetemp))
		end
		else
		begin
			set @mon = convert(char(2),datepart(month,@startdatetemp))
		end
		
		if len(datepart(day,@startdatetemp))=1
		begin
			set @dy = '0' + convert(char(1),datepart(day,@startdatetemp))
		end
		else
		begin
			set @dy = convert(char(2),datepart(day,@startdatetemp))
		end
		set @potype = '/'+@potype	
		set @dt= @mon + @dy

 		set @strsql = 'update #finalcur set potp' + @dt + '= rtrim(potp' + @dt + ') + ''' + @potype + ''',mrp' + @dt + '= mrp' + @dt +' + '+ convert(char(10),@poqty) + ''
 		set @strsql = @strsql + ' where material_no = ''' + @matno + ''''
 
 	 	execute(@strsql)
		
		select @strsql = ''
		update #restmrpCur set flag='N' where flag='Y'
		Set @COND=@COND+1
	end
	set rowcount 0

	--requirement consolidated as per week req.
	SELECT date=dateadd(dd,-(datepart(dw,bm.plan_issue_dt)-1),bm.plan_issue_dt),
	bm.material_no,
	invt_qty=ISNULL(inv.invt_qty,0),
	issues=sum(plan_qty), 
	po_type=' ',
	reserved_qty=ISNULL(inv.reserved_qty,0) into #issuecur1
	FROM  mrp_bom_temp bm 
	left outer join  mrp_invt_temp inv on bm.section_c = inv.section_c 
	and bm.material_no=inv.material_no
	WHERE bm.section_c=@seccode and BM.MATERIAL_NO = isnull(@material_no, BM.MATERIAL_NO) 
	and plan_issue_dt < DATEADD(day,7*(@No_of_column-1),@fromdate)
	group by bm.material_no,inv.invt_qty,reserved_qty,
	dateadd(dd,-(datepart(dw,bm.plan_issue_dt)-1),bm.plan_issue_dt)
		UNION ALL 
	--15 weeks & consolidate remaining data for for @No_of_column week
	SELECT date=DATEADD(day,7*(@No_of_column-1),@fromdate),
	bm.material_no,
	invt_qty=ISNULL(inv.invt_qty,0),
	issues=sum(bm.plan_qty),
	po_type='',
	reserved_qty=ISNULL(inv.reserved_qty,0) 
	FROM  mrp_bom_temp bm left outer join mrp_invt_temp inv on 
	bm.section_c = inv.section_c 
	and bm.material_no=inv.material_no 
	WHERE bm.section_c=@seccode AND BM.MATERIAL_NO = isnull(@material_no, BM.MATERIAL_NO)
	and plan_issue_dt >= DATEADD(day,7*(@No_of_column-1),@fromdate)
	GROUP BY bm.material_no,inv.invt_qty,reserved_qty 
--	ORDER BY bm.material_no,date
	
	-- Create @No_of_column Records for Each Material which is used in MRP 
	SELECT DISTINCT iss.material_no,iss.invt_qty,iss.reserved_qty,dt.date1 as date
	INTO #dtCur1 FROM #issueCur1 iss,#dtCur dt 
	
	SELECT dt.material_no ,dt.invt_qty AS invt_qty, ISNULL(iss.issues,0) AS issues, 
	ISNULL(iss.po_type,'') AS po_type ,dt.reserved_qty, dt.date ,'Y' Flag
	INTO #issueCur FROM #dtCur1 dt LEFT OUTER JOIN #issueCur1 iss ON 
	iss.material_no = dt.material_no AND dt.date=iss.date
	ORDER BY dt.material_no,dt.date
	--select * from #issuecur 

	--Put Issues & Invt & diff2 in finalcur
	update #issueCur set flag='Y'
	set @StartDatetemp=@fromdate
	set @EndDatetemp=dateadd(day,7*@No_of_column,@StartDatetemp)
	set @i = 1
	set rowcount 0



	While datediff(day,@StartDatetemp,@EndDatetemp)>0
	begin
		if len(datepart(month,@startdatetemp))=1
		begin
			set @mon = '0' + convert(char(1),datepart(month,@startdatetemp))
		end
		else
		begin
			set @mon = convert(char(2),datepart(month,@startdatetemp))
		end
		
		if len(datepart(day,@startdatetemp))=1
		begin
			set @dy = '0' + convert(char(1),datepart(day,@startdatetemp))
		end
		else
		begin
			set @dy = convert(char(2),datepart(day,@startdatetemp))
		end
	
		set @dt= RTRIM(@mon) + RTRIM(@dy)
		
		--for 1st column
		if @startdatetemp=@fromdate
		begin
			set @strsql = ' update #finalcur set iss' + @dt + '=isnull(b.issues,0),
			invt' + @dt + '=isnull(b.invt_qty,0) ,
			diff1' + @dt + '=(b.invt_qty - b.issues) - b.reserved_qty,
			diff2' + convert(varchar(2),@i) + '=(b.invt_qty - b.issues) - b.reserved_qty,
			reserved_qty=b.reserved_qty from #finalcur, #issuecur b
			where #finalcur.material_no = b.material_no COLLATE database_default and
			#finalcur.dt' + convert(varchar(2),@i) + '=date and 
			#finalcur.dt' + convert(varchar(2),@i) + '=''' + convert(char(11),@startdatetemp) + ''''
		end
		else
		begin
			set @strsql = 'update #finalcur set iss' + @dt + '=b.issues,
			diff2' + convert(char(2),@i) + '=(select a.diff2' + convert(char(2),@i-1) + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no)-b.issues, 
			invt' + @dt + '=isnull((select a.diff1' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) + 
			isnull((select a.mrp' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) + 
			isnull((select a.po_qt' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) + 
			isnull((select a.sumpo' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0),
			diff1' + @dt + '=isnull((select a.diff1' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) + 
			isnull((select a.po_qt' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) + 
			isnull((select a.mrp' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) + 
			isnull((select a.sumpo' + @dt1 + ' from #finalcur a where 
 			a.material_no=#finalcur.material_no),0) - b.issues
			from #finalcur, #issuecur b
			where #finalcur.material_no = b.material_no COLLATE database_default and
			#finalcur.dt' + convert(varchar(2),@i) + '=date and 
			#finalcur.dt' + convert(varchar(2),@i) + '=''' + convert(char(11),@startdatetemp) + ''''
		end
	 	--print @strsql
	 	execute(@strsql)

		set @i=@i+1
		set @dt1=@dt
		Set @StartDatetemp = dateadd(day, 7, @StartDatetemp)
	end

        set @i=1
	set @StartDatetemp=@fromdate
        set @strsql=    'select material_no,material_nm,loc_c,bom_unit,lot_size1,lot_size2,lot_size3,supp_c,
                        po_ratio,po_minord_qty1,po_minord_qty2,po_minord_qty3,order_ltime,trans_ltime,
                        safetyinvt_qty,safetyinvt_ratio,unit_confac_po,fixed_ord_qty,po_method_c,
                        supp_priority,plant_lt,daily_order,bckqty,reserved_qty,trg_fr,trg_to'

        while @i <=@No_of_column
        begin

		set @mon = replicate('0',2-datalength(rtrim(datepart(month,@startdatetemp))))+ ltrim(str(datepart(month,@startdatetemp))) 
		set @dy =  replicate('0',2-datalength(rtrim(datepart(day,@startdatetemp))))+ ltrim(str(datepart(day,@startdatetemp))) 

		set @dt= rtrim(@mon) + rtrim(@dy)

                if  @i<=15
                begin
        	set @strsql= @strsql+ ',dt' + convert(varchar(2),@i) + ', iss' + @dt + ' ,
	        	po_no' + @dt + ',po_qt' + @dt + ' ,sumpo' + @dt + ',
			invt' + @dt + ' ,mrp' + @dt + ' ,diff2' + convert(varchar(2),@i) + ',
			potp' + @dt + ' ,diff1' + @dt + ' ,
			dqty' + @dt + ' '
                end
        
                if @i=16
                   begin     
                            set @dt1= @dt
                            set @strsql= @strsql+ ',dt' + convert(varchar(2),@i) + ', po_no' + @dt + 
                               ',invt' + @dt + ' ,potp' + @dt + ' '

                            set @iss1=',iss' + @dt+'= iss' + @dt+''    
                            set @po_qt1=',po_qt' + @dt+'= iss' + @dt+''    
                            set @sumpo1=',sumpo' + @dt+'= sumpo' + @dt+''    
                            set @dqty1=',dqty' + @dt+'= dqty' + @dt+''    
                            set @mrp1=',mrp' + @dt+'= mrp' + @dt+''
                   end             

                if @i>16
                   begin         
                            set @iss1=@iss1+'+iss' + @dt+''
                            set @po_qt1=@po_qt1+'+po_qt'+ @dt+''
                            set @sumpo1=@sumpo1+'+sumpo' + @dt
                            set @dqty1=@dqty1+'+dqty' + @dt
                            set @mrp1=@mrp1+'+mrp' + @dt

                   end                     

                if @i=@no_of_column 
                begin
                    set @strsql=@strsql+',diff1'+ @dt1+'= diff1'+ @dt+',diff216=diff2'+ convert(varchar(2),@i)+'' 
                end

		set @i=@i+1

		Set @StartDatetemp = dateadd(day, 7, @StartDatetemp)		
                
        end                
        
        select * into ##finalcur from #finalcur 


        execute(@strsql+@iss1+@po_qt1+@sumpo1+@dqty1+@mrp1+' from ##finalcur order by material_no')

        SET DATEFIRST 7	
        
        exec MRP_Weekly_Tran_excel_proc @seccode ,@fromdate ,@multilang_str,@sql_date_option ,@No_of_column 

	
       drop table ##finalCur        
	drop table #finalCur
	drop table #matcur
	drop table #matcur1
	drop table #poqtycur
	drop table #poqty1cur
	drop table #mrpCur
	drop table #mrp1Cur
	drop table #dtcur
	drop table #dtcur1
	drop table #issueCur
	drop table #issueCur1
	drop table #FirstPOqtyCur
	drop table #restPOqtyCur
	drop table #firstmrpCur
	drop table #restmrpCur
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[MRP_Weekly_Tran_36weeks_proc]  TO [mrpgrp]
GO

