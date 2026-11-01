-- ============================================================
-- GENIUS DATABASE: ALL STORED PROCEDURES DEFINITIONS
-- ============================================================

-- PROCEDURE: [dbo].[Proc_InsertCustomerEInvoices]
-- Modified: 2026-10-09 21:46:52.753000
CREATE PROCEDURE [dbo].[Proc_InsertCustomerEInvoices] @FromDate datetime ,@ToDate datetime, @CustIdStr varchar(1000)
AS


INSERT INTO CustomerEInvoices ([CustId], [cust_name_ar], total,InvoiceDate,FromDate,ToDate) 
SELECT [CustId]
      ,[cust_name_ar]
      
      ,Sum([TransValue]) as total
	  ,GetDate()
	  ,@FromDate
	  ,@ToDate
  FROM [dbo].[Customer_EInvoice]
  where (TransDate between @FromDate and @ToDate) --and [CustId] in (@CustIdStr)
   and Charindex(','+cast(CustId as varchar(8000))+',', @CustIdStr) > 0
  group by [CustId]
      ,[cust_name_ar]
GO

----------------------------------------------------------------------

-- PROCEDURE: [dbo].[SaveMonthClose]
-- Modified: 2026-10-09 21:46:52.753000
CREATE procedure [dbo].[SaveMonthClose] as
declare @lastdate as datetime
declare @currdate as datetime
set @currdate=getdate()
set @lastdate =(select isnull(last_month_close_dt,'2013/05/01') From sys_setting)
if month(@lastdate)+year(@lastdate) < month(@currdate)+year(@currdate)
	Begin
		insert into ICS_Month_Close(c_id,itm_id,sto_id,ven_id,itm_qty,pharm_price,sell_price,sell_tax,itm_expiry_date,sec_insert_date,itm_purchase_unit,itm_unit1_unit2,itm_unit1_unit3)
		select c_id,ics.itm_id,sto_id,ven_id
		,itm_qty
		,pharm_price,itm_def_sell_price sell_price,sell_tax,itm_expiry_date,@currdate sec_insert_date
		,itm_purchase_unit,itm_unit1_unit2,itm_unit1_unit3
		From item_class_store ics join item_catalog ic on ic.itm_id=ics.itm_id
		Where sto_id in (Select sto_id from store where activated=1)
		and itm_qty<>0
		update sys_setting set last_month_close_dt=@currdate
	End
GO

----------------------------------------------------------------------

