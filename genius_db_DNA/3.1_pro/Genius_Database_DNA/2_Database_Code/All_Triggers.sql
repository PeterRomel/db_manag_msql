-- ==========================================
-- TRIGGER: [dbo].[UPDATEDATE] ON TABLE [Customer]
-- ==========================================
CREATE TRIGGER [UPDATEDATE] ON [dbo].[Customer] 
FOR UPDATE
AS

IF  UPDATE(cust_current_credit) 
BEGIN
	DECLARE @id decimal(18,0)
	select @id=(SELECT cust_id FROM inserted)
	update customer set sec_update_date=getdate() where cust_id=@id
END 





GO

-- ==========================================
-- TRIGGER: [dbo].[VUPDATEDATE] ON TABLE [Vendor]
-- ==========================================
CREATE TRIGGER [VUPDATEDATE] ON [dbo].[vendor] 
FOR UPDATE
AS
IF  UPDATE(ven_current_credit) 
BEGIN
	DECLARE @id decimal(18,0)
	select @id=(SELECT ven_id FROM inserted)
	update vendor set sec_update_date=getdate() where ven_id=@id
END 







GO

-- ==========================================
-- TRIGGER: [dbo].[delete_duplicate_records] ON TABLE [F_Transaction_Header]
-- ==========================================
create TRIGGER [dbo].[delete_duplicate_records] ON [dbo].[f_transaction_header] 
FOR INSERT, UPDATE 
AS 
DECLARE @fh_notes VARCHAR(200) 
DECLARE @fh_trans_type decimal(18,0) 
DECLARE @fh_trans_type2 decimal(18,0)
DECLARE @fh_form_type decimal(18,0)
DECLARE @fh_id decimal(18,0)

select @fh_notes = (SELECT fh_notes FROM inserted)
select @fh_trans_type = (SELECT fh_trans_type FROM inserted)
select @fh_trans_type2 = (SELECT fh_trans_type2 FROM inserted)
select @fh_form_type = (SELECT fh_form_type FROM inserted)
select @fh_id = (SELECT fh_id FROM inserted)

if( @fh_trans_type=1 and @fh_trans_type2=1 and @fh_form_type=1 and @fh_notes like '%فاتورة بيع  رقم%')
begin
if not exists (select * from sysobjects where name='F_Transaction_Header_SaveDeleteRecords' and xtype='U')
CREATE TABLE [dbo].[F_Transaction_Header_SaveDeleteRecords] (
	[fh_id] [decimal](18, 0)  NULL ,
	[fh_trans_type] [varchar] (2) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_trans_type2] [varchar] (2) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_code] [varchar] (10) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_value] [money] NULL ,
	[fh_from_type] [varchar] (2) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_from_id] [decimal](18, 0) NULL ,
	[fh_to_type] [varchar] (2) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_to_id] [decimal](18, 0) NULL ,
	[fh_notes] [varchar] (250) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_actual_date] [datetime] NULL ,
	[fh_computer] [varchar] (100) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[fh_actual_cashier_id] [decimal](18, 0) NULL ,
	[fh_form_type] [int] NULL ,
	[sec_insert_uid] [varchar] (50) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[sec_insert_date] [datetime] NULL ,
	[sec_update_uid] [varchar] (50) COLLATE SQL_Latin1_General_CP1256_CI_AS NULL ,
	[sec_update_date] [datetime] NULL ,
	[fh_sto_id] [decimal](18, 0) NULL ,
	[fh_cost_sto_id] [int] NULL 
)
Insert into F_Transaction_Header_SaveDeleteRecords select * from f_transaction_header where fh_trans_type=1 and fh_trans_type2=1 and fh_form_type=1 and fh_notes=@fh_notes and fh_id<>@fh_id
   delete from f_transaction_header where fh_trans_type=1 and fh_trans_type2=1 and fh_form_type=1 and fh_notes=@fh_notes and fh_id<>@fh_id
end

GO

-- ==========================================
-- TRIGGER: [dbo].[watch_stock] ON TABLE [Item_Class_Store]
-- ==========================================
---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//---//
Create TRIGGER watch_stock ON Item_Class_Store 
FOR INSERT, UPDATE 
AS 
DECLARE @AppName VARCHAR(200) 
DECLARE @PCName VARCHAR(200) 
DECLARE @StoreId decimal(18,0) 
DECLARE @ItemId decimal(18,0) 
DECLARE @ClassId decimal(18,0) 
DECLARE @OldQty decimal(18,0) 
DECLARE @NewQty decimal(18,0) 
DECLARE @itm_exp datetime 
select @AppName = (SELECT top 1 program_name from master.dbo.sysprocesses (nolock) where spid=@@spid) 
select @PCName = (SELECT top 1 hostname from master.dbo.sysprocesses (nolock) where spid=@@spid) 
IF EXISTS (select * from deleted) 
BEGIN 
	IF UPDATE(itm_qty) or UPDATE(itm_expiry_date)
	BEGIN 
		select @StoreId = (SELECT sto_id FROM inserted) 		
		if EXISTS (select sto_id from store where sto_id=@StoreId and activated<>2) 
		Begin
			select @ItemId = (SELECT itm_id FROM inserted) 
			select @ClassId = (SELECT c_id FROM inserted) 
			select @OldQty = (SELECT itm_qty FROM deleted) 
			select @NewQty = (SELECT itm_qty FROM inserted) 
			select @itm_exp = (SELECT itm_expiry_date FROM deleted)
			insert into watch_qty_chng (c_id,itm_id,sto_id,old_qty,new_qty,prog_nm,pc_nm,oper,itm_expiry_date) 
				values (@ClassId,@ItemId ,@StoreId,@OldQty,@NewQty,@AppName,@PCName,'UPDATE',@itm_exp) 
		END
	END 
END 
ELSE 
BEGIN 
	IF EXISTS (select * from inserted) 
	BEGIN 
		select @StoreId = (SELECT sto_id FROM inserted) 
		if EXISTS (select sto_id from store where sto_id=@StoreId and activated<>2)
		Begin
			select @ItemId = (SELECT itm_id FROM inserted) 
			select @ClassId = (SELECT c_id FROM inserted) 
			select @OldQty = (SELECT isnull(itm_qty,0) FROM deleted) 
			select @NewQty = (SELECT itm_qty FROM inserted) 
			select @itm_exp = (SELECT itm_expiry_date FROM deleted)
			insert into watch_qty_chng (c_id,itm_id,sto_id,old_qty,new_qty,prog_nm,pc_nm,oper,itm_expiry_date) 
				values (@ClassId,@ItemId ,@StoreId,@OldQty,@NewQty,@AppName,@PCName,'INSERT',@itm_exp) 
		End
	END 
END


GO

-- ==========================================
-- TRIGGER: [dbo].[close_stock] ON TABLE [Item_Class_Store]
-- ==========================================
Create TRIGGER [close_stock] ON [Item_Class_Store] After UPDATE, Insert 
AS 
DECLARE @StoreId decimal(18,0) 
select @StoreId = (SELECT top 1 sto_id FROM inserted) 

if EXISTS (select sto_id from store where sto_id=@StoreId and activated<>2)
Begin
	------ ---------- --------------- SaveMonthClose
	declare @lastdate as datetime
	declare @currdate as datetime
	set @currdate= getdate()
	set @lastdate =(select isnull(last_month_close_dt,'2013/05/01') From sys_setting)
	if DATEDIFF ( month , @lastdate , @currdate ) > 0
	Begin
		insert into ICS_Month_Close(c_id,itm_id,sto_id,ven_id,itm_qty,pharm_price,sell_price,sell_tax,itm_expiry_date,sec_insert_date,itm_purchase_unit,itm_unit1_unit2,itm_unit1_unit3)
		select c_id,ics.itm_id,sto_id,ven_id,itm_qty,pharm_price,itm_def_sell_price sell_price,sell_tax,itm_expiry_date,@currdate sec_insert_date,itm_purchase_unit,itm_unit1_unit2,itm_unit1_unit3
		From item_class_store ics join item_catalog ic on ic.itm_id=ics.itm_id
		Where sto_id in (Select sto_id from store where activated=1)
		and itm_qty<>0
	
		update sys_setting set last_month_close_dt=@currdate
	
	------ ---------- --------------- Undo item change 
	DECLARE @ItemId decimal(18,0) 
	DECLARE @ClassId decimal(18,0) 
	DECLARE @OldQty decimal(18,0) 
	DECLARE @OldPharmPrice money
	DECLARE @OldTaxPrice money

	IF EXISTS (select * from deleted) -- case update
	Begin
		select @ItemId = (SELECT itm_id FROM inserted) 
		select @ClassId = (SELECT c_id FROM inserted) 
		select @OldQty = (SELECT isnull(itm_qty,0) FROM deleted)
		select @OldPharmPrice = (SELECT isnull(pharm_price,0) FROM deleted)
		select @OldTaxPrice = (SELECT isnull(sell_tax,0) FROM deleted)

		update ICS_Month_Close set itm_qty=@OldQty,pharm_price=@OldPharmPrice,sell_tax=@OldTaxPrice
			where itm_id=@ItemId and c_id=@ClassId and sto_id=@StoreId and sec_insert_date=(select last_month_close_dt from sys_setting)
	End
	else--Insert
	Begin
		IF EXISTS (select * from inserted) 
		BEGIN
			select @ItemId = (SELECT itm_id FROM inserted) 
			select @ClassId = (SELECT c_id FROM inserted) 
			delete from ICS_Month_Close where itm_id=@ItemId and c_id=@ClassId and sto_id=@StoreId
		End
	End
	End
End


GO

-- ==========================================
-- TRIGGER: [dbo].[watch_cust_credit_chnge] ON TABLE [Customer]
-- ==========================================
CREATE TRIGGER watch_cust_credit_chnge ON customer FOR INSERT, UPDATE AS DECLARE @custId decimal(18,0) DECLARE @old_credit money DECLARE @new_credit money DECLARE @program_name VARCHAR(200) DECLARE @pc_name VARCHAR(200) IF EXISTS (select * from deleted) BEGIN IF UPDATE(cust_current_credit) BEGIN  select @custId = (SELECT cust_id FROM inserted)  select @old_credit= (SELECT cust_current_credit FROM deleted)  select @new_credit= (SELECT cust_current_credit FROM inserted)  select @program_name= (SELECT program_name from master.dbo.sysprocesses (nolock) where spid=@@spid)  select @pc_name= (SELECT hostname from master.dbo.sysprocesses (nolock) where spid=@@spid)  INSERT INTO customer_credit_chng(cc_cust_id, cc_old_credit, cc_new_credit, cc_prog_name, cc_pc_name, cc_operation) values (@custId, @old_credit, @new_credit, @program_name, @pc_name, 'Update') END END ELSE BEGIN IF EXISTS (select * from inserted) BEGIN  select @custId = (SELECT cust_id FROM inserted)  select @old_credit= (SELECT cust_current_credit FROM deleted)  select @new_credit= (SELECT cust_current_credit FROM inserted)  select @program_name= (SELECT program_name from master.dbo.sysprocesses (nolock) where spid=@@spid)  select @pc_name= (SELECT hostname from master.dbo.sysprocesses (nolock) where spid=@@spid)  INSERT INTO customer_credit_chng(cc_cust_id, cc_old_credit, cc_new_credit, cc_prog_name, cc_pc_name, cc_operation) values (@custId, @old_credit, @new_credit, @program_name, @pc_name, 'Insert') END END

GO

-- ==========================================
-- TRIGGER: [dbo].[watch_vendor_credit_chnge] ON TABLE [Vendor]
-- ==========================================
CREATE TRIGGER watch_vendor_credit_chnge ON vendor FOR INSERT, UPDATE AS DECLARE @vendId decimal(18,0) DECLARE @old_credit money DECLARE @new_credit money DECLARE @program_name VARCHAR(200) DECLARE @pc_name VARCHAR(200) IF EXISTS (select * from deleted) BEGIN IF UPDATE(ven_current_credit) BEGIN  select @vendId = (SELECT ven_id FROM inserted)  select @old_credit= (SELECT ven_current_credit FROM deleted)  select @new_credit= (SELECT ven_current_credit FROM inserted)  select @program_name= (SELECT program_name from master.dbo.sysprocesses (nolock) where spid=@@spid)  select @pc_name= (SELECT hostname from master.dbo.sysprocesses (nolock) where spid=@@spid)  INSERT INTO vendor_credit_chng(vc_cust_id, vc_old_credit, vc_new_credit, vc_prog_name, vc_pc_name, vc_operation) values (@vendId, @old_credit, @new_credit, @program_name, @pc_name, 'Update') END END ELSE BEGIN IF EXISTS (select * from inserted) BEGIN  select @vendId = (SELECT ven_id FROM inserted)  select @old_credit= (SELECT ven_current_credit FROM deleted)  select @new_credit= (SELECT ven_current_credit FROM inserted)  select @program_name= (SELECT program_name from master.dbo.sysprocesses (nolock) where spid=@@spid)  select @pc_name= (SELECT hostname from master.dbo.sysprocesses (nolock) where spid=@@spid)  INSERT INTO vendor_credit_chng(vc_cust_id, vc_old_credit, vc_new_credit, vc_prog_name, vc_pc_name, vc_operation) values (@vendId, @old_credit, @new_credit, @program_name, @pc_name, 'Insert') END END

GO

