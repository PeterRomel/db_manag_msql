-- =============================================================================
-- GENIUS (E-PLUS) ERP: CATALOG DEDUPLICATION MIGRATION SCRIPT
-- Target Database: Genius (Microsoft SQL Server)
-- Safety Directives:
--   1. ZERO HARD DELETES: Preserves referential integrity in sales_trans_d & pur_trans_d.
--   2. Secondary barcodes repointed to Master SKU in Item_Objects.
--   3. Active branch inventory balances consolidated in Item_Class_Store.
--   4. Duplicate items soft-deactivated in Item_Catalog (itm_active = '0', itm_stop_sell = 1).
-- =============================================================================

USE [Genius];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;
BEGIN TRY
    -- 1. Create Staging Table for Verified Deduplication Pairs
    IF OBJECT_ID('tempdb..#VerifiedDuplicates') IS NOT NULL
        DROP TABLE #VerifiedDuplicates;

    CREATE TABLE #VerifiedDuplicates (
        Master_Code VARCHAR(50) NOT NULL,
        Duplicate_Code VARCHAR(50) NOT NULL,
        PRIMARY KEY CLUSTERED (Master_Code, Duplicate_Code)
    );

    PRINT 'Loading verified duplicate clusters into staging table...';
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84662', '85169');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81715', '84603');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('65630', '74138');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('26307', '65275');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('79898', '85370');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('52989', '65299');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('67744', '62230289');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('18448', '50813');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('26725', '82565');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('13655', '64124');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('36106', '84887');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('52778', '64369');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84605', '85067');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('80222', '82045');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('77027', '82191');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('55372', '79167');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('52775', '66615');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('29679', '85183');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('53863', '64921');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('63644', '72852');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('52928', '84275');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('73328', '84901');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('61691', '85520');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84053', '85519');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('25146', '25416');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('30193', '66115');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('70170', '81899');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('34307', '65592');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('13259', '74313');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('13993', '65867');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('27919', '84919');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('30531', '63276');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81812', '85565');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('8092', '75340');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('64056', '83872');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('56644', '66848');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('56300', '65460');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('30472', '64867');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84265', '84870');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('55955', '85614');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('21037', '63896');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('33428', '63845');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('67833', '75958');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('27780', '85373');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('57919', '83688');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('49638', '62230156');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('59071', '66076');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('4705', '64533');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('49355', '64072');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('49603', '85299');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('63865', '84935');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('76868', '85188');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('27204', '66556');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('26301', '83721');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('61267', '79588');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('61268', '80267');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83892', '85044');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('79902', '81006');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('64225', '77459');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('63483', '73644');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('73645', '85555');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('37967', '63855');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('57128', '59927');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('34876', '66620');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('61040', '64537');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('59273', '66080');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('10844', '85330');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('72292', '80679');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('59585', '81773');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('76395', '85324');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('43591', '65411');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81331', '85274');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('80903', '85488');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('64855', '83684');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('77529', '80908');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('77529', '85334');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('54242', '80385');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('27474', '64854');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83953', '85552');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('38191', '83950');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83795', '85620');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('82042', '85309');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('61742', '82965');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('40696', '85225');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81475', '85510');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81537', '83293');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84981', '85601');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('74141', '79965');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('80723', '82296');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83876', '85095');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('80997', '85652');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('77299', '85564');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('82614', '85468');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('53599', '84992');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('33255', '66333');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('9431', '83033');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('43741', '60325');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('58690', '65072');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('43554', '66162');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('43554', '84574');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('53408', '56867');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('75211', '84373');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('79537', '80767');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('31836', '80337');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('78314', '84912');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('61367', '85168');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('33183', '65732');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('30277', '82219');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84577', '84985');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('47598', '72161');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('62612', '64815');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('1894', '66709');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('16288', '77630');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('30165', '64204');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('31967', '64998');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('31967', '84681');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('34001', '81528');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('48315', '85056');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('33987', '77308');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('20124', '65724');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('80491', '84529');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('64123', '84561');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('2386', '76290');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('r', 'RV');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('56655', '65439');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('54814', '80892');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('55220', '64181');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('55130', '65303');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81085', '84780');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('27654', '73657');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83379', '84873');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('4594', '82482');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('81721', '84200');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('77640', '85387');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('39467', '53843');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('26251', '68083');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('47670', '65431');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83287', '84398');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('83287', '85687');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('63641', 'so');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('85088', '85092');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('50648', '64546');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('82019', '83018');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('50817', '80325');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('35570', '80525');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('53768', '85635');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('31037', '63940');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('21949', '64344');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('82124', '85469');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('62112', '84986');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('37477', '85403');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('49322', '49323');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84749', '85349');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('82747', '84806');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('51020', '64692');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('19636', '64391');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('82751', '85677');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('68751', '82807');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('26656', '56719');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('64577', '81543');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('84471', '84945');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('60484', '85032');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('48290', '74597');
    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('60881', '63936');

    PRINT 'Loaded 164 duplicate relationships.';

    -- 2. Repoint Secondary Barcodes in Item_Objects to Master SKU
    PRINT 'Step 1: Repointing secondary barcodes in Item_Objects...';
    UPDATE obj
    SET obj.io_itm_id = m.itm_id
    FROM Item_Objects obj
    INNER JOIN Item_Catalog dup ON obj.io_itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Catalog m ON m.itm_code = map.Master_Code
    WHERE obj.io_itm_id <> m.itm_id;

    -- 3. Consolidate Active Branch Stock Balances in Item_Class_Store
    PRINT 'Step 2: Consolidating branch inventory in Item_Class_Store...';

    -- A. Transfer quantities from duplicate batch records to master batch records where same store exists
    UPDATE master_ics
    SET master_ics.itm_qty = master_ics.itm_qty + dup_ics.itm_qty
    FROM Item_Class_Store master_ics
    INNER JOIN Item_Catalog m ON master_ics.itm_id = m.itm_id
    INNER JOIN #VerifiedDuplicates map ON m.itm_code = map.Master_Code
    INNER JOIN Item_Catalog dup ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Class_Store dup_ics ON dup_ics.itm_id = dup.itm_id 
        AND dup_ics.sto_id = master_ics.sto_id;

    -- B. Zero out inventory on duplicate records
    UPDATE dup_ics
    SET dup_ics.itm_qty = 0
    FROM Item_Class_Store dup_ics
    INNER JOIN Item_Catalog dup ON dup_ics.itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code;

    -- 4. Soft-Deactivate Redundant Catalog Rows
    PRINT 'Step 3: Soft-deactivating redundant items in Item_Catalog...';
    UPDATE cat
    SET cat.itm_active = '0',
        cat.itm_stop_sell = 1,
        cat.itm_stop_pur = 1,
        cat.itm_notes = ISNULL(cat.itm_notes, '') + ' [MERGED INTO MASTER CODE ' + map.Master_Code + ']'
    FROM Item_Catalog cat
    INNER JOIN #VerifiedDuplicates map ON cat.itm_code = map.Duplicate_Code;

    COMMIT TRANSACTION;
    PRINT 'SUCCESS: Catalog deduplication migration completed and committed.';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;
    PRINT 'ERROR OCCURRED: Transaction rolled back successfully.';
    PRINT ERROR_MESSAGE();
END CATCH;
GO
