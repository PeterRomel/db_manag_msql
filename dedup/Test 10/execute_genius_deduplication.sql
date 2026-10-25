-- =============================================================================
-- GENIUS / E-PLUS PHARMACY ERP: PRODUCTION CATALOG DEDUPLICATION SCRIPT
-- Target Database: [Genius] on Microsoft SQL Server
-- Directives:
--   1. ZERO hard deletes (preserves foreign keys on sales_trans_d and pur_trans_d).
--   2. Secondary barcodes re-linked to Master SKU in Item_Objects.
--   3. Warehouse stock quantities consolidated in Item_Class_Store.
--   4. Duplicate SKUs soft-deactivated (itm_stop_sell=1, itm_stop_pur=1, itm_active='0').
-- =============================================================================

USE [Genius];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;
BEGIN TRY
    -- 1. Create In-Memory Mapping Table
    IF OBJECT_ID('tempdb..#VerifiedDuplicates') IS NOT NULL
        DROP TABLE #VerifiedDuplicates;

    CREATE TABLE #VerifiedDuplicates (
        Duplicate_Code VARCHAR(50) NOT NULL,
        Master_Code VARCHAR(50) NOT NULL,
        PRIMARY KEY CLUSTERED (Duplicate_Code, Master_Code)
    );

    PRINT 'Staging verified duplicate mappings...';
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('68400', '44486');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85169', '84662');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84603', '81715');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('74138', '65630');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65275', '26307');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85370', '79898');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65299', '52989');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('62230289', '67744');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('68253', '64018');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('68252', '44798');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85522', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85523', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85524', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85526', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85527', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85528', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85529', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85531', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85532', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85533', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85534', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85562', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85615', '85521');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('74611', '74147');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84384', '56393');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66454', '44116');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('74447', '59267');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('50813', '18448');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82565', '26725');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64124', '13655');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84887', '36106');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64369', '52778');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85067', '84605');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82045', '80222');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82191', '77027');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('79167', '55372');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66615', '52775');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85183', '29679');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64921', '53863');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64011', '27264');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65574', '64408');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66593', '64408');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('69053', '64408');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('72852', '63644');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84901', '73328');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84047', '84046');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85520', '61691');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85519', '84053');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80923', '80922');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82096', '80922');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('25416', '25146');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66115', '30193');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81899', '70170');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65592', '34307');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('74313', '13259');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65867', '13993');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83141', '83007');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84919', '27919');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84029', '50253');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('67205', '64337');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80600', '80487');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63275', '30531');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63276', '30531');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('76631', '30531');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80343', '30531');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80591', '30531');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85565', '81812');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('75340', '8092');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85490', '83718');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81121', '81119');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81143', '81119');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('35428', '34693');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65460', '56300');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66848', '56644');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64867', '30472');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84870', '84265');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84035', '74229');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65746', '55955');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85614', '55955');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63896', '21037');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63845', '33428');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('75958', '67833');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85421', '85185');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85484', '85185');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85373', '27780');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83688', '57919');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('62230156', '49638');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66076', '59071');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66312', '59071');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('75265', '56364');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64533', '4705');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64072', '49355');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85299', '49603');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84935', '63865');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80977', '80968');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('77016', '57300');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85188', '76868');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('76856', '76739');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64534', '28118');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66556', '27204');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83721', '26301');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('79588', '61267');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80267', '61268');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85044', '83892');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81006', '79902');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('77459', '64225');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('73644', '63483');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85555', '73645');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85184', '30735');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83111', '79125');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('53983', '53982');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('26445', '2635');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('30266', '2635');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('51540', '2658');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('2637', '2636');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63855', '37967');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('59927', '57128');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66620', '34876');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64817', '29295');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64537', '61040');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80251', '80134');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('75661', '74464');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66080', '59273');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85330', '10844');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66154', '54804');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80679', '72292');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81773', '59585');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85324', '76395');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65411', '43591');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85274', '81331');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85488', '80903');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83592', '56198');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('77473', '26342');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('73160', '49026');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80908', '77529');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85334', '77529');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80385', '54242');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64854', '27474');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85552', '83953');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83950', '38191');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('58465', '57830');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('52403', '13385');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('29272', '18648');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('29296', '18648');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('47808', '7716');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('54297', '7716');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('72450', '7716');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('73227', '7716');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('56774', '54302');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85309', '82042');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82965', '61742');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85225', '40696');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83293', '81537');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85601', '84981');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('79965', '74141');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82296', '80723');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85511', '85295');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85095', '83876');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85498', '83876');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85652', '80997');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85564', '77299');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85660', '85176');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85468', '82614');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66333', '33255');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85673', '85672');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83033', '9431');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('60325', '43741');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65072', '58690');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84030', '32733');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84574', '43554');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('56867', '53408');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84373', '75211');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84671', '55912');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80767', '79537');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80337', '31836');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85698', '85007');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84912', '78314');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85168', '61367');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65732', '33183');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82219', '30277');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84985', '84577');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('51862', '50729');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('52643', '50729');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('17465', '1595');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('42516', '1595');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('53347', '53346');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('72161', '47598');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('70059', '63734');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81815', '63734');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64815', '62612');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('P4', 'p3');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('P5', 'p3');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('p2', 'P1');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81297', '72484');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('66709', '1894');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('74006', '22785');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('51604', '37528');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('77265', '72245');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64204', '30165');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64998', '31967');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84681', '31967');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81528', '34001');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85056', '48315');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('77308', '33987');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65724', '20124');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84529', '80491');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84561', '64123');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85132', '62477');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('76290', '2386');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('RV', 'r');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80892', '54814');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64181', '55220');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65303', '55130');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84780', '81085');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('73657', '27654');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84873', '83379');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82482', '4594');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84200', '81721');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85387', '77640');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('53843', '39467');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('68083', '26251');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('65431', '47670');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83287', '63641');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84398', '63641');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85687', '63641');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('so', '63641');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85092', '85088');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64546', '50648');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('67818', '64733');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83018', '82019');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('77822', '63835');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('d20', '53726');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('d25', '53726');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80325', '50817');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('80525', '35570');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83251', '74861');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63940', '31037');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85635', '53768');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('83639', '53648');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64344', '21949');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81350', '68658');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84986', '62112');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85403', '37477');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('79120', '70678');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('49323', '49322');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85349', '84749');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84806', '82747');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64692', '51020');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64391', '19636');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('82807', '68751');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('64995', '63617');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('56719', '26656');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('85497', '85496');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('81543', '64577');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('84945', '84471');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('74597', '48290');
    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('63936', '60881');

    PRINT 'Loaded 257 duplicate mappings into staging table.';

    -- 2. Migrate Secondary Barcodes in Item_Objects
    PRINT 'Step 1: Re-pointing secondary barcodes in Item_Objects to Master SKU...';
    UPDATE obj
    SET obj.io_itm_id = master_cat.itm_id
    FROM Item_Objects obj
    INNER JOIN Item_Catalog dup_cat ON obj.io_itm_id = dup_cat.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup_cat.itm_code = map.Duplicate_Code
    INNER JOIN Item_Catalog master_cat ON master_cat.itm_code = map.Master_Code
    WHERE obj.io_itm_id <> master_cat.itm_id;

    -- 3. Consolidate Active Branch Stock Balances in Item_Class_Store
    PRINT 'Step 2: Consolidating branch stock balances in Item_Class_Store...';
    
    -- A. Add duplicate quantity to master record where batch store record exists
    UPDATE master_ics
    SET master_ics.itm_qty = master_ics.itm_qty + dup_ics.itm_qty
    FROM Item_Class_Store master_ics
    INNER JOIN Item_Catalog master_cat ON master_ics.itm_id = master_cat.itm_id
    INNER JOIN #VerifiedDuplicates map ON master_cat.itm_code = map.Master_Code
    INNER JOIN Item_Catalog dup_cat ON dup_cat.itm_code = map.Duplicate_Code
    INNER JOIN Item_Class_Store dup_ics ON dup_ics.itm_id = dup_cat.itm_id 
        AND dup_ics.sto_id = master_ics.sto_id;

    -- B. Zero out inventory balances on duplicate items
    UPDATE dup_ics
    SET dup_ics.itm_qty = 0
    FROM Item_Class_Store dup_ics
    INNER JOIN Item_Catalog dup_cat ON dup_ics.itm_id = dup_cat.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup_cat.itm_code = map.Duplicate_Code;

    -- 4. Soft-Deactivate Redundant Catalog Rows in Item_Catalog
    PRINT 'Step 3: Soft-deactivating redundant records in Item_Catalog...';
    UPDATE cat
    SET cat.itm_active = '0',
        cat.itm_stop_sell = 1,
        cat.itm_stop_pur = 1,
        cat.itm_notes = ISNULL(cat.itm_notes, '') + ' [MERGED INTO MASTER CODE ' + map.Master_Code + ']'
    FROM Item_Catalog cat
    INNER JOIN #VerifiedDuplicates map ON cat.itm_code = map.Duplicate_Code;

    COMMIT TRANSACTION;
    PRINT 'Deduplication migration completed successfully. Transaction committed.';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;

    PRINT 'ERROR OCCURRED DURING MIGRATION! Transaction rolled back.';
    PRINT ERROR_MESSAGE();
END CATCH;
GO
