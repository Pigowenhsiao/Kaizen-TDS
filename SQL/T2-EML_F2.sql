SELECT * FROM(
    SELECT
        HDR.STARTTIME,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        HDR.BATCHNUMBER,
        MAX(CASE MISCDESC WHEN 'RecipeName-Macro' THEN MISCVALUE END) AS "RecipeName_Macro",
        MAX(CASE MISCDESC WHEN 'RecipeName-Program' THEN MISCVALUE END) AS "RecipeName_Program",
        MAX(CASE MISCDESC WHEN 'RecipeName-Folder' THEN MISCVALUE END) AS "RecipeName_Folder",
        OPS.OPERATIONSTEPNAME || '_' || PARAMETERNAME AS TESTSTEPS,
        MSD.VALUE,
        MAX(CASE TEQ.DEVICENAME WHEN 'SEM' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "SEM",
        MAX(CASE TEQ.DEVICENAME WHEN 'Polaron' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "Polaron",
        MAX(CASE MO.DEVICENAME WHEN 'MOCVD' THEN MO.DEVICENAME || MO.DEVICESERIALNUMBER END) AS "MOCVD"
    FROM
        TDSMFG.TESTHEADER           HDR,
        TDSMFG.TESTHEADERMISC       THM,
        TDSMFG.TESTHEADERSTEP       HDS,
        TDSMFG.OPERATIONSTEP        OPS,
        TDSMFG.MEASUREMENTPARAMETER MSP,
        TDSMFG.MEASUREMENT          MSD,
        TDSMFG.PARAMETER            PAR,
        TDSMFG.TESTEQUIPMENT        TEQ,
        TDSMFG.TESTEQUIPMENT        MO
    WHERE 1=1
        AND HDR.TESTHEADERID            = HDS.TESTHEADERID
        AND HDS.TESTHEADERID            = THM.TESTHEADERID
        AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
        AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
        AND MSP.PARAMETERID             = PAR.PARAMETERID
        AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND MO.TESTHEADERID(+)          = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'T2-EML'
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD2'
    GROUP BY
        STARTTIME,
        PARTNUMBER,
        SERIALNUMBER,
        BATCHNUMBER,
        OPERATORID,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE
)
PIVOT(
    MAX(VALUE) FOR TESTSTEPS IN(
        'Coordinate_X' AS "X",
        'Coordinate_Y' AS "Y",
        'CarrierConcentration_CarrierConcentration_Clad' AS "CarrierConcentration_Clad",
        'CarrierConcentration_CarrierConcentration_Contact' AS "CarrierConcentration_Contact",
        'CarrierConcentration_CarrierConcentration_Minimum' AS "CarrierConcentration_Minimum",
        'Thickness_Thickness_Clad' AS "Thickness_Clad",
        'Thickness_Thickness_Contact' AS "Thickness_Contact",
        'Dulation_Step9' AS "Dulation_Step9",
        'Dulation_Step10' AS "Dulation_Step10",
        'Dulation_Step11' AS "Dulation_Step11",
        'Dulation_Step12' AS "Dulation_Step12",
        'Dulation_Step13' AS "Dulation_Step13",
        'Dulation_Step14' AS "Dulation_Step14",
        'Dulation_Step20' AS "Dulation_Step20",
        'Dulation_Step26' AS "Dulation_Step26",
        'Dulation_Step32' AS "Dulation_Step32",
        'Dulation_Step33' AS "Dulation_Step33",
        'Dulation_Step34' AS "Dulation_Step34",
        'Dulation_Step35' AS "Dulation_Step35",
        'Dulation_Step40' AS "Dulation_Step40",
        'Dulation_Step42' AS "Dulation_Step42",
        'MO1-TMI_Step20' AS "MO1-TMI_Step20",
        'MO2-TEG_Step20' AS "MO2-TEG_Step20",
        'MO3-TMI_Step26' AS "MO3-TMI_Step26",
        'MO3-TMI_Step32' AS "MO3-TMI_Step32",
        'MO3-TMI_Step33' AS "MO3-TMI_Step33",
        'MO3-TMI_Step34' AS "MO3-TMI_Step34",
        'MO3-TMI_Step35' AS "MO3-TMI_Step35",
        'MO4-TEG_Step26' AS "MO4-TEG_Step26",
        'MO4-TEG_Step32' AS "MO4-TEG_Step32",
        'MO4-TEG_Step33' AS "MO4-TEG_Step33",
        'MO4-TEG_Step34' AS "MO4-TEG_Step34",
        'MO4-TEG_Step35' AS "MO4-TEG_Step35",
        'MO5-TMI_Step9' AS "MO5-TMI_Step9",
        'MO5-TMI_Step10' AS "MO5-TMI_Step10",
        'MO5-TMI_Step11' AS "MO5-TMI_Step11",
        'MO5-TMI_Step12' AS "MO5-TMI_Step12",
        'MO5-TMI_Step13' AS "MO5-TMI_Step13",
        'MO5-TMI_Step14' AS "MO5-TMI_Step14",
        'MO5-TMI_Step40' AS "MO5-TMI_Step40",
        'AsH3-A-20percent_Step26' AS "AsH3-A-20percent_Step26",
        'AsH3-A-20percent_Step32' AS "AsH3-A-20percent_Step32",
        'AsH3-A-20percent_Step33' AS "AsH3-A-20percent_Step33",
        'AsH3-A-20percent_Step34' AS "AsH3-A-20percent_Step34",
        'AsH3-A-20percent_Step35' AS "AsH3-A-20percent_Step35",
        'AsH3-5percent_AsH3-5percent-Step20' AS "AsH3-5percent_AsH3-5percent-Step20",
        'PH3-A-50percent_Step9' AS "PH3-A-50percent_Step9",
        'PH3-A-50percent_Step10' AS "PH3-A-50percent_Step10",
        'PH3-A-50percent_Step11' AS "PH3-A-50percent_Step11",
        'PH3-A-50percent_Step12' AS "PH3-A-50percent_Step12",
        'PH3-A-50percent_Step13' AS "PH3-A-50percent_Step13",
        'PH3-A-50percent_Step14' AS "PH3-A-50percent_Step14",
        'PH3-A-50percent_Step20' AS "PH3-A-50percent_Step20",
        'PH3-A-50percent_Step26' AS "PH3-A-50percent_Step26",
        'PH3-A-50percent_Step40' AS "PH3-A-50percent_Step40",
        'PH3-A-50percent_Step42' AS "PH3-A-50percent_Step42",
        'DMZn-A-0.1percent_Step13' AS "DMZn-A-0.1percent_Step13",
        'DMZn-A-0.1percent_Step14' AS "DMZn-A-0.1percent_Step14",
        'DMZn-A-0.1percent_Step20' AS "DMZn-A-0.1percent_Step20",
        'DMZn-A-0.1percent_Step26' AS "DMZn-A-0.1percent_Step26",
        'DMZn-A-0.1percent_Step32' AS "DMZn-A-0.1percent_Step32",
        'DMZn-A-0.1percent_Step33' AS "DMZn-A-0.1percent_Step33",
        'DMZn-A-0.1percent_Step34' AS "DMZn-A-0.1percent_Step34",
        'DMZn-A-0.1percent_Step35' AS "DMZn-A-0.1percent_Step35",
        'DMZn-B-0.1percent_Step9' AS "DMZn-B-0.1percent_Step9",
        'DMZn-B-0.1percent_Step10' AS "DMZn-B-0.1percent_Step10",
        'DMZn-B-0.1percent_Step11' AS "DMZn-B-0.1percent_Step11",
        'DMZn-B-0.1percent_Step12' AS "DMZn-B-0.1percent_Step12",
        'Temperature_Step9' AS "Temperature_Step9",
        'Temperature_Step10' AS "Temperature_Step10",
        'Temperature_Step11' AS "Temperature_Step11",
        'Temperature_Step12' AS "Temperature_Step12",
        'Temperature_Step13' AS "Temperature_Step13",
        'Temperature_Step14' AS "Temperature_Step14",
        'Temperature_Step20' AS "Temperature_Step20",
        'Temperature_Step26' AS "Temperature_Step26",
        'Temperature_Step32' AS "Temperature_Step32",
        'Temperature_Step33' AS "Temperature_Step33",
        'Temperature_Step34' AS "Temperature_Step34",
        'Temperature_Step35' AS "Temperature_Step35",
        'Temperature_Step40' AS "Temperature_Step40",
        'Temperature_Step42' AS "Temperature_Step42",
        'Piezocon_F1' AS "Piezocon_F1",
        'Piezocon_F2-Start' AS "Piezocon_F2-Start",
        'Piezocon_F2-Current' AS "Piezocon_F2-Current",
        'Piezocon_F2-T2CH' AS "Piezocon_F2-T2CH",
        'BallastN2_BallastN2' AS "BallastN2_BallastN2",
        'MO-Temperature_MO1-TMI' AS "MO-Temperature_MO1-TMI",
        'MO-Temperature_MO2-TEG' AS "MO-Temperature_MO2-TEG",
        'MO-Temperature_MO3-TMI' AS "MO-Temperature_MO3-TMI",
        'MO-Temperature_MO4-TEG' AS "MO-Temperature_MO4-TEG",
        'MO-Temperature_MO5-TMI' AS "MO-Temperature_MO5-TMI",
        'MO-Temperature_MO6-TEG' AS "MO-Temperature_MO6-TEG",
        'MO-Temperature_Blank' AS "MO-Temperature_Blank",
        'SORTED_DATA_STARTTIME_SORTED' AS "SORTED_DATA_STARTTIME_SORTED",
        'SORTED_DATA_SORTNUMBER' AS "SORTED_DATA_SORTNUMBER"
    )
)Int_Data
LEFT JOIN
    (
        SELECT
            MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_LotNumber_5' THEN VALUESTRING END) AS "FIVE_SERIALNUMBER",
            MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_LotNumber_9' THEN VALUESTRING END) AS "NINE_SERIALNUMBER"
        FROM
            TDSMFG.TESTHEADER           HDR,
            TDSMFG.TESTHEADERSTEP       HDS,
            TDSMFG.OPERATIONSTEP        OPS,
            TDSMFG.PARAMETER            PAR,
            TDSMFG.TESTEQUIPMENT        TEQ,
            TDSMFG.STRINGPARAMETER      STP,
            TDSMFG.STRINGMEASUREMENT    SMM
        WHERE 1=1
            AND HDR.TESTHEADERID            = HDS.TESTHEADERID
            AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND HDS.TESTHEADERSTEPID        = STP.TESTHEADERSTEPID
            AND STP.STRINGPARAMETERID       = SMM.STRINGPARAMETERID
            AND PAR.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND PAR.PARAMETERID             = STP.PARAMETERID
            AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
            AND HDR.OPERATION               = 'T2-EML'
            AND TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER = 'MOCVD2'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    "SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
