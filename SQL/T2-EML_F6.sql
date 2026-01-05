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
        MAX(CASE TEQ.DEVICENAME WHEN 'XRD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "XRD",
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
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD6'
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
        'XRayDiffraction_XRayDiffraction_Strain' AS "XRayDiffraction_Strain",
        'CarrierConcentration_CarrierConcentration_Clad' AS "CarrierConcentration_Clad",
        'CarrierConcentration_CarrierConcentration_Contact' AS "CarrierConcentration_Contact",
        'CarrierConcentration_CarrierConcentration_Minimum' AS "CarrierConcentration_Minimum",
        'Thickness_Thickness_Clad' AS "Thickness_Clad",
        'Thickness_Thickness_Contact' AS "Thickness_Contact",
        'Particles_Particles' AS "Particles",
        'Dulation_Step9' AS "Dulation_Step9",
        'Dulation_Step10' AS "Dulation_Step10",
        'Dulation_Step11' AS "Dulation_Step11",
        'Dulation_Step12' AS "Dulation_Step12",
        'Dulation_Step13' AS "Dulation_Step13",
        'Dulation_Step14' AS "Dulation_Step14",
        'Dulation_Step18' AS "Dulation_Step18",
        'Dulation_Step22' AS "Dulation_Step22",
        'Dulation_Step26' AS "Dulation_Step26",
        'Dulation_Step27' AS "Dulation_Step27",
        'Dulation_Step28' AS "Dulation_Step28",
        'Dulation_Step29' AS "Dulation_Step29",
        'Dulation_Step30' AS "Dulation_Step30",
        'Dulation_Step31' AS "Dulation_Step31",
        'Dulation_Step35' AS "Dulation_Step35",
        'Dulation_Step37' AS "Dulation_Step37",
        'MO1-TEG_Step18' AS "MO1-TEG_Step18",
        'MO2-TMI_Step18' AS "MO2-TMI_Step18",
        'MO3-TEG_Step22' AS "MO3-TEG_Step22",
        'MO3-TEG_Step26' AS "MO3-TEG_Step26",
        'MO3-TEG_Step27' AS "MO3-TEG_Step27",
        'MO3-TEG_Step28' AS "MO3-TEG_Step28",
        'MO3-TEG_Step29' AS "MO3-TEG_Step29",
        'MO3-TEG_Step30' AS "MO3-TEG_Step30",
        'MO3-TEG_Step31' AS "MO3-TEG_Step31",
        'MO4-TMI_Step22' AS "MO4-TMI_Step22",
        'MO4-TMI_Step26' AS "MO4-TMI_Step26",
        'MO4-TMI_Step27' AS "MO4-TMI_Step27",
        'MO4-TMI_Step28' AS "MO4-TMI_Step28",
        'MO4-TMI_Step29' AS "MO4-TMI_Step29",
        'MO4-TMI_Step30' AS "MO4-TMI_Step30",
        'MO4-TMI_Step31' AS "MO4-TMI_Step31",
        'MO5-TMI_Step9' AS "MO5-TMI_Step9",
        'MO5-TMI_Step10' AS "MO5-TMI_Step10",
        'MO5-TMI_Step11' AS "MO5-TMI_Step11",
        'MO5-TMI_Step12' AS "MO5-TMI_Step12",
        'MO5-TMI_Step13' AS "MO5-TMI_Step13",
        'MO5-TMI_Step14' AS "MO5-TMI_Step14",
        'MO5-TMI_Step35' AS "MO5-TMI_Step35",
        'AsH3-1-20percent_Step26' AS "AsH3-1-20percent_Step26",
        'AsH3-1-20percent_Step27' AS "AsH3-1-20percent_Step27",
        'AsH3-1-20percent_Step28' AS "AsH3-1-20percent_Step28",
        'AsH3-1-20percent_Step29' AS "AsH3-1-20percent_Step29",
        'AsH3-1-20percent_Step30' AS "AsH3-1-20percent_Step30",
        'AsH3-1-20percent_Step31' AS "AsH3-1-20percent_Step31",
        'AsH3-21-5percent_Step18' AS "AsH3-21-5percent_Step18",
        'AsH3-3-20percent_Step22' AS "AsH3-3-20percent_Step22",
        'PH3-1-50percent_Step9' AS "PH3-1-50percent_Step9",
        'PH3-1-50percent_Step10' AS "PH3-1-50percent_Step10",
        'PH3-1-50percent_Step11' AS "PH3-1-50percent_Step11",
        'PH3-1-50percent_Step12' AS "PH3-1-50percent_Step12",
        'PH3-1-50percent_Step13' AS "PH3-1-50percent_Step13",
        'PH3-1-50percent_Step14' AS "PH3-1-50percent_Step14",
        'PH3-1-50percent_Step35' AS "PH3-1-50percent_Step35",
        'PH3-1-50percent_Step37' AS "PH3-1-50percent_Step37",
        'PH3-2-50percent_Step18' AS "PH3-2-50percent_Step18",
        'PH3-2-50percent_Step22' AS "PH3-2-50percent_Step22",
        'DMZn-1-0.1percent_Step9' AS "DMZn-1-0.1percent_Step9",
        'DMZn-1-0.1percent_Step10' AS "DMZn-1-0.1percent_Step10",
        'DMZn-1-0.1percent_Step11' AS "DMZn-1-0.1percent_Step11",
        'DMZn-1-0.1percent_Step12' AS "DMZn-1-0.1percent_Step12",
        'DMZn-2-0.1percent_Step13' AS "DMZn-2-0.1percent_Step13",
        'DMZn-2-0.1percent_Step14' AS "DMZn-2-0.1percent_Step14",
        'DMZn-2-0.1percent_Step18' AS "DMZn-2-0.1percent_Step18",
        'DMZn-2-0.1percent_Step22' AS "DMZn-2-0.1percent_Step22",
        'DMZn-2-0.1percent_Step26' AS "DMZn-2-0.1percent_Step26",
        'DMZn-2-0.1percent_Step27' AS "DMZn-2-0.1percent_Step27",
        'DMZn-2-0.1percent_Step28' AS "DMZn-2-0.1percent_Step28",
        'DMZn-2-0.1percent_Step29' AS "DMZn-2-0.1percent_Step29",
        'DMZn-2-0.1percent_Step30' AS "DMZn-2-0.1percent_Step30",
        'DMZn-2-0.1percent_Step31' AS "DMZn-2-0.1percent_Step31",
        'Temperature_Step9' AS "Temperature_Step9",
        'Temperature_Step10' AS "Temperature_Step10",
        'Temperature_Step11' AS "Temperature_Step11",
        'Temperature_Step12' AS "Temperature_Step12",
        'Temperature_Step13' AS "Temperature_Step13",
        'Temperature_Step14' AS "Temperature_Step14",
        'Temperature_Step18' AS "Temperature_Step18",
        'Temperature_Step22' AS "Temperature_Step22",
        'Temperature_Step26' AS "Temperature_Step26",
        'Temperature_Step27' AS "Temperature_Step27",
        'Temperature_Step28' AS "Temperature_Step28",
        'Temperature_Step29' AS "Temperature_Step29",
        'Temperature_Step30' AS "Temperature_Step30",
        'Temperature_Step31' AS "Temperature_Step31",
        'Temperature_Step35' AS "Temperature_Step35",
        'Temperature_Step37' AS "Temperature_Step37",
        'Piezocon_F1' AS "F1",
        'Piezocon_F1_Inverse' AS "F1_Inverse",
        'BallastN2_BallastN2' AS "BallastN2",
        'MO-Temperature_MO1-TEG' AS "MO1-TEG",
        'MO-Temperature_MO2-TMI' AS "MO2-TMI",
        'MO-Temperature_MO3-TEG' AS "MO3-TEG",
        'MO-Temperature_MO4-TMI' AS "MO4-TMI",
        'MO-Temperature_MO5-TMI' AS "MO5-TMI",
        'MO-Temperature_MO6-CBr4' AS "MO6-CBr4",
        'MO-Temperature_MO7-Cp2Mg' AS "MO7-Cp2Mg",
        'MO-Temperature_MO8-TMA' AS "MO8-TMA",
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
            AND TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER = 'MOCVD6'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
