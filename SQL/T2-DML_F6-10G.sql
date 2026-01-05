SELECT * FROM(
    SELECT
        HDR.STARTTIME,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        HDR.BATCHNUMBER,
        MAX(CASE THM2.MISCDESC WHEN 'RecipeName-Macro' THEN THM2.MISCVALUE END) AS "RecipeName_Macro",
        MAX(CASE THM.MISCDESC WHEN 'RecipeName-Program' THEN THM.MISCVALUE END) AS "RecipeName_Program",
        MAX(CASE THM.MISCDESC WHEN 'RecipeName-Folder' THEN THM.MISCVALUE END) AS "RecipeName_Folder",
        OPS.OPERATIONSTEPNAME || '_' || PARAMETERNAME AS TESTSTEPS,
        MSD.VALUE,
        MAX(CASE TEQ.DEVICENAME WHEN 'SEM' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "SEM",
        MAX(CASE TEQ.DEVICENAME WHEN 'Polaron' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "POLARON",
        MAX(CASE TEQ.DEVICENAME WHEN 'XRD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "XRD",
        MAX(CASE MO.DEVICENAME WHEN 'MOCVD' THEN MO.DEVICENAME || MO.DEVICESERIALNUMBER END) AS "MOCVD"
    FROM
        TDSMFG.TESTHEADER           HDR,
        TDSMFG.TESTHEADERMISC       THM,
        TDSMFG.TESTHEADERMISC       THM2,
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
        AND HDS.TESTHEADERID            = THM2.TESTHEADERID
        AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
        AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
        AND MSP.PARAMETERID             = PAR.PARAMETERID
        AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND MO.TESTHEADERID(+)          = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'T2-DML'
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD6'
        AND THM2.MISCVALUE LIKE '%10GDML-HIMEJI-T2.exe'
    GROUP BY
        STARTTIME,
        PARTNUMBER,
        SERIALNUMBER,
        OPERATORID,
        BATCHNUMBER,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE
)
PIVOT(
    MAX(VALUE) FOR TESTSTEPS IN(
        'Coordinate_X' AS X,
        'Coordinate_Y' AS Y,
        'XRayDiffraction_XRayDiffraction_Strain' AS XRayDiffraction_Strain,
        'CarrierConcentration_CarrierConcentration_Clad' AS CarrierConcentration_Clad,
        'CarrierConcentration_CarrierConcentration_Contact' AS CarrierConcentration_Contact,
        'CarrierConcentration_CarrierConcentration_Minimum' AS CarrierConcentration_Minimum,
        'Thickness_Thickness_Contact' AS Thickness_Contact,
        'Thickness_Thickness_Clad' AS Thickness_Clad,
        'Particles_Particles' AS Particles,
        'GrowthRate_GrowthRate' AS GrowthRate,
        'Dulation_Step7' AS Dulation_Step7,
        'Dulation_Step8' AS Dulation_Step8,
        'Dulation_Step11' AS Dulation_Step11,
        'Dulation_Step12' AS Dulation_Step12,
        'Dulation_Step13' AS Dulation_Step13,
        'Dulation_Step14' AS Dulation_Step14,
        'Dulation_Step15' AS Dulation_Step15,
        'Dulation_Step19' AS Dulation_Step19,
        'Dulation_Step23' AS Dulation_Step23,
        'Dulation_Step27' AS Dulation_Step27,
        'Dulation_Step28' AS Dulation_Step28,
        'Dulation_Step29' AS Dulation_Step29,
        'Dulation_Step30' AS Dulation_Step30,
        'Dulation_Step31' AS Dulation_Step31,
        'Dulation_Step32' AS Dulation_Step32,
        'Dulation_Step36' AS Dulation_Step36,
        'MO1-TEG_Step19' AS "MO1-TEG_Step19",
        'MO2-TMI-1_Step19' AS "MO2-TMI-1_Step19",
        'MO3-TEG-1_Step23' AS "MO3-TEG-1_Step23",
        'MO3-TEG-1_Step27' AS "MO3-TEG-1_Step27",
        'MO3-TEG-1_Step28' AS "MO3-TEG-1_Step28",
        'MO3-TEG-1_Step29' AS "MO3-TEG-1_Step29",
        'MO3-TEG-1_Step30' AS "MO3-TEG-1_Step30",
        'MO3-TEG-1_Step31' AS "MO3-TEG-1_Step31",
        'MO3-TEG-1_Step32' AS "MO3-TEG-1_Step32",
        'MO4-TMI-2_Step23' AS "MO4-TMI-2_Step23",
        'MO4-TMI-2_Step27' AS "MO4-TMI-2_Step27",
        'MO4-TMI-2_Step28' AS "MO4-TMI-2_Step28",
        'MO4-TMI-2_Step29' AS "MO4-TMI-2_Step29",
        'MO4-TMI-2_Step30' AS "MO4-TMI-2_Step30",
        'MO4-TMI-2_Step31' AS "MO4-TMI-2_Step31",
        'MO4-TMI-2_Step32' AS "MO4-TMI-2_Step32",
        'MO5-TMI-3_Step11' AS "MO5-TMI-3_Step11",
        'MO5-TMI-3_Step12' AS "MO5-TMI-3_Step12",
        'MO5-TMI-3_Step13' AS "MO5-TMI-3_Step13",
        'MO5-TMI-3_Step14' AS "MO5-TMI-3_Step14",
        'MO5-TMI-3_Step15' AS "MO5-TMI-3_Step15",
        'MO5-TMI-3_Step36' AS "MO5-TMI-3_Step36",
        'AsH3-1-20percent_Step27' AS "AsH3-1-20percent_Step27",
        'AsH3-1-20percent_Step28' AS "AsH3-1-20percent_Step28",
        'AsH3-1-20percent_Step29' AS "AsH3-1-20percent_Step29",
        'AsH3-1-20percent_Step30' AS "AsH3-1-20percent_Step30",
        'AsH3-1-20percent_Step31' AS "AsH3-1-20percent_Step31",
        'AsH3-1-20percent_Step32' AS "AsH3-1-20percent_Step32",
        'AsH3-21-5percent_Step19' AS "AsH3-21-5percent_Step19",
        'AsH3-3-20percent_Step23' AS "AsH3-3-20percent_Step23",
        'PH3-1-50percent_Step7' AS "PH3-1-50percent_Step7",
        'PH3-1-50percent_Step8' AS "PH3-1-50percent_Step8",
        'PH3-1-50percent_Step11' AS "PH3-1-50percent_Step11",
        'PH3-1-50percent_Step12' AS "PH3-1-50percent_Step12",
        'PH3-1-50percent_Step13' AS "PH3-1-50percent_Step13",
        'PH3-1-50percent_Step14' AS "PH3-1-50percent_Step14",
        'PH3-1-50percent_Step15' AS "PH3-1-50percent_Step15",
        'PH3-1-50percent_Step36' AS "PH3-1-50percent_Step36",
        'PH3-2-50percent_Step19' AS "PH3-2-50percent_Step19",
        'PH3-2-50percent_Step23' AS "PH3-2-50percent_Step23",
        'DMZn-1-0.1percent_Step11' AS "DMZn-1-0.1percent_Step11",
        'DMZn-1-0.1percent_Step12' AS "DMZn-1-0.1percent_Step12",
        'DMZn-1-0.1percent_Step13' AS "DMZn-1-0.1percent_Step13",
        'DMZn-1-0.1percent_Step36' AS "DMZn-1-0.1percent_Step36",
        'DMZn-2-0.1percent_Step14' AS "DMZn-2-0.1percent_Step14",
        'DMZn-2-0.1percent_Step15' AS "DMZn-2-0.1percent_Step15",
        'DMZn-2-0.1percent_Step19' AS "DMZn-2-0.1percent_Step19",
        'DMZn-2-0.1percent_Step23' AS "DMZn-2-0.1percent_Step23",
        'DMZn-2-0.1percent_Step27' AS "DMZn-2-0.1percent_Step27",
        'DMZn-2-0.1percent_Step28' AS "DMZn-2-0.1percent_Step28",
        'DMZn-2-0.1percent_Step29' AS "DMZn-2-0.1percent_Step29",
        'DMZn-2-0.1percent_Step30' AS "DMZn-2-0.1percent_Step30",
        'DMZn-2-0.1percent_Step31' AS "DMZn-2-0.1percent_Step31",
        'DMZn-2-0.1percent_Step32' AS "DMZn-2-0.1percent_Step32",
        'CH3Cl_CH3Cl' AS "CH3Cl",
        'Temperature_Step7' AS "Temperature_Step7",
        'Temperature_Step8' AS "Temperature_Step8",
        'Temperature_Step11' AS "Temperature_Step11",
        'Temperature_Step12' AS "Temperature_Step12",
        'Temperature_Step13' AS "Temperature_Step13",
        'Temperature_Step14' AS "Temperature_Step14",
        'Temperature_Step15' AS "Temperature_Step15",
        'Temperature_Step19' AS "Temperature_Step19",
        'Temperature_Step23' AS "Temperature_Step23",
        'Temperature_Step27' AS "Temperature_Step27",
        'Temperature_Step28' AS "Temperature_Step28",
        'Temperature_Step29' AS "Temperature_Step29",
        'Temperature_Step30' AS "Temperature_Step30",
        'Temperature_Step31' AS "Temperature_Step31",
        'Temperature_Step32' AS "Temperature_Step32",
        'Temperature_Step36' AS "Temperature_Step36",
        'Piezocon_F1'AS "Piezocon_F1",
        'Piezocon_F1_Inverse' AS "Piezocon_F1_Inverse",
        'BallastN2_BallastN2' AS "BallastN2",
        'MO-Temperature_MO1-TEG(1)' AS "MO-Temperature_MO1-TEG(1)",
        'MO-Temperature_MO2-TMI(1)' AS "MO-Temperature_MO2-TMI(1)",
        'MO-Temperature_MO3-TEG(2)' AS "MO-Temperature_MO3-TEG(2)",
        'MO-Temperature_MO4-TMI(2)' AS "MO-Temperature_MO4-TMI(2)",
        'MO-Temperature_MO5-TMI(3)' AS "MO-Temperature_MO5-TMI(3)",
        'MO-Temperature_MO6-CBr4' AS "MO-Temperature_MO6-CBr4",
        'MO-Temperature_MO7-Cp2Mg' AS "MO-Temperature_MO7-Cp2Mg",
        'MO-Temperature_MO8-TMA' AS "MO-Temperature_MO8-TMA",
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
            TDSMFG.STRINGPARAMETER      STP,
            TDSMFG.STRINGMEASUREMENT    SMM
        WHERE 1=1
            AND HDR.TESTHEADERID            = HDS.TESTHEADERID
            AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND HDS.TESTHEADERSTEPID        = STP.TESTHEADERSTEPID
            AND STP.STRINGPARAMETERID       = SMM.STRINGPARAMETERID
            AND PAR.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND PAR.PARAMETERID             = STP.PARAMETERID
            AND HDR.OPERATION               = 'T2-DML'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
