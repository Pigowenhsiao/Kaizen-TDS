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
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD5'
        AND THM2.MISCVALUE LIKE '%13BFCP50-T2.exe'
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
        'Dulation_Step6' AS Dulation_Step6,
        'Dulation_Step7' AS Dulation_Step7,
        'Dulation_Step10' AS Dulation_Step10,
        'Dulation_Step11' AS Dulation_Step11,
        'Dulation_Step12' AS Dulation_Step12,
        'Dulation_Step13' AS Dulation_Step13,
        'Dulation_Step14' AS Dulation_Step14,
        'Dulation_Step15' AS Dulation_Step15,
        'Dulation_Step21' AS Dulation_Step21,
        'Dulation_Step27' AS Dulation_Step27,
        'Dulation_Step33' AS Dulation_Step33,
        'Dulation_Step34' AS Dulation_Step34,
        'Dulation_Step35' AS Dulation_Step35,
        'Dulation_Step36' AS Dulation_Step36,
        'Dulation_Step41' AS Dulation_Step41,
        'MO1-TEG_Step21' AS "MO1-TEG_Step21",
        'MO1-TEG_Step27' AS "MO1-TEG_Step27",
        'MO1-TEG_Step33' AS "MO1-TEG_Step33",
        'MO1-TEG_Step34' AS "MO1-TEG_Step34",
        'MO1-TEG_Step35' AS "MO1-TEG_Step35",
        'MO1-TEG_Step36' AS "MO1-TEG_Step36",
        'MO9-TMI_Step10' AS "MO9-TMI_Step10",
        'MO9-TMI_Step11' AS "MO9-TMI_Step11",
        'MO9-TMI_Step12' AS "MO9-TMI_Step12",
        'MO9-TMI_Step13' AS "MO9-TMI_Step13",
        'MO9-TMI_Step14' AS "MO9-TMI_Step14",
        'MO9-TMI_Step15' AS "MO9-TMI_Step15",
        'MO9-TMI_Step21' AS "MO9-TMI_Step21",
        'MO9-TMI_Step27' AS "MO9-TMI_Step27",
        'MO9-TMI_Step33' AS "MO9-TMI_Step33",
        'MO9-TMI_Step34' AS "MO9-TMI_Step34",
        'MO9-TMI_Step35' AS "MO9-TMI_Step35",
        'MO9-TMI_Step36' AS "MO9-TMI_Step36",
        'MO9-TMI_Step41' AS "MO9-TMI_Step41",
        'AsH3-A-20percent_Step27' AS "AsH3-1-20percent_Step27",
        'AsH3-A-20percent_Step33' AS "AsH3-1-20percent_Step33",
        'AsH3-A-20percent_Step34' AS "AsH3-1-20percent_Step34",
        'AsH3-A-20percent_Step35' AS "AsH3-1-20percent_Step35",
        'AsH3-A-20percent_Step36' AS "AsH3-1-20percent_Step36",
        'PH3-50percent_Step6' AS "PH3-50percent_Step6",
        'PH3-50percent_Step7' AS "PH3-50percent_Step7",
        'PH3-50percent_Step10' AS "PH3-50percent_Step10",
        'PH3-50percent_Step11' AS "PH3-50percent_Step11",
        'PH3-50percent_Step12' AS "PH3-50percent_Step12",
        'PH3-50percent_Step13' AS "PH3-50percent_Step13",
        'PH3-50percent_Step14' AS "PH3-50percent_Step14",
        'PH3-50percent_Step15' AS "PH3-50percent_Step15",
        'PH3-50percent_Step21' AS "PH3-50percent_Step21",
        'PH3-50percent_Step27' AS "PH3-50percent_Step27",
        'PH3-50percent_Step41' AS "PH3-50percent_Step41",
        'CH3Cl-5percent_Step7' AS "CH3Cl-5percent_Step7",
        'Si2H6-10ppm_Step41' AS "Si6H6-10ppm_Step41",
        'DMZn-B-0.1percent_Step10' AS "DMZn-B-0.1percent_Step10",
        'DMZn-B-0.1percent_Step11' AS "DMZn-B-0.1percent_Step11",
        'DMZn-B-0.1percent_Step12' AS "DMZn-B-0.1percent_Step12",
        'DMZn-B-0.1percent_Step13' AS "DMZn-B-0.1percent_Step13",
        'DMZn-B-0.1percent_Step14' AS "DMZn-B-0.1percent_Step14",
        'DMZn-B-0.1percent_Step15' AS "DMZn-B-0.1percent_Step15",
        'DMZn-B-0.1percent_Step21' AS "DMZn-B-0.1percent_Step21",
        'DMZn-B-0.1percent_Step27' AS "DMZn-B-0.1percent_Step27",
        'DMZn-B-0.1percent_Step33' AS "DMZn-B-0.1percent_Step33",
        'DMZn-B-0.1percent_Step34' AS "DMZn-B-0.1percent_Step34",
        'DMZn-B-0.1percent_Step35' AS "DMZn-B-0.1percent_Step35",
        'DMZn-B-0.1percent_Step36' AS "DMZn-B-0.1percent_Step36",
        'AsH3-5percent_Step21' AS "AsH3-5percent_Step21",
        'Temperature_Step6' AS "Temperature_Step6",
        'Temperature_Step7' AS "Temperature_Step7",
        'Temperature_Step10' AS "Temperature_Step10",
        'Temperature_Step11' AS "Temperature_Step11",
        'Temperature_Step12' AS "Temperature_Step12",
        'Temperature_Step13' AS "Temperature_Step13",
        'Temperature_Step14' AS "Temperature_Step14",
        'Temperature_Step15' AS "Temperature_Step15",
        'Temperature_Step21' AS "Temperature_Step21",
        'Temperature_Step27' AS "Temperature_Step27",
        'Temperature_Step33' AS "Temperature_Step33",
        'Temperature_Step34' AS "Temperature_Step34",
        'Temperature_Step35' AS "Temperature_Step35",
        'Temperature_Step36' AS "Temperature_Step36",
        'Temperature_Step41' AS "Temperature_Step41",
        'Piezocon_F1'AS "Piezocon_F1",
        'Piezocon_F1_Inverse' AS "Piezocon_F1_Inverse",
        'Piezocon_F5_Start' AS "Piezocon_F5_Start",
        'Piezocon_F5_Current' AS "Piezocon_F5_Current",
        'BallastN2_BallastN2' AS "BallastN2",
        'CH3Cl-Judge_StartTime' AS "CH3Cl-Judge_StartTime",
        'CH3Cl-Judge_StablWaitTime' AS "CH3Cl-Judge_StablWaitTime",
        'CH3Cl-Judge_EtchingTime' AS "CH3Cl-Judge_EtchingTime",
        'CH3Cl-Judge_EtchingStartTemperature' AS "CH3Cl-Judge_EtchingStartTemperature",
        'MO-Temperature_MO1-TEG(A)' AS "MO-Temperature_MO1-TEG(A)",
        'MO-Temperature_MO2-TMA(A)' AS "MO-Temperature_MO2-TMI(A)",
        'MO-Temperature_MO3-DEZ' AS "MO-Temperature_MO3-DEZ",
        'MO-Temperature_MO4-TMI(A)' AS "MO-Temperature_MO4-TMI(A)",
        'MO-Temperature_MO5-TMI(B)' AS "MO-Temperature_MO5-TMI(B)",
        'MO-Temperature_MO6-TEG(B)' AS "MO-Temperature_MO6-TEG(B)",
        'MO-Temperature_MO7-TMA(B)' AS "MO-Temperature_MO7-TMA(B)",
        'MO-Temperature_MO8-CBr4' AS "MO-Temperature_MO8-CBr4",
        'MO-Temperature_MO9-TMI' AS "MO-Temperature_MO9-TMI",
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
