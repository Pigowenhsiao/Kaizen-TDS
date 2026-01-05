SELECT 
    Int_Data.STARTTIME,
    Int_Data.PARTNUMBER,
    Int_Data.SERIALNUMBER,
    String_Data."NINE_SERIALNUMBER",
    Int_Data.OPERATORID,
    Int_Data.BATCHNUMBER,
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    Int_Data."SORTED_DATA_SORTNUMBER",
    Int_Data."MOCVD",
    Int_Data."SEM",
    Int_Data."RecipeName_Macro",
    Int_Data."RecipeName_Program",
    Int_Data."RecipeName_Folder",
    Int_Data."X",
    Int_Data."Y",
    Int_Data."Particles_All_Address",
    Int_Data."Particles_Center",
    Int_Data."Particles_NG_Count",
    String_Data."Particles_NG_Address",
    Int_Data."Particles",
    Int_Data."BallastN2",
    Int_Data."MO-Temperature_MO1-Ru",
    Int_Data."MO-Temperature_MO2-Fe",
    Int_Data."MO-Temperature_MO3-TMI",
    Int_Data."MO-Temperature_MO4-TMA-A",
    Int_Data."MO-Temperature_MO5-TEG-A",
    Int_Data."Dulation_Step7",
    Int_Data."Dulation_Step8",
    Int_Data."Dulation_Step9",
    Int_Data."Dulation_Step10",
    Int_Data."Dulation_Step11",
    Int_Data."MO1-Ru_Step10",
    Int_Data."MO1-Ru_Step11",
    Int_Data."MO3-TMI_Step10",
    Int_Data."MO3-TMI_Step11",
    Int_Data."CH3Cl_Step7",
    Int_Data."CH3Cl_Step8",
    Int_Data."CH3Cl_Step9",
    Int_Data."CH3Cl_Step10",
    Int_Data."CH3Cl_Step11",
    Int_Data."PH3_Step7",
    Int_Data."PH3_Step8",
    Int_Data."PH3_Step9",
    Int_Data."PH3_Step10",
    Int_Data."PH3_Step11",
    Int_Data."Temperature_Step7",
    Int_Data."Temperature_Step8",
    Int_Data."Temperature_Step9",
    Int_Data."Temperature_Step10",
    Int_Data."Temperature_Step11",
    Int_Data."TMInPiezoconConc_Check",
    Int_Data."TMInPiezoconConc_EPI",
    Int_Data."TMInPiezoconConc_CheckFlow",
    Int_Data."CH3Cl_ConcBonbe",
    Int_Data."Thickness_Total",
    Int_Data."Thickness_Ru-InP",
    Int_Data."Thickness_Cap"
FROM(
    SELECT
        HDR.TESTHEADERID,
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
        MAX(CASE TEQ.DEVICENAME WHEN 'MOCVD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "MOCVD"
    FROM
        TDSMFG.TESTHEADER           HDR,
        TDSMFG.TESTHEADERMISC       THM,
        TDSMFG.TESTHEADERSTEP       HDS,
        TDSMFG.OPERATIONSTEP        OPS,
        TDSMFG.MEASUREMENTPARAMETER MSP,
        TDSMFG.MEASUREMENT          MSD,
        TDSMFG.PARAMETER            PAR,
        TDSMFG.TESTEQUIPMENT        TEQ
    WHERE 1=1
        AND HDR.TESTHEADERID            = HDS.TESTHEADERID
        AND HDR.TESTHEADERID            = THM.TESTHEADERID
        AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
        AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
        AND MSP.PARAMETERID             = PAR.PARAMETERID
        AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'Ru-EML_F4ICP'
    GROUP BY
        HDR.TESTHEADERID,
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
        'Coordinate_X' AS "X",
        'Coordinate_Y' AS "Y",
        'Particles_Particles_All_Address' AS "Particles_All_Address",
        'Particles_Particles_Center' AS "Particles_Center",
        'Particles_Particles_NG_Count' AS "Particles_NG_Count",
        'Particles_Particles' AS "Particles",
        'BallastN2_BallastN2' AS "BallastN2",
        'MO-Temperature_MO1-Ru' AS "MO-Temperature_MO1-Ru",
        'MO-Temperature_MO2-Fe' AS "MO-Temperature_MO2-Fe",
        'MO-Temperature_MO3-TMI' AS "MO-Temperature_MO3-TMI",
        'MO-Temperature_MO4-TMA-A' AS "MO-Temperature_MO4-TMA-A",
        'MO-Temperature_MO5-TEG-A' AS "MO-Temperature_MO5-TEG-A",
        'Dulation_Step7' AS "Dulation_Step7",
        'Dulation_Step8' AS "Dulation_Step8",
        'Dulation_Step9' AS "Dulation_Step9",
        'Dulation_Step10' AS "Dulation_Step10",
        'Dulation_Step11' AS "Dulation_Step11",
        'MO1-Ru_Step10' AS "MO1-Ru_Step10",
        'MO1-Ru_Step11' AS "MO1-Ru_Step11",
        'MO3-TMI_Step10' AS "MO3-TMI_Step10",
        'MO3-TMI_Step11' AS "MO3-TMI_Step11",
        'CH3Cl_Step7' AS "CH3Cl_Step7",
        'CH3Cl_Step8' AS "CH3Cl_Step8",
        'CH3Cl_Step9' AS "CH3Cl_Step9",
        'CH3Cl_Step10' AS "CH3Cl_Step10",
        'CH3Cl_Step11' AS "CH3Cl_Step11",
        'PH3_Step7' AS "PH3_Step7",
        'PH3_Step8' AS "PH3_Step8",
        'PH3_Step9' AS "PH3_Step9",
        'PH3_Step10' AS "PH3_Step10",
        'PH3_Step11' AS "PH3_Step11",
        'Temperature_Step7' AS "Temperature_Step7",
        'Temperature_Step8' AS "Temperature_Step8",
        'Temperature_Step9' AS "Temperature_Step9",
        'Temperature_Step10' AS "Temperature_Step10",
        'Temperature_Step11' AS "Temperature_Step11",
        'TMInPiezoconConc_TMInPiezoconConc_Check' AS "TMInPiezoconConc_Check",
        'TMInPiezoconConc_TMInPiezoconConc_EPI' AS "TMInPiezoconConc_EPI",
        'TMInPiezoconConc_TMInPiezoconConc_CheckFlow' AS "TMInPiezoconConc_CheckFlow",
        'CH3Cl_Conc_CH3Cl_ConcBonbe' AS "CH3Cl_ConcBonbe",
        'Thickness_Thickness_Total' AS "Thickness_Total",
        'Thickness_Thickness_Ru-InP' AS "Thickness_Ru-InP",
        'Thickness_Thickness_Cap' AS "Thickness_Cap",
        'SORTED_DATA_STARTTIME_SORTED' AS "SORTED_DATA_STARTTIME_SORTED",
        'SORTED_DATA_SORTNUMBER' AS "SORTED_DATA_SORTNUMBER"
    )
)Int_Data
,(
    SELECT
        HDR.TESTHEADERID,
        HDR.SERIALNUMBER,
        MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'Particles_Particles_NG_Address' THEN VALUESTRING END) AS "Particles_NG_Address",
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
        AND HDR.OPERATION               = 'Ru-EML_F4ICP'
    GROUP BY
        HDR.TESTHEADERID,
        HDR.SERIALNUMBER
)String_Data
WHERE 1=1
  AND Int_Data.TESTHEADERID = String_Data.TESTHEADERID
ORDER BY
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
