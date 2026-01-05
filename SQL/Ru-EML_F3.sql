SELECT 
    Int_Data.STARTTIME,
    Int_Data.PARTNUMBER,
    Int_Data.SERIALNUMBER,
    String_Data."NINE_SERIALNUMBER",
    Int_Data.OPERATORID,
    Int_Data.BATCHNUMBER,
    Int_Data.OPERATION,
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    Int_Data."SORTED_DATA_SORTNUMBER",
    Int_Data."MOCVD",
    Int_Data."SEM",
    Int_Data."RecipeName_Macro",
    Int_Data."RecipeName_Program",
    Int_Data."RecipeName_Folder",
    Int_Data."Particles.Particles_All_Address",
    Int_Data."Particles.Particles_Center",
    Int_Data."Particles.Particles_NG_Count",
    String_Data."Particles_NG_Address",
    Int_Data."Particles.Particles",
    Int_Data."Thickness.Thickness_Total",
    Int_Data."Thickness.Thickness_Ru-InP",
    Int_Data."Thickness.Thickness_Cap"
FROM(
    SELECT
        HDR.TESTHEADERID,
        HDR.STARTTIME,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        HDR.BATCHNUMBER,
	HDR.OPERATION,
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
        AND HDR.OPERATION               LIKE 'Ru-EML_F3%'
    GROUP BY
        HDR.TESTHEADERID,
        STARTTIME,
        PARTNUMBER,
        SERIALNUMBER,
        OPERATORID,
        BATCHNUMBER,
        OPERATION,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE
)
PIVOT(
    MAX(VALUE) FOR TESTSTEPS IN(
        'Particles_Particles_All_Address' AS "Particles.Particles_All_Address"
        ,'Particles_Particles_Center' AS "Particles.Particles_Center"
        ,'Particles_Particles_NG_Count' AS "Particles.Particles_NG_Count"
        ,'Particles_Particles' AS "Particles.Particles"
        ,'BallastN2_BallastN2' AS "BallastN2.BallastN2"
        ,'Thickness_Thickness_Total' AS "Thickness.Thickness_Total"
        ,'Thickness_Thickness_Ru-InP' AS "Thickness.Thickness_Ru-InP"
        ,'Thickness_Thickness_Cap' AS "Thickness.Thickness_Cap"
        ,'SORTED_DATA_STARTTIME_SORTED' AS "SORTED_DATA_STARTTIME_SORTED"
        ,'SORTED_DATA_SORTNUMBER' AS "SORTED_DATA_SORTNUMBER"

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
        AND HDR.OPERATION               LIKE 'Ru-EML_%'
    GROUP BY
        HDR.TESTHEADERID,
        HDR.SERIALNUMBER
)String_Data
WHERE 1=1
  AND Int_Data.TESTHEADERID = String_Data.TESTHEADERID
ORDER BY
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
