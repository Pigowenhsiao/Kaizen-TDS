SELECT * FROM (
    SELECT
      HDR.STARTTIME,
      HDR.OPERATORID,
      HDR.PARTNUMBER,
      SUBSTR(HDR.SERIALNUMBER,1,5) AS SERIALNUMBER,
      TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
      OPS.OPERATIONSTEPNAME || '.' || PAR.PARAMETERNAME AS STEP_AND_PARAMETER,
      MSD.VALUE
    FROM
        TDSMFG.TESTHEADER           HDR,
        TDSMFG.TESTHEADERSTEP       HDS,
        TDSMFG.OPERATIONSTEP        OPS,
        TDSMFG.MEASUREMENTPARAMETER MSP,
        TDSMFG.MEASUREMENT          MSD,
        TDSMFG.PARAMETER            PAR,
        TDSMFG.TESTEQUIPMENT        TEQ
    WHERE 1=1
        AND HDR.TESTHEADERID            = HDS.TESTHEADERID
        AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
        AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
        AND MSP.PARAMETERID             = PAR.PARAMETERID
        AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND HDR.OPERATION IN (
            'ICP1_MESA_APCPress',
            'ICP1_MESA_APCPosition',
            'ICP1_MESA_PFC1Press',
            'ICP1_MESA_PFC1Flow',
            'ICP1_MESA_ArFlow',
            'ICP1_MESA_SiCl4Flow',
            'ICP1_MESA_AntennaPf',
            'ICP1_MESA_AntennaPr',
            'ICP1_MESA_BiasPf',
            'ICP1_MESA_BiasPr',
            'ICP1_MESA_BiasVpp',
            'ICP1_MESA_Discharge',
            'ICP1_MESA_ChamberBakeHeat',
            'ICP1_MESA_HotPlateHeat'
        )
    GROUP BY
        STARTTIME,
        OPERATORID,
        PARTNUMBER,
        SERIALNUMBER,
        DEVICENAME,
        DEVICESERIALNUMBER,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE
)
PIVOT(
    MAX(VALUE) FOR STEP_AND_PARAMETER IN (
        'APCPress.APCPress' AS "APCPress.APCPress",
        'APCPosition.APCPosition' AS "APCPosition.APCPosition",
        'PFC1Press.PFC1Press' AS "PFC1Press.PFC1Press",
        'PFC1Flow.PFC1Flow' AS "PFC1Flow.PFC1Flow",
        'ArFlow.ArFlow' AS "ArFlow.ArFlow",
        'SiCl4Flow.SiCl4Flow' AS "SiCl4Flow.SiCl4Flow",
        'AntennaPf.AntennaPf' AS "AntennaPf.AntennaPf",
        'AntennaPr.AntennaPr' AS "AntennaPr.AntennaPr",
        'BiasPf.BiasPf' AS "BiasPf.BiasPf",
        'BiasPr.BiasPr' AS "BiasPr.BiasPr",
        'BiasVpp.BiasVpp' AS "BiasVpp.BiasVpp",
        'Discharge.Discharge' AS "Discharge.Discharge",
        'ChamberBakeHeat.ChamberBakeHeat' AS "ChamberBakeHeat.ChamberBakeHeat",
        'HotPlateHeat.HotPlateHeat' AS "HotPlateHeat.HotPlateHeat"
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
            AND HDR.OPERATION IN(
            'ICP1_MESA_APCPress',
            'ICP1_MESA_APCPosition',
            'ICP1_MESA_PFC1Press',
            'ICP1_MESA_PFC1Flow',
            'ICP1_MESA_ArFlow',
            'ICP1_MESA_SiCl4Flow',
            'ICP1_MESA_AntennaPf',
            'ICP1_MESA_AntennaPr',
            'ICP1_MESA_BiasPf',
            'ICP1_MESA_BiasPr',
            'ICP1_MESA_BiasVpp',
            'ICP1_MESA_Discharge',
            'ICP1_MESA_ChamberBakeHeat',
            'ICP1_MESA_HotPlateHeat'
        )
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    STARTTIME
