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
            'Dry6_CHF3-10sccm_ValvePosition',
            'Dry6_CHF3-10sccm_CHF3Flow',
            'Dry6_CHF3-10sccm_O2Flow',
            'Dry6_CHF3-10sccm_RFPower',
            'Dry6_CHF3-10sccm_Pressure',
            'Dry6_CHF3-10sccm_InitialVpp',
            'Dry6_CHF3-10sccm_FinalVpp'
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
        'ValvePosition.ValvePosition' AS "ValvePosition.ValvePosition",
        'CHF3Flow.CHF3Flow' AS "CHF3Flow.CHF3Flow",
        'O2Flow.O2Flow' AS "O2Flow.O2Flow",
        'RFPower.RFPower' AS "RFPower.RFPower",
        'Pressure.Pressure' AS "Pressure.Pressure",
        'InitialVpp.InitialVpp' AS "InitialVpp.InitialVpp",
        'FinalVpp.FinalVpp' AS "FinalVpp.FinalVpp"
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
            'Dry6_CHF3-10sccm_ValvePosition',
            'Dry6_CHF3-10sccm_CHF3Flow',
            'Dry6_CHF3-10sccm_O2Flow',
            'Dry6_CHF3-10sccm_RFPower',
            'Dry6_CHF3-10sccm_Pressure',
            'Dry6_CHF3-10sccm_InitialVpp',
            'Dry6_CHF3-10sccm_FinalVpp'
        )
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    STARTTIME
