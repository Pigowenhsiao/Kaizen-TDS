SELECT * FROM (
    SELECT
      HDR.STARTTIME,
      HDR.OPERATORID,
      HDR.PARTNUMBER,
      MAX(CASE DEVICENAME WHEN 'Nanospec' THEN DEVICENAME || DEVICESERIALNUMBER END) AS "Nanospec",
      MAX(CASE DEVICENAME WHEN 'DryEtch' THEN DEVICENAME || DEVICESERIALNUMBER END) AS "DryEtch",
      HDR.SERIALNUMBER,
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
        AND HDR.OPERATIONID             = OPS.OPERATIONID
        AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
        AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
        AND MSP.PARAMETERID             = PAR.PARAMETERID
        AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'MESA_CVD_EtchingRate'
    GROUP BY
        STARTTIME,
        OPERATORID,
        PARTNUMBER,
        SERIALNUMBER,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE

    )
PIVOT (
    MAX(VALUE) FOR STEP_AND_PARAMETER IN (
        'Thickness1.Initial1' AS "Thickness1.Initial1",
        'Thickness1.Initial2' AS "Thickness1.Initial2",
        'Thickness1.Initial3' AS "Thickness1.Initial3",
        'Thickness1.Initial4' AS "Thickness1.Initial4",
        'Thickness1.Initial5' AS "Thickness1.Initial5",
        'Thickness1.Initial_Ave' AS "Thickness1.Initial_Ave",
        'Thickness2.Final1' AS "Thickness2.Final1",
        'Thickness2.Final2' AS "Thickness2.Final2",
        'Thickness2.Final3' AS "Thickness2.Final3",
        'Thickness2.Final4' AS "Thickness2.Final4",
        'Thickness2.Final5' AS "Thickness2.Final5",
        'Thickness2.Final_Ave' AS "Thickness2.Final_Ave",
        'Rate.Rate1' AS "Rate.Rate1",
        'Rate.Rate2' AS "Rate.Rate2",
        'Rate.Rate3' AS "Rate.Rate3",
        'Rate.Rate4' AS "Rate.Rate4",
        'Rate.Rate5' AS "Rate.Rate5",
        'Rate.Rate_Ave' AS "Rate.Rate_Ave",
        'Rate.Rate_3sigma' AS "Rate.Rate_3sigma",
        'Time.Time' AS "Time.Time",
        'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
        'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER",
        'Coordinate.X1'     AS "X1",
        'Coordinate.X2'     AS "X2",
        'Coordinate.X3'     AS "X3",
        'Coordinate.X4'     AS "X4",
        'Coordinate.X5'     AS "X5",
        'Coordinate.Y1'     AS "Y1",
        'Coordinate.Y2'     AS "Y2",
        'Coordinate.Y3'     AS "Y3",
        'Coordinate.Y4'     AS "Y4",
        'Coordinate.Y5'     AS "Y5"
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
            AND HDR.OPERATION               = 'MESA_CVD_EtchingRate'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
