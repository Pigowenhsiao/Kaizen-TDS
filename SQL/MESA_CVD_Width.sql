SELECT * 
FROM(
    SELECT
        TST.STARTTIME,
        TST.OPERATORID,
        TST."EQUIPMENT",
        TST.PARTNUMBER,
        TST.SERIALNUMBER,
        TST."Width1.Width",
        TST."Width2.Width",
        TST."Width3.Width",
        TST."Width4.Width",
        TST."Width5.Width",
        TST."SORTED_DATA.STARTTIME_SORTED",
        TST."SORTED_DATA.SORTNUMBER",
        AAA."Width_Average",
        AAA."Width_3sigma",
        TST.RESULT,
        TST.X1,
        TST.X2,
        TST.X3,
        TST.X4,
        TST.X5,
        TST.Y1,
        TST.Y2,
        TST.Y3,
        TST.Y4,
        TST.Y5
        FROM(
            SELECT * FROM(
                SELECT
                    HDR.TESTHEADERID,
                    HDR.STARTTIME,
                    HDR.OPERATORID,
                    TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
                    HDR.PARTNUMBER,
                    HDR.SERIALNUMBER,
                    OPS.OPERATIONSTEPNAME || '.' || PAR.PARAMETERNAME AS VALUENAME,
                    MSD.VALUE,
                    HDR.RESULT
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
                    AND HDR.OPERATION               = 'MESA_CVD_Width'
            )
            PIVOT(
                MAX(VALUE) FOR VALUENAME IN(
                    'Width1.Width' AS "Width1.Width",
                    'Width2.Width' AS "Width2.Width",
                    'Width3.Width' AS "Width3.Width",
                    'Width4.Width' AS "Width4.Width",
                    'Width5.Width' AS "Width5.Width",
                    'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
                    'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER",
                    'Width1.X'     AS "X1",
                    'Width2.X'     AS "X2",
                    'Width3.X'     AS "X3",
                    'Width4.X'     AS "X4",
                    'Width5.X'     AS "X5",
                    'Width1.Y'     AS "Y1",
                    'Width2.Y'     AS "Y2",
                    'Width3.Y'     AS "Y3",
                    'Width4.Y'     AS "Y4",
                    'Width5.Y'     AS "Y5"
                )
            )
        )TST
    ,(
        SELECT
            HDR.TESTHEADERID,
            HDR.SERIALNUMBER,
            AVG(MSD.VALUE) AS "Width_Average",
            3*STDDEV(MSD.VALUE) AS "Width_3sigma"
        FROM
            TDSMFG.TESTHEADER           HDR,
            TDSMFG.TESTHEADERSTEP       HDS,
            TDSMFG.OPERATIONSTEP        OPS,
            TDSMFG.MEASUREMENTPARAMETER MSP,
            TDSMFG.MEASUREMENT          MSD,
            TDSMFG.PARAMETER            PAR
        WHERE 1=1
            AND HDR.TESTHEADERID            = HDS.TESTHEADERID
            AND HDR.OPERATIONID             = OPS.OPERATIONID
            AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
            AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
            AND MSP.PARAMETERID             = PAR.PARAMETERID
            AND HDR.OPERATION               = 'MESA_CVD_Width'
            AND PAR.PARAMETERNAME           = 'Width'
        GROUP BY
            HDR.TESTHEADERID,
            SERIALNUMBER
    )AAA    
    WHERE 1=1
        AND TST.TESTHEADERID = AAA.TESTHEADERID
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
            AND HDR.OPERATION               = 'MESA_CVD_Width'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
