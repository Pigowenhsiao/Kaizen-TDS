SELECT * FROM (
    SELECT
        HDR.STARTTIME,
        HDR.OPERATORID,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
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
        AND HDR.OPERATION               = 'MESA_Crystal_Depth'
        AND TEQ.DEVICESERIALNUMBER      = 'ICP'
        
)
PIVOT(
    MAX(VALUE) FOR STEP_AND_PARAMETER IN (
        'Order.Order' AS "Order.Order",
        'Time.Time' AS "Time.Time",
        'Depth.First1' AS "Depth.First1",
        'Depth.First2' AS "Depth.First2",
        'Depth.First3' AS "Depth.First3",
        'Depth.First4' AS "Depth.First4",
        'Depth.First5' AS "Depth.First5",
        'Depth.First_Ave' AS "Depth.First_Ave",
        'Thickness.First1' AS "Thickness.First1",
        'Thickness.First2' AS "Thickness.First2",
        'Thickness.First3' AS "Thickness.First3",
        'Thickness.First4' AS "Thickness.First4",
        'Thickness.First5' AS "Thickness.First5",
        'Thickness.First_Ave' AS "Thickness.First_Ave",
        'Etching.Etching1' AS "Etching.Etching1",
        'Etching.Etching2' AS "Etching.Etching2",
        'Etching.Etching3' AS "Etching.Etching3",
        'Etching.Etching4' AS "Etching.Etching4",
        'Etching.Etching5' AS "Etching.Etching5",
        'Etching.Etching_Ave' AS "Etching.Etching_Ave",
        'Etching.Etching_Max-Min' AS "Etching.Etching_Max-Min",
        'Etching.Etching_3sigma' AS "Etching.Etching_3sigma",
        'Etching.Etching_Rate' AS "Etching.Etching_Rate",
        'Etching.Etching_Error' AS "Etching.Etching_Error",
        'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
        'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER",
        'Coordinate.X1' AS "X1",
        'Coordinate.X2' AS "X2",
        'Coordinate.X3' AS "X3",
        'Coordinate.X4' AS "X4",
        'Coordinate.X5' AS "X5",
        'Coordinate.Y1' AS "Y1",
        'Coordinate.Y2' AS "Y2",
        'Coordinate.Y3' AS "Y3",
        'Coordinate.Y4' AS "Y4",
        'Coordinate.Y5' AS "Y5" 
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
            TDSMFG.TESTEQUIPMENT        TEQ,
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
            AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
            AND HDR.OPERATION               = 'MESA_Crystal_Depth'
            AND TEQ.DEVICESERIALNUMBER      = 'ICP'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
