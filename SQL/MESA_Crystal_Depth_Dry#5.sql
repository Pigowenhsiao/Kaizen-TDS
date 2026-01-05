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
        AND TEQ.DEVICESERIALNUMBER      = '#5'
)
PIVOT(
    MAX(VALUE) FOR STEP_AND_PARAMETER IN (
        'Order.Order' AS "Order.Order",
        'Depth1.First1' AS "Depth1.First1",
        'Depth1.First2' AS "Depth1.First2",
        'Depth1.First3' AS "Depth1.First3",
        'Depth1.First4' AS "Depth1.First4",
        'Depth1.First5' AS "Depth1.First5",
        'Depth1.First_Ave' AS "Depth1.First_Ave",
        'Thickness1.First1' AS "Thickness1.First1",
        'Thickness1.First2' AS "Thickness1.First2",
        'Thickness1.First3' AS "Thickness1.First3",
        'Thickness1.First4' AS "Thickness1.First4",
        'Thickness1.First5' AS "Thickness1.First5",
        'Thickness1.First_Ave' AS "Thickness1.First_Ave",
        'Etching1.First1' AS "Etching1.First1",
        'Etching1.First2' AS "Etching1.First2",
        'Etching1.First3' AS "Etching1.First3",
        'Etching1.First4' AS "Etching1.First4",
        'Etching1.First5' AS "Etching1.First5",
        'Etching1.First_Ave' AS "Etching1.First_Ave",
        'Etching1.First_3sigma' AS "Etching1.First_3sigma",
        'Etching1.First_Rate' AS "Etching1.First_Rate",
        'Additional1.Time' AS "Additional1.Time",
        'Depth2.Second1' AS "Depth2.Second1",
        'Depth2.Second2' AS "Depth2.Second2",
        'Depth2.Second3' AS "Depth2.Second3",
        'Depth2.Second4' AS "Depth2.Second4",
        'Depth2.Second5' AS "Depth2.Second5",
        'Depth2.Second_Ave' AS "Depth2.Second_Ave",
        'Thickness2.Second1' AS "Thickness2.Second1",
        'Thickness2.Second2' AS "Thickness2.Second2",
        'Thickness2.Second3' AS "Thickness2.Second3",
        'Thickness2.Second4' AS "Thickness2.Second4",
        'Thickness2.Second5' AS "Thickness2.Second5",
        'Thickness2.Second_Ave' AS "Thickness2.Second_Ave",
        'Etching2.Second1' AS "Etching2.Second1",
        'Etching2.Second2' AS "Etching2.Second2",
        'Etching2.Second3' AS "Etching2.Second3",
        'Etching2.Second4' AS "Etching2.Second4",
        'Etching2.Second5' AS "Etching2.Second5",
        'Etching2.Second_Ave' AS "Etching2.Second_Ave",
        'Etching2.Second_3sigma' AS "Etching2.Second_3sigma",
        'Etching2.Second_Depth' AS "Etching2.Second_Depth",
        'Etching2.Second_Rate' AS "Etching2.Second_Rate",
        'Etching2.Second_Error' AS "Etching2.Second_Error",
        'Final.Final1' AS "Final.Final1",
        'Final.Final2' AS "Final.Final2",
        'Final.Final3' AS "Final.Final3",
        'Final.Final4' AS "Final.Final4",
        'Final.Final5' AS "Final.Final5",
        'Final.Final_Ave' AS "Final.Final_Ave",
        'Final.Final_3sigma' AS "Final.Final_3sigma",
        'Final.Final_Max-Min' AS "Final.Final_Max-Min",
        'Final.Final_Time1' AS "Final.Final_Time1",
        'Final.Final_Time2' AS "Final.Final_Time2",
        'SiO2.Etching' AS "SiO2.Etching",
        'SiO2.Rate' AS "SiO2.Rate",
        'SiO2.Selectivity' AS "SiO2.Selectivity",
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
            AND TEQ.DEVICESERIALNUMBER      = '#5'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
