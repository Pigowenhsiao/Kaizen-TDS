SELECT * FROM (
    SELECT
      HDR.STARTTIME,
      HDR.OPERATORID,
      HDR.PARTNUMBER,
      HDR.SERIALNUMBER,
      TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
      MSD.OPERATIONSTEPNAME || '.' || MSD.PARAMETERNAME AS STEP_AND_PARAMETER,
      MSD.VALUE
    FROM TESTHEADER_V HDR 
        INNER JOIN MEASUREMENT_V MSD ON HDR.TESTHEADERID = MSD.TESTHEADERID
        LEFT JOIN TDSMFG.TESTEQUIPMENT TEQ ON TEQ.TESTHEADERID = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'GRATING_CVD_Duty'
)
PIVOT(
    MAX(VALUE) FOR STEP_AND_PARAMETER IN (
        'Length1.Pitch' AS "Length1.Pitch1",
        'Length2.Pitch' AS "Length2.Pitch2",
        'Length3.Pitch' AS "Length3.Pitch3",
        'Length4.Pitch' AS "Length4.Pitch4",
        'Length5.Pitch' AS "Length5.Pitch5",
        'Length1.Space' AS "Length1.Space1",
        'Length2.Space' AS "Length2.Space2",
        'Length3.Space' AS "Length3.Space3",
        'Length4.Space' AS "Length4.Space4",
        'Length5.Space' AS "Length5.Space5",
        'Length1.Duty' AS "Length1.Duty1",
        'Length2.Duty' AS "Length2.Duty2",
        'Length3.Duty' AS "Length3.Duty3",
        'Length4.Duty' AS "Length4.Duty4",
        'Length5.Duty' AS "Length5.Duty5",
        'Average.Duty' AS "Average.Duty",
        'Average.3sigma' AS "Average.Duty_3sigma",
        'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
        'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER",
        'Length1.X'     AS "X1",
        'Length2.X'     AS "X2",
        'Length3.X'     AS "X3",
        'Length4.X'     AS "X4",
        'Length5.X'     AS "X5",
        'Length1.Y'     AS "Y1",
        'Length2.Y'     AS "Y2",
        'Length3.Y'     AS "Y3",
        'Length4.Y'     AS "Y4",
        'Length5.Y'     AS "Y5"
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
            AND HDR.OPERATION               = 'GRATING_CVD_Duty'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    "SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
