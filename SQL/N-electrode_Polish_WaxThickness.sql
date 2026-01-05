SELECT * FROM(
    SELECT
        TST.STARTTIME,
        TST.OPERATORID,
        TST.EQUIPMENT,
        TST.PARTNUMBER,
        TST.SERIALNUMBER,
        TST."Thickness1.Thickness",
        TST."Thickness2.Thickness",
        TST."Thickness3.Thickness",
        TST."Thickness4.Thickness",
        TST."Thickness5.Thickness",
        TST."Thickness6.Thickness",
        TST."Thickness7.Thickness",
        TST."Thickness8.Thickness",
        TST."Thickness9.Thickness",
        TST."Thickness10.Thickness",
        TST."Thickness11.Thickness",
        TST."Thickness12.Thickness",
        TST."Thickness13.Thickness",
        AAA."Thickness_Average",
        AAA."Thickness_Max",
        AAA."Thickness_Min",
        TST."Thickness1.X",
        TST."Thickness2.X",
        TST."Thickness3.X",
        TST."Thickness4.X",
        TST."Thickness5.X",
        TST."Thickness6.X",
        TST."Thickness7.X",
        TST."Thickness8.X",
        TST."Thickness9.X",
        TST."Thickness10.X",
        TST."Thickness11.X",
        TST."Thickness12.X",
        TST."Thickness13.X",
        TST."Thickness1.Y",
        TST."Thickness2.Y",
        TST."Thickness3.Y",
        TST."Thickness4.Y",
        TST."Thickness5.Y",
        TST."Thickness6.Y",
        TST."Thickness7.Y",
        TST."Thickness8.Y",
        TST."Thickness9.Y",
        TST."Thickness10.Y",
        TST."Thickness11.Y",
        TST."Thickness12.Y",
        TST."Thickness13.Y"
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
                    AND HDR.OPERATION               = 'N-electrode_Polish_WaxThickness'
            )
            PIVOT(
                MAX(VALUE) FOR VALUENAME IN (
                    'Thickness1.Thickness'  AS "Thickness1.Thickness",
                    'Thickness2.Thickness'  AS "Thickness2.Thickness",
                    'Thickness3.Thickness'  AS "Thickness3.Thickness",
                    'Thickness4.Thickness'  AS "Thickness4.Thickness",
                    'Thickness5.Thickness'  AS "Thickness5.Thickness",
                    'Thickness6.Thickness'  AS "Thickness6.Thickness",
                    'Thickness7.Thickness'  AS "Thickness7.Thickness",
                    'Thickness8.Thickness'  AS "Thickness8.Thickness",
                    'Thickness9.Thickness'  AS "Thickness9.Thickness",
                    'Thickness10.Thickness'  AS "Thickness10.Thickness",
                    'Thickness11.Thickness'  AS "Thickness11.Thickness",
                    'Thickness12.Thickness'  AS "Thickness12.Thickness",
                    'Thickness13.Thickness'  AS "Thickness13.Thickness",
                    'Thickness1.X'     AS "Thickness1.X",
                    'Thickness2.X'     AS "Thickness2.X",
                    'Thickness3.X'     AS "Thickness3.X",
                    'Thickness4.X'     AS "Thickness4.X",
                    'Thickness5.X'     AS "Thickness5.X",
                    'Thickness6.X'     AS "Thickness6.X",
                    'Thickness7.X'     AS "Thickness7.X",
                    'Thickness8.X'     AS "Thickness8.X",
                    'Thickness9.X'     AS "Thickness9.X",
                    'Thickness10.X'     AS "Thickness10.X",
                    'Thickness11.X'     AS "Thickness11.X",
                    'Thickness12.X'     AS "Thickness12.X",
                    'Thickness13.X'     AS "Thickness13.X",
                    'Thickness1.Y'     AS "Thickness1.Y",
                    'Thickness2.Y'     AS "Thickness2.Y",
                    'Thickness3.Y'     AS "Thickness3.Y",
                    'Thickness4.Y'     AS "Thickness4.Y",
                    'Thickness5.Y'     AS "Thickness5.Y",
                    'Thickness6.Y'     AS "Thickness6.Y",
                    'Thickness7.Y'     AS "Thickness7.Y",
                    'Thickness8.Y'     AS "Thickness8.Y",
                    'Thickness9.Y'     AS "Thickness9.Y",
                    'Thickness10.Y'     AS "Thickness10.Y",
                    'Thickness11.Y'     AS "Thickness11.Y",
                    'Thickness12.Y'     AS "Thickness12.Y",
                    'Thickness13.Y'     AS "Thickness13.Y"
                )
            )
        )TST
        ,(
            SELECT
                HDR.TESTHEADERID,
                HDR.SERIALNUMBER,
                AVG(MSD.VALUE) AS "Thickness_Average",
                MAX(MSD.VALUE) AS "Thickness_Max",
                MIN(MSD.VALUE) AS "Thickness_Min"
            FROM
                TDSMFG.TESTHEADER           HDR,
                TDSMFG.TESTHEADERSTEP       HDS,
                TDSMFG.OPERATIONSTEP        OPS,
                TDSMFG.MEASUREMENTPARAMETER MSP,
                TDSMFG.MEASUREMENT          MSD,
                TDSMFG.PARAMETER            PAR
            WHERE 1=1
                AND HDR.TESTHEADERID            = HDS.TESTHEADERID
                AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
                AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
                AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
                AND MSP.PARAMETERID             = PAR.PARAMETERID
                AND HDR.OPERATION               = 'N-electrode_Polish_WaxThickness'
                AND PAR.PARAMETERNAME           = 'Thickness'
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
            AND HDR.OPERATION               IN 
                (   
                    'N-electrode_Polish_InitialWaferThickness',
                    'N-electrode_Polish_WaxThickness',
                    'N-electrode_Polish_MirrorPolishedThickness',
                    'N-electrode_Polish_RoughPolishedThickness',
                    'N-electrode_Polish_EtchedThickness'
                )
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    STARTTIME

