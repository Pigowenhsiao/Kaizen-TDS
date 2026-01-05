SELECT * FROM(
    SELECT
        PIX1."PIX1_STARTTIME",
        PIX1."PIX1_Operator",
        PIX1."SORTED_DATA.STARTTIME_SORTED",
        PIX1."SORTED_DATA.SORTNUMBER",
        PIX1.PARTNUMBER,
        PIX1.SERIALNUMBER,
        PIX1."EQUIPMENT",
        PIX1."PIX1.Step1",
        PIX1."PIX1.Step2",
        PIX1."PIX1.Step3",
        PIX1."PIX1.Step_Ave",
        PIX1."PIX1.Step_3sigma",
        PIX2."PIX2_STARTTIME",
        PIX2."PIX2_Operator",
        PIX2."PIX2.Step1",
        PIX2."PIX2.Step2",
        PIX2."PIX2.Step3",
        PIX2."PIX2.Step_Ave",
        PIX2."PIX2.Step_3sigma",
        PIX3."PIX3_STARTTIME",
        PIX3."PIX3_Operator",
        PIX3."PIX3.Step1",
        PIX3."PIX3.Step2",
        PIX3."PIX3.Step3",
        PIX3."PIX3.Step_Ave",
        PIX3."PIX3.Step_3sigma"
        FROM(
                SELECT 
                    PIX1_TST.TESTHEADERID,
                    PIX1_TST."PIX1_STARTTIME",
                    PIX1_TST."SORTED_DATA.STARTTIME_SORTED",
                    PIX1_TST."SORTED_DATA.SORTNUMBER",
                    PIX1_TST."PIX1_Operator",
                    PIX1_TST.PARTNUMBER,
                    PIX1_TST.SERIALNUMBER,
                    PIX1_TST."EQUIPMENT",
                    PIX1_TST."PIX1.Step1",
                    PIX1_TST."PIX1.Step2",
                    PIX1_TST."PIX1.Step3",
                    PIX1_TST."PIX1.Step_Ave",
                    PIX1_TST."PIX1.Step_3sigma"
                 FROM (
                    SELECT * FROM(
                       SELECT
                            HDR.TESTHEADERID,
                            MAX(CASE THM.MISCDESC WHEN  'PIX1_Operator' THEN HDS.STARTDATETIME END) AS "PIX1_STARTTIME",
                            MAX(CASE THM.MISCDESC WHEN  'PIX1_Operator' THEN THM.MISCVALUE END) AS "PIX1_Operator",
                            HDR.PARTNUMBER,
                            HDR.SERIALNUMBER,
                            TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
                            OPS.OPERATIONSTEPNAME || '.' || PAR.PARAMETERNAME STEP_AND_PARAMETER,
                            MSD.VALUE
                        FROM
                            TDSMFG.TESTHEADER           HDR,
                            TDSMFG.TESTHEADERMISC        THM,
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
                            AND THM.MISCDESC                = 'PIX1_Operator'
                            AND HDR.OPERATION               = 'PIX_SiN_Step'
                        GROUP BY
                          HDR.TESTHEADERID,
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
                        'PIX1.Step1' AS "PIX1.Step1",
                        'PIX1.Step2' AS "PIX1.Step2",
                        'PIX1.Step3' AS "PIX1.Step3",
                        'PIX1.Step_Ave' AS "PIX1.Step_Ave",
                        'PIX1.Step_3sigma' AS "PIX1.Step_3sigma",
                        'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
                        'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER"
                    )
                )
                )PIX1_TST
            )PIX1
           ,(
            SELECT 
                    PIX2_TST.TESTHEADERID,
                    PIX2_TST."PIX2_STARTTIME",
                    PIX2_TST."PIX2_Operator",
                    PIX2_TST.PARTNUMBER,
                    PIX2_TST.SERIALNUMBER,
                    PIX2_TST."EQUIPMENT",
                    PIX2_TST."PIX2.Step1",
                    PIX2_TST."PIX2.Step2",
                    PIX2_TST."PIX2.Step3",
                    PIX2_TST."PIX2.Step_Ave",
                    PIX2_TST."PIX2.Step_3sigma"
                 FROM (
                    SELECT * FROM(
                       SELECT
                          HDR.TESTHEADERID,
                          MAX(CASE THM.MISCDESC WHEN  'PIX2_Operator' THEN HDS.STARTDATETIME END) AS "PIX2_STARTTIME",
                          MAX(CASE THM.MISCDESC WHEN  'PIX2_Operator' THEN THM.MISCVALUE END) AS "PIX2_Operator",
                          HDR.PARTNUMBER,
                          HDR.SERIALNUMBER,
                          TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
                          OPS.OPERATIONSTEPNAME || '.' || PAR.PARAMETERNAME AS STEP_AND_PARAMETER,
                          MSD.VALUE
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
                            AND THM.MISCDESC                = 'PIX2_Operator'
                            AND HDR.OPERATION               = 'PIX_SiN_Step'
                        GROUP BY
                          HDR.TESTHEADERID,
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
                        'PIX2.Step1' AS "PIX2.Step1",
                        'PIX2.Step2' AS "PIX2.Step2",
                        'PIX2.Step3' AS "PIX2.Step3",
                        'PIX2.Step_Ave' AS "PIX2.Step_Ave",
                        'PIX2.Step_3sigma' AS "PIX2.Step_3sigma"
                    )
                )
                )PIX2_TST
            )PIX2
           ,(
              SELECT
                    PIX3_TST.TESTHEADERID,
                    PIX3_TST."PIX3_STARTTIME",
                    PIX3_TST."PIX3_Operator",
                    PIX3_TST.PARTNUMBER,
                    PIX3_TST.SERIALNUMBER,
                    PIX3_TST."EQUIPMENT",
                    PIX3_TST."PIX3.Step1",
                    PIX3_TST."PIX3.Step2",
                    PIX3_TST."PIX3.Step3",
                    PIX3_TST."PIX3.Step_Ave",
                    PIX3_TST."PIX3.Step_3sigma"
                 FROM (
                    SELECT * FROM(
                       SELECT
                          HDR.TESTHEADERID,
                          MAX(CASE THM.MISCDESC WHEN  'PIX3_Operator' THEN HDS.STARTDATETIME END) AS "PIX3_STARTTIME",
                          MAX(CASE THM.MISCDESC WHEN  'PIX3_Operator' THEN THM.MISCVALUE END) AS "PIX3_Operator",
                          HDR.PARTNUMBER,
                          HDR.SERIALNUMBER,
                          TEQ.DEVICESERIALNUMBER AS "EQUIPMENT",
                          OPS.OPERATIONSTEPNAME || '.' || PAR.PARAMETERNAME AS STEP_AND_PARAMETER,
                          MSD.VALUE
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
                            AND THM.MISCDESC                = 'PIX3_Operator'
                            AND HDR.OPERATION               = 'PIX_SiN_Step'
                        GROUP BY
                          HDR.TESTHEADERID,
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
                        'PIX3.Step1' AS "PIX3.Step1",
                        'PIX3.Step2' AS "PIX3.Step2",
                        'PIX3.Step3' AS "PIX3.Step3",
                        'PIX3.Step_Ave' AS "PIX3.Step_Ave",
                        'PIX3.Step_3sigma' AS "PIX3.Step_3sigma"
                    )
                )
                )PIX3_TST
            )PIX3
    WHERE 1=1
        AND PIX1.TESTHEADERID = PIX2.TESTHEADERID
        AND PIX2.TESTHEADERID = PIX3.TESTHEADERID
        AND PIX1."PIX1.Step1" IS NOT NULL
        AND PIX2."PIX2.Step1" IS NOT NULL
        AND PIX3."PIX3.Step1" IS NOT NULL
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
            AND HDR.OPERATION               = 'PIX_SiN_Step'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    PIX1_STARTTIME
