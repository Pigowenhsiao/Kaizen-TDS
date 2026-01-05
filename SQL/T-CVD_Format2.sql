SELECT * FROM(
    SELECT
        HDR.STARTTIME,
        HDR.OPERATION,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        HDR.BATCHNUMBER,
        MAX(CASE MISCDESC WHEN 'Group' THEN MISCVALUE END) AS "Group",
        MAX(CASE MISCDESC WHEN 'CVD-No' THEN MISCVALUE END) AS "CVD-No",
        MAX(CASE MISCDESC WHEN 'TypeOfDeposition' THEN MISCVALUE END) AS "TypeOfDeposition",
        MAX(CASE MISCDESC WHEN 'Recipe' THEN MISCVALUE END) AS "Recipe",
        MAX(CASE TEQ.DEVICENAME WHEN 'Nanospec' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "Nanospec",
        MAX(CASE TEQ.DEVICESERIALNUMBER WHEN '1' THEN 'T-CVD#1' ELSE 'T-CVD#2' END) AS "T-CVD",
        OPS.OPERATIONSTEPNAME || '.' || PARAMETERNAME AS TESTSTEPS,
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
        AND HDR.OPERATION               LIKE '%T-CVD%'
        AND TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER = 'Nanospec2'
    GROUP BY
        OPERATION,
        STARTTIME,
        PARTNUMBER,
        SERIALNUMBER,
        OPERATORID,
        BATCHNUMBER,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE
    )
    PIVOT(
        MAX(VALUE) FOR TESTSTEPS IN(
            'Coordinate.X' AS "X",
            'Coordinate.Y' AS "Y",
            'Thickness.Thickness' AS "Thickness.Thickness",
            'HEAD1.O2' AS "HEAD1.O2",
            'HEAD1.N2_O2' AS "HEAD1.N2_O2",
            'HEAD1.N2_SiH4' AS "HEAD1.N2_SiH4",
            'HEAD1.SiH4_4percent' AS "HEAD1.SiH4_4percent",
            'HEAD1.PH3_1percent' AS "HEAD1.PH3_1percent",
            'HEAD1.SiH4_100percent' AS "HEAD1.SiH4_100percent",
            'HEAD1.PH3_100percent' AS "HEAD1.PH3_100percent",
            'HEAD1.P_P+Si' AS "HEAD1.P_P+Si",
            'HEAD1.Ratio_O2/Gas' AS "HEAD1.Ratio_O2/Gas",
            'HEAD2.O2' AS "HEAD2.O2",
            'HEAD2.N2_SiH4' AS "HEAD2.N2_SiH4",
            'HEAD2.N2_O2' AS "HEAD2.N2_O2",
            'HEAD2.SiH4_4percent' AS "HEAD2.SiH4_4percent",
            'HEAD2.PH3_1percent' AS "HEAD2.PH3_1percent",
            'HEAD2.SiH4_100percent' AS "HEAD2.SiH4_100percent",
            'HEAD2.Ratio_O2' AS "HEAD2.Ratio_O2",
            'Temp.Target' AS "Temp.Target",
            'Temp.H1' AS "Temp.H1",
            'Temp.H2' AS "Temp.H2",
            'Temp.H3' AS "Temp.H3",
            'Temp.H4' AS "Temp.H4",
            'Temp.H5' AS "Temp.H5",
            'Speed.Speed' AS "Speed.Speed",
            'Position.Start' AS "Position.Start",
            'Position.Stop' AS "Position.Stop",
            'Position.Delta' AS "Position.Delta",
            'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
            'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER"
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
            TDSMFG.STRINGMEASUREMENT    SMM,
            TDSMFG.TESTEQUIPMENT        TEQ
        WHERE 1=1
            AND HDR.TESTHEADERID            = HDS.TESTHEADERID
            AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND HDS.TESTHEADERSTEPID        = STP.TESTHEADERSTEPID
            AND STP.STRINGPARAMETERID       = SMM.STRINGPARAMETERID
            AND PAR.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND PAR.PARAMETERID             = STP.PARAMETERID
            AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
            AND HDR.OPERATION               LIKE '%T-CVD%'
            AND TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER = 'Nanospec2'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
