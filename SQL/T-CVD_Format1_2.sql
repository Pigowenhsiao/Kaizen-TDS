SELECT * FROM(
    SELECT
        STARTTIME,
        HDR.OPERATION,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        MAX(CASE MISCDESC WHEN 'Group' THEN MISCVALUE END) AS "Group",
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
        AND TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER    = 'Nanospec1'
        AND HDR.OPERATION                               LIKE '%T-CVD%'
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
            'DepotTime.PSG' AS "DepotTime.PSG",
            'DepotTime.SiO2' AS "DepotTime.SiO2",
            'Temperature.Ts' AS "Temperature.Ts",
            'Temperature.T' AS "Temperature.T",
            'Temperature.T-Ts' AS "Temperature.T-Ts",
            'Temperature_Plus14.Ts_Plus14' AS "Temperature_Plus14.Ts_Plus14",
            'Temperature_Plus14.T_Plus14' AS "Temperature_Plus14.T_Plus14",
            'Temperature_Plus14.T-Ts_Plus14' AS "Temperature_Plus14.T-Ts_Plus14",
            'ReplacementTime.ReplacementTime' AS "ReplacementTime.ReplacementTime",
            'GasFlowRate.O2' AS "GasFlowRate.O2",
            'GasFlowRate.N2' AS "GasFlowRate.N2",
            'GasFlowRate.PH3' AS "GasFlowRate.PH3",
            'GasFlowRate.SiH4' AS "GasFlowRate.SiH4",
            'GasFlowRate.N2-carrier' AS "GasFlowRate.N2-carrier",
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
            TDSMFG.STRINGMEASUREMENT    SMM
        WHERE 1=1
            AND HDR.TESTHEADERID            = HDS.TESTHEADERID
            AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND HDS.TESTHEADERSTEPID        = STP.TESTHEADERSTEPID
            AND STP.STRINGPARAMETERID       = SMM.STRINGPARAMETERID
            AND PAR.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
            AND PAR.PARAMETERID             = STP.PARAMETERID
            AND HDR.OPERATION               LIKE '%T-CVD%'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    --Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
