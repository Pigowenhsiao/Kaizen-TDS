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
            'Ion-milling2_Electrode_AcceleratorA',
            'Ion-milling2_Electrode_AcceleratorV',
            'Ion-milling2_Electrode_ArFlow',
            'Ion-milling2_Electrode_BeamA',
            'Ion-milling2_Electrode_BeamV',
            'Ion-milling2_Electrode_CathodeFilamentA',
            'Ion-milling2_Electrode_ChamberPress',
            'Ion-milling2_Electrode_DischargeA',
            'Ion-milling2_Electrode_DischargeV',
            'Ion-milling2_Electrode_EPDPress',
            'Ion-milling2_Electrode_NeutralizerEmissionA',
            'Ion-milling2_Electrode_NeutralizerFilamentA',
            'Ion-milling2_Electrode_StageAngle',
            'Ion-milling2_Electrode_StageHeFlow',
            'Ion-milling2_Electrode_StageHePress',
            'Ion-milling2_Electrode_StageSpin',
            'Ion-milling2_Electrode_Time',
            'Ion-milling2_Electrode_WaterFlow'
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
        'AcceleratorA.AcceleratorA' AS "AcceleratorA.AcceleratorA",
        'AcceleratorV.AcceleratorV' AS "AcceleratorV.AcceleratorV",
        'ArFlow.ArFlow' AS "ArFlow.ArFlow",
        'BeamA.BeamA' AS "BeamA.BeamA",
        'BeamV.BeamV' AS "BeamV.BeamV",
        'CathodeFilamentA.CathodeFilamentA' AS "CathodeFilamentA.CathodeFilamentA",
        'ChamberPress.ChamberPress' AS "ChamberPress.ChamberPress",
        'DischargeA.DischargeA' AS "DischargeA.DischargeA",
        'DischargeV.DischargeV' AS "DischargeV.DischargeV",
        'EPDPress.EPDPress' AS "EPDPress.EPDPress",
        'NeutralizerEmissionA.NeutralizerEmissionA' AS "NeutralizerEmissionA.NeutralizerEmissionA",
        'NeutralizerFilamentA.NeutralizerFilamentA' AS "NeutralizerFilamentA.NeutralizerFilamentA",
        'StageAngle.StageAngle' AS "StageAngle.StageAngle",
        'StageHeFlow.StageHeFlow' AS "StageHeFlow.StageHeFlow",
        'StageHePress.StageHePress' AS "StageHePress.StageHePress",
        'StageSpin.StageSpin' AS "StageSpin.StageSpin",
        'Time.Time' AS "Time.Time",
        'WaterFlow.WaterFlow' AS "WaterFlow.WaterFlow"
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
            'Ion-milling2_Electrode_AcceleratorA',
            'Ion-milling2_Electrode_AcceleratorV',
            'Ion-milling2_Electrode_ArFlow',
            'Ion-milling2_Electrode_BeamA',
            'Ion-milling2_Electrode_BeamV',
            'Ion-milling2_Electrode_CathodeFilamentA',
            'Ion-milling2_Electrode_ChamberPress',
            'Ion-milling2_Electrode_DischargeA',
            'Ion-milling2_Electrode_DischargeV',
            'Ion-milling2_Electrode_EPDPress',
            'Ion-milling2_Electrode_NeutralizerEmissionA',
            'Ion-milling2_Electrode_NeutralizerFilamentA',
            'Ion-milling2_Electrode_StageAngle',
            'Ion-milling2_Electrode_StageHeFlow',
            'Ion-milling2_Electrode_StageHePress',
            'Ion-milling2_Electrode_StageSpin',
            'Ion-milling2_Electrode_Time',
            'Ion-milling2_Electrode_WaterFlow'
        )
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    STARTTIME
