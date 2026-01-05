SELECT * FROM (
    SELECT
      HDR.STARTTIME,
      HDR.OPERATORID,
      HDR.PARTNUMBER,
      SUBSTR(HDR.SERIALNUMBER,1,5) AS SERIALNUMBER,
      SUBSTR(HDR.SERIALNUMBER,7,9) AS Banchi,
      HDR.SERIALNUMBER AS SERIALNUMBER_ADD_BANCHI,
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
	    AND HDR.STARTTIME		> '2010-01-01'
        AND HDR.OPERATION IN (
            'SEM-DML_Wmi_Width',
            'SEM-DML_ISO_Depth',
            'SEM-DML_MESA_Height',
            'SEM-DML_MESA_Width',
            'SEM-DML_SiO2_EtchingDepth',
            'SEM-DML_SiO2_Height',
            'SEM-DML_SiO2_Thickness'
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
        'MESA.Width_1' AS "MESA.Width_1",
        'MESA.Width_2' AS "MESA.Width_2",
        'MESA.Width_3' AS "MESA.Width_3",
        'MESA.Width_Ave' AS "MESA.Width_Ave",
        'MESA.Height_1' AS "MESA.Height_1",
        'MESA.Height_2' AS "MESA.Height_2",
        'MESA.Height_3' AS "MESA.Height_3",
        'MESA.Height_Ave' AS "MESA.Height_Ave",
        'SiO2.Height_1' AS "SiO2.Height_1",
        'SiO2.Height_2' AS "SiO2.Height_2",
        'SiO2.Height_3' AS "SiO2.Height_3",
        'SiO2.Height_Ave' AS "SiO2.Height_Ave",
        'SiO2.Thickness_1' AS "SiO2.Thickness_1",
        'SiO2.Thickness_2' AS "SiO2.Thickness_2",
        'SiO2.Thickness_3' AS "SiO2.Thickness_3",
        'SiO2.Thickness_Ave' AS "SiO2.Thickness_Ave", 
        'Wmi.Width_1' AS "Wmi.Width_1",
        'Wmi.Width_2' AS "Wmi.Width_2",
        'Wmi.Width_3' AS "Wmi.Width_3",
        'Wmi.Width_Ave' AS "Wmi.Width_Ave",
        'ISO.Depth_1' AS "ISO.Depth_1",
        'ISO.Depth_2' AS "ISO.Depth_2",
        'ISO.Depth_3' AS "ISO.Depth_3",
        'ISO.Depth_Ave' AS "ISO.Depth_Ave",
        'SiO2.EtchingDepth_1' AS "SiO2.EtchingDepth_1",
        'SiO2.EtchingDepth_2' AS "SiO2.EtchingDepth_2",
        'SiO2.EtchingDepth_3' AS "SiO2.EtchingDepth_3",
        'SiO2.EtchingDepth_Ave' AS "SiO2.EtchingDepth_Ave"
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
            AND HDR.OPERATION IN (
                'SEM-DML_Wmi_Width',
                'SEM-DML_ISO_Depth',
                'SEM-DML_MESA_Height',
                'SEM-DML_MESA_Width',
                'SEM-DML_SiO2_EtchingDepth',
                'SEM-DML_SiO2_Height',
                'SEM-DML_SiO2_Thickness'
            )
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER_ADD_BANCHI = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    STARTTIME
