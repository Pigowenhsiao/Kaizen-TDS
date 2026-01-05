SELECT * FROM(
    SELECT 
        HDR.STARTTIME,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        HDR.BATCHNUMBER,
        MAX(CASE MISCDESC WHEN 'RecipeName-Macro' THEN MISCVALUE END) AS "RecipeName_Macro",
        MAX(CASE MISCDESC WHEN 'RecipeName-Program' THEN MISCVALUE END) AS "RecipeName_Program",
        MAX(CASE MISCDESC WHEN 'RecipeName-Folder' THEN MISCVALUE END) AS "RecipeName_Folder",
        OPS.OPERATIONSTEPNAME || '_' || PARAMETERNAME AS TESTSTEPS,
        MSD.VALUE,
        MAX(CASE TEQ.DEVICENAME WHEN 'SEM' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "SEM",
        MAX(CASE MO.DEVICENAME  WHEN 'MOCVD' THEN MO.DEVICENAME || MO.DEVICESERIALNUMBER END) AS "MOCVD"
    FROM
        TDSMFG.TESTHEADER           HDR,
        TDSMFG.TESTHEADERMISC       THM,
        TDSMFG.TESTHEADERSTEP       HDS,
        TDSMFG.OPERATIONSTEP        OPS,
        TDSMFG.MEASUREMENTPARAMETER MSP,
        TDSMFG.MEASUREMENT          MSD,
        TDSMFG.PARAMETER            PAR,
        TDSMFG.TESTEQUIPMENT        TEQ,
        TDSMFG.TESTEQUIPMENT        MO
    WHERE 1=1
        AND HDR.TESTHEADERID            = HDS.TESTHEADERID
        AND HDR.TESTHEADERID            = THM.TESTHEADERID
        AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND HDS.TESTHEADERSTEPID        = MSP.TESTHEADERSTEPID
        AND MSP.MEASUREMENTPARAMETERID  = MSD.MEASUREMENTPARAMETERID
        AND MSP.PARAMETERID             = PAR.PARAMETERID
        AND TEQ.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND MO.TESTHEADERID(+)         = HDR.TESTHEADERID
        AND HDR.OPERATION          = 'WG-EML'
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD6'
    GROUP BY
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
        'Coordinate_X' AS "X",
        'Coordinate_Y' AS "Y",
        'Thickness_Thickness_Cap' AS "Thickness_Cap",
        'Thickness_Thickness_Core' AS "Thickness_Core",
        'Thickness_Thickness_Total' AS "Thickness_Total",
        'Dulation_Step9' AS "Dulation_Step9",
        'Dulation_Step13' AS "Dulation_Step13",
        'Dulation_Step17' AS "Dulation_Step17",
        'Dulation_Step21' AS "Dulation_Step21",
        'MO1-TMI_Step9' AS "MO1-TMI_Step9",
        'MO1-TMI_Step13' AS "MO1-TMI_Step13",
        'MO1-TMI_Step17' AS "MO1-TMI_Step17",
        'MO2-TEG_Step9' AS "MO2-TEG_Step9",
        'MO2-TEG_Step13' AS "MO2-TEG_Step13",
        'MO2-TEG_Step17' AS "MO2-TEG_Step17",
        'MO5-TMI-3_Step21' AS "MO5-TMI-3_Step21",
        'AsH3-21-5percent_Step9' AS "AsH3-5percent_Step9",
        'AsH3-21-5percent_Step13' AS "AsH3-5percent_Step13",
        'AsH3-21-5percent_Step17' AS "AsH3-5percent_Step17",
        'AsH3-21-5percent_Cylinder' AS "AsH3-5percent_Cylinder",
        'PH3-1-50percent_Step21' AS "PH3-1-50percent_Step21",
        'PH3-2-50percent_Step9' AS "PH3-2-50percent_Step9",
        'PH3-2-50percent_Step13' AS "PH3-2-50percent_Step13",
        'PH3-2-50percent_Step17' AS "PH3-2-50percent_Step17",
        'Temperature_Step9' AS "Temperature_Step9",
        'Temperature_Step13' AS "Temperature_Step13",
        'Temperature_Step17' AS "Temperature_Step17",
        'Temperature_Step21' AS "Temperature_Step21",
        'BallastN2_BallastN2' AS "BallastN2_BallastN2",
        'MO-Temperature_MO1-TMI' AS "MO-Temperature_MO1-TMI",
        'MO-Temperature_MO2-TEG' AS "MO-Temperature_MO2-TEG",
        'MO-Temperature_MO3-TMI' AS "MO-Temperature_MO3-TMI",
        'MO-Temperature_MO4-TEG' AS "MO-Temperature_MO4-TEG",
        'MO-Temperature_MO5-TMI' AS "MO-Temperature_MO5-TMI",
        'MO-Temperature_MO6-TEG' AS "MO-Temperature_MO6-TEG",
        'MO-Temperature_MO7-Mg' AS "MO-Temperature_MO7-Mg",
        'MO-Temperature_MO6-TMA1' AS "MO-Temperature_MO6-TMA1",
        'Strain_Straing_WG' AS "Straing_WG",
        'Wavelength_Wavelength_WG' AS "Wavelength_WG",
        'Wavelength_Wavelength_Intensity' AS "Wavelength_Intensity",
        'Wavelength_Wavelength_FWHM' AS "Wavelength_FWHM",
        'SORTED_DATA_STARTTIME_SORTED' AS "SORTED_DATA_STARTTIME_SORTED",
        'SORTED_DATA_SORTNUMBER' AS "SORTED_DATA_SORTNUMBER"
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
            AND HDR.OPERATION               = 'WG-EML'
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    Int_Data."SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
