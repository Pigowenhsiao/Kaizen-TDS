SELECT * FROM (
    SELECT
      HDR.STARTTIME,
      HDR.OPERATORID,
      HDR.PARTNUMBER,
      SUBSTR(HDR.SERIALNUMBER,1,5) AS SERIALNUMBER,
      HDR.SERIALNUMBER AS SERIALNUMBER_ADD_BANCHI,
      SUBSTR(HDR.SERIALNUMBER,7,9) AS Banchi,
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
            'SEM-EML_MESA_Width',
            'SEM-EML_Ru_Thickness',
            'SEM-EML_MESA_Height',
            'SEM-EML_Contact_Width',
            'SEM-EML_Contact_Degree'
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
        'Contact.Width_1' AS "Contact.Width_1",
        'Contact.Width_2' AS "Contact.Width_2",
        'Contact.Width_3' AS "Contact.Width_3",
        'Contact.Width_Ave' AS "Contact.Width_Ave",
        'Contact.Degree_1' AS "Contact.Degree_1",
        'Contact.Degree_2' AS "Contact.Degree_2",
        'Contact.Degree_3' AS "Contact.Degree_3",
        'Contact.Degree_Ave' AS "Contact.Degree_Ave",
        'Ru.Thickness_1' AS "Ru.Thickness_1",
        'Ru.Thickness_2' AS "Ru.Thickness_2",
        'Ru.Thickness_3' AS "Ru.Thickness_3",
        'Ru.Thickness_Ave' AS "Ru.Thickness_Ave"
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
            'SEM-EML_MESA_Width',
            'SEM-EML_Ru_Thickness',
            'SEM-EML_MESA_Height',
            'SEM-EML_Contact_Width',
            'SEM-EML_Contact_Degree'
        )
        GROUP BY
            HDR.SERIALNUMBER
    )String_Data
    ON Int_Data.SERIALNUMBER_ADD_BANCHI = String_Data."FIVE_SERIALNUMBER"
ORDER BY
    STARTTIME
