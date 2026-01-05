SELECT
    Int_Data.STARTTIME,
    Int_Data.PARTNUMBER,
    Int_Data.SERIALNUMBER,
    String_Data."NINE_SERIALNUMBER",
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    Int_Data."SORTED_DATA.SORTNUMBER",
    Int_Data."PL",
    Int_Data."XRD",
    String_Data."GrowthID",
    String_Data."SubstrateID",
    Int_Data."PL-Wavelength-Center",
    Int_Data."PL-Wavelength-Average",
    Int_Data."PL-Wavelength-Median",
    Int_Data."PL-Wavelength-Sigma",
    Int_Data."PL-Wavelength-Delta",
    Int_Data."PL-Intensity-Rate_HH-LH",
    Int_Data."PL-Wavelength-HH_LH",
    Int_Data."PL-Intensity",
    Int_Data."PL-FWHM",
    Int_Data."PL-Intensity_Rate",
    Int_Data."PL-Intensity_Rate_Calk",
    Int_Data."XRD-Strain",
    Int_Data."XRD-Thickness"
FROM(
    SELECT
        HDR.TESTHEADERID,
        HDR.STARTTIME,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        MAX(CASE TEQ.DEVICENAME WHEN 'PL' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "PL",
        MAX(CASE TEQ.DEVICENAME WHEN 'XRD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "XRD",
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
        AND HDR.OPERATION               = 'T1-DML'
    GROUP BY
        HDR.TESTHEADERID,
        STARTTIME,
        PARTNUMBER,
        SERIALNUMBER,
        OPERATIONSTEPNAME,
        PARAMETERNAME,
        VALUE
    )
    PIVOT(
        MAX(VALUE) FOR STEP_AND_PARAMETER IN (
            'PL.PL-Wavelength-Center' AS "PL-Wavelength-Center",
            'PL.PL-Wavelength-Average' AS "PL-Wavelength-Average",
            'PL.PL-Wavelength-Median' AS "PL-Wavelength-Median",
            'PL.PL-Wavelength-Sigma' AS "PL-Wavelength-Sigma",
            'PL.PL-Wavelength-Delta' AS "PL-Wavelength-Delta",
            'PL.PL-Intensity-Rate_HH-LH' AS "PL-Intensity-Rate_HH-LH",
            'PL.PL-Wavelength_HH-LH' AS "PL-Wavelength-HH_LH",
            'PL.PL-Intensity' AS "PL-Intensity",
            'PL.PL-FWHM' AS "PL-FWHM",
            'PL.PL-Intensity_Rate' AS "PL-Intensity_Rate",
            'PL.PL-Intensity_Rate_Calk' AS "PL-Intensity_Rate_Calk",
            'XRD.XRD-Strain' AS "XRD-Strain",
            'XRD.XRD-Thickness' AS "XRD-Thickness",
            'SORTED_DATA.STARTTIME_SORTED' AS "SORTED_DATA.STARTTIME_SORTED",
            'SORTED_DATA.SORTNUMBER' AS "SORTED_DATA.SORTNUMBER"

        )    
    )Int_Data,
    (
        SELECT
            HDR.TESTHEADERID,
            MAX(CASE PARAMETERNAME WHEN 'GrowthID' THEN VALUESTRING END) AS "GrowthID",
            MAX(CASE PARAMETERNAME WHEN 'SubstrateID' THEN VALUESTRING END) AS "SubstrateID",
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
            AND HDR.OPERATION               = 'T1-DML'
        GROUP BY
            HDR.TESTHEADERID
    )String_Data
WHERE
    1 = 1
    AND Int_Data.TESTHEADERID = String_Data.TESTHEADERID
ORDER BY
    Int_Data."SORTED_DATA.STARTTIME_SORTED",
    STARTTIME
