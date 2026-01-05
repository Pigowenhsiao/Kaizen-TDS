SELECT
    Int_Data.STARTTIME,
    String_Data."NINE_SERIALNUMBER",
    Int_Data.PARTNUMBER,
    Int_Data."STARTTIME_SORTED",
    Int_Data."SORTNUMBER",
    Int_Data.OPERATORID,
    Int_Data.BATCHNUMBER,
    String_Data."SIX_BATCHNUMBER",
    Int_Data."SEM",
    Int_Data."PLmapper",
    Int_Data."XRD",
    Int_Data."MOCVD",
    Int_Data."RecipeName_Macro",
    Int_Data."RecipeName_Program",
    Int_Data."RecipeName_Folder",
    Int_Data."X",
    Int_Data."Y",
    Int_Data."Thickness_Cap",
    Int_Data."Thickness_Core",
    Int_Data."Thickness_Total",
    Int_Data."Straing_WG",
    Int_Data."Wavelength_WG",
    Int_Data."Wavelength_Intensity",
    Int_Data."Wavelength_FWHM",
    Int_Data."EpiTT_WG",
    Int_Data."Dulation_Step4",
    Int_Data."Dulation_Step6",
    Int_Data."Dulation_Step8",
    Int_Data."Dulation_Step10",
    Int_Data."TMGa_3_Step4",
    Int_Data."TMGa_3_Step6",
    Int_Data."TMGa_3_Step8",
    Int_Data."TMIn_1_Step4",
    Int_Data."TMIn_1_Step6",
    Int_Data."TMIn_1_Step8",
    Int_Data."TMIn_1_Step10",
    Int_Data."TMIn_2_Step4",
    Int_Data."TMIn_2_Step6",
    Int_Data."TMIn_2_Step8",
    Int_Data."TMIn_2_Step10",
    Int_Data."PH3_1_Step4",
    Int_Data."PH3_1_Step6",
    Int_Data."PH3_1_Step8",
    Int_Data."PH3_1_Step10",
    Int_Data."PH3_2_Step4",
    Int_Data."PH3_2_Step6",
    Int_Data."PH3_2_Step8",
    Int_Data."PH3_2_Step10",
    Int_Data."AsH3_3_100percent_Step4",
    Int_Data."AsH3_3_100percent_Step6",
    Int_Data."AsH3_3_100percent_Step8",
    Int_Data."Temperature_Step4",
    Int_Data."Temperature_Step6",
    Int_Data."Temperature_Step8",
    Int_Data."Temperature_Step10",
    String_Data.ReactorSetting_CoverStar,
    String_Data.ReactorSetting_Magazine,
    String_Data.ReactorSetting_Deck,
    Int_Data."Remaining_MO_TMGa_1",
    Int_Data."Remaining_MO_TMGa_2",
    Int_Data."Remaining_MO_TMGa_3",
    Int_Data."Remaining_MO_TMIn_1",
    Int_Data."Remaining_MO_TMIn_2",
    Int_Data."Remaining_MO_TMIn_3",
    Int_Data."Remaining_MO_DEZn_1",
    Int_Data."Remaining_MO_DMZn_1",
    Int_Data."Remaining_MO_TMAl_1",
    Int_Data."Remaining_MO_TMAl_2",
    Int_Data."Remaining_MO_TMAl_3",
    Int_Data."As-Ratio_Step4",
    Int_Data."As-Ratio_Step6",
    Int_Data."As-Ratio_Step8",
    Int_Data."Ga-Ratio_Step4",
    Int_Data."Ga-Ratio_Step6",
    Int_Data."Ga-Ratio_Step8"
FROM(
    SELECT 
        HDR.TESTHEADERID,
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
        MAX(CASE TEQ.DEVICENAME WHEN 'PLmapper' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "PLmapper",
        MAX(CASE TEQ.DEVICENAME WHEN 'XRD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "XRD",
        MAX(CASE MO.DEVICENAME||MO.DEVICESERIALNUMBER WHEN 'MOCVD7' THEN MO.DEVICENAME || MO.DEVICESERIALNUMBER END) AS "MOCVD"
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
        AND MO.TESTHEADERID(+)          = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'WG-EML'
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD7'
    GROUP BY
        HDR.TESTHEADERID,
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
            'Strain_Straing_WG' AS "Straing_WG",
            'Wavelength_Wavelength_WG' AS "Wavelength_WG",
            'Wavelength_Wavelength_Intensity' AS "Wavelength_Intensity",
            'Wavelength_Wavelength_FWHM' AS "Wavelength_FWHM",
            'EpiTT_EpiTT_WG' AS "EpiTT_WG",
            'Dulation_Step4' AS "Dulation_Step4",
            'Dulation_Step6' AS "Dulation_Step6",
            'Dulation_Step8' AS "Dulation_Step8",
            'Dulation_Step10' AS "Dulation_Step10",
            'TMGa_3_Step4' AS "TMGa_3_Step4",
            'TMGa_3_Step6' AS "TMGa_3_Step6",
            'TMGa_3_Step8' AS "TMGa_3_Step8",
            'TMIn_1_Step4' AS "TMIn_1_Step4",
            'TMIn_1_Step6' AS "TMIn_1_Step6",
            'TMIn_1_Step8' AS "TMIn_1_Step8",
            'TMIn_1_Step10' AS "TMIn_1_Step10",
            'TMIn_2_Step4' AS "TMIn_2_Step4",
            'TMIn_2_Step6' AS "TMIn_2_Step6",
            'TMIn_2_Step8' AS "TMIn_2_Step8",
            'TMIn_2_Step10' AS "TMIn_2_Step10",
            'PH3_1_Step4' AS "PH3_1_Step4",
            'PH3_1_Step6' AS "PH3_1_Step6",
            'PH3_1_Step8' AS "PH3_1_Step8",
            'PH3_1_Step10' AS "PH3_1_Step10",
            'PH3_2_Step4' AS "PH3_2_Step4",
            'PH3_2_Step6' AS "PH3_2_Step6",
            'PH3_2_Step8' AS "PH3_2_Step8",
            'PH3_2_Step10' AS "PH3_2_Step10",
            'AsH3_3_100percent_Step4' AS "AsH3_3_100percent_Step4",
            'AsH3_3_100percent_Step6' AS "AsH3_3_100percent_Step6",
            'AsH3_3_100percent_Step8' AS "AsH3_3_100percent_Step8",
            'Temperature_Step4' AS "Temperature_Step4",
            'Temperature_Step6' AS "Temperature_Step6",
            'Temperature_Step8' AS "Temperature_Step8",
            'Temperature_Step10' AS "Temperature_Step10",
            'Remaining_MO_Remaining_MO_TMGa_1' AS "Remaining_MO_TMGa_1",
            'Remaining_MO_Remaining_MO_TMGa_2' AS "Remaining_MO_TMGa_2",
            'Remaining_MO_Remaining_MO_TMGa_3' AS "Remaining_MO_TMGa_3",
            'Remaining_MO_Remaining_MO_TMIn_1' AS "Remaining_MO_TMIn_1",
            'Remaining_MO_Remaining_MO_TMIn_2' AS "Remaining_MO_TMIn_2",
            'Remaining_MO_Remaining_MO_TMIn_3' AS "Remaining_MO_TMIn_3",
            'Remaining_MO_Remaining_MO_DEZn_1' AS "Remaining_MO_DEZn_1",
            'Remaining_MO_Remaining_MO_DMZn_1' AS "Remaining_MO_DMZn_1",
            'Remaining_MO_Remaining_MO_TMAl_1' AS "Remaining_MO_TMAl_1",
            'Remaining_MO_Remaining_MO_TMAl_2' AS "Remaining_MO_TMAl_2",
            'Remaining_MO_Remaining_MO_TMAl_3' AS "Remaining_MO_TMAl_3",
            'As-Ratio_Step4' AS "As-Ratio_Step4",
            'As-Ratio_Step6' AS "As-Ratio_Step6",
            'As-Ratio_Step8' AS "As-Ratio_Step8",
            'Ga-Ratio_Step4' As "Ga-Ratio_Step4",
            'Ga-Ratio_Step6' As "Ga-Ratio_Step6",
            'Ga-Ratio_Step8' As "Ga-Ratio_Step8",
            'SORTED_DATA_STARTTIME_SORTED' AS "STARTTIME_SORTED",
            'SORTED_DATA_SORTNUMBER' AS "SORTNUMBER"
        )
    )Int_Data
,(
    SELECT
        HDR.TESTHEADERID,
        MAX(CASE PARAMETERNAME WHEN 'CoverStar' THEN VALUESTRING END) AS ReactorSetting_CoverStar,
        MAX(CASE PARAMETERNAME WHEN 'Magazine' THEN VALUESTRING END) AS ReactorSetting_Magazine,
        MAX(CASE PARAMETERNAME WHEN 'Deck' THEN VALUESTRING END) AS ReactorSetting_Deck,
        MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_LotNumber_5' THEN VALUESTRING END) AS "FIVE_SERIALNUMBER",
        MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_LotNumber_9' THEN VALUESTRING END) AS "NINE_SERIALNUMBER",
        MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_BATCHNUMBER_SORTED' THEN VALUESTRING END) AS "SIX_BATCHNUMBER"
    FROM
        TDSMFG.TESTHEADER           HDR,
        TDSMFG.TESTHEADERSTEP       HDS,
        TDSMFG.OPERATIONSTEP        OPS,
        TDSMFG.PARAMETER            PAR,
        TDSMFG.STRINGPARAMETER      STP,
        TDSMFG.STRINGMEASUREMENT    SMM,
        TDSMFG.TESTEQUIPMENT        MO
        
    WHERE 1=1
        AND HDR.TESTHEADERID            = HDS.TESTHEADERID
        AND HDS.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND HDS.TESTHEADERSTEPID        = STP.TESTHEADERSTEPID
        AND STP.STRINGPARAMETERID       = SMM.STRINGPARAMETERID
        AND PAR.OPERATIONSTEPID         = OPS.OPERATIONSTEPID
        AND PAR.PARAMETERID             = STP.PARAMETERID
        AND MO.TESTHEADERID(+)          = HDR.TESTHEADERID
        AND HDR.OPERATION               = 'WG-EML'
        AND MO.DEVICENAME || MO.DEVICESERIALNUMBER = 'MOCVD7'
    GROUP BY
        HDR.TESTHEADERID
)String_Data

WHERE 1=1
  AND Int_Data.TESTHEADERID = String_Data.TESTHEADERID
ORDER BY
    Int_Data."SORTNUMBER",
    STARTTIME
