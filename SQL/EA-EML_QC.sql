SELECT 
    Format1_Int_Data.STARTTIME,
    Format1_String_Data."NINE_SERIALNUMBER",
    Format1_Int_Data.PARTNUMBER,
    Format1_Int_Data."SORTED_DATA_STARTTIME_SORTED",
    Format1_Int_Data."SORTED_DATA_SORTNUMBER",
    Format1_Int_Data.SERIALNUMBER,
    Format1_Int_Data.OPERATORID,
    Format1_Int_Data.BATCHNUMBER,
    Format1_Int_Data."MOCVD",
    Format1_Int_Data."SEM",
    Format1_Int_Data."PLmapper",
    Format1_Int_Data."XRD",
    Format1_Int_Data."Thickness_Cap",
    Format1_Int_Data."Thickness_Core",
    Format1_Int_Data."Thickness_Total",
    Format1_Int_Data."Xray_Thickness",
    Format1_Int_Data."Xray_Strain",
    Format1_String_Data."InPBoardLot_No",
    Format1_Int_Data."InPBoardLot_CC",
    Format1_Int_Data."InPBoardLot_EPD",
    Format2.STARTTIME AS "Format2_STARTTIME",
    Format2.SERIALNUMBER AS "Format2_SERIALNUMBER",
    Format2.BATCHNUMBER AS "Format2_BATCHNUMBER",
    Format2."Mapper_PL_AreaStandards_B2",
    Format2."Mapper_Average_Mapper_PL_Lambda"-Format2."Mapper_TargetWavelength" AS "Mapper_Average_Mapper_PL_Lambda_Diff",
    Format2."Mapper_Average_Mapper_PL_Intensity",
    Format2."Mapper_Average_Mapper_PL_FWHM",
    (Format2."A_Tails"+Format2."B_Tails"+Format2."C_Tails"+Format2."D_Tails") / 4 as "Mapper_Average_Mapper_PL_Tails",
    Format2."AreaStandards_1",
    Format2."AreaStandards_2"
FROM(
    SELECT
        HDR.TESTHEADERID,
        HDR.STARTTIME,
        HDR.PARTNUMBER,
        HDR.SERIALNUMBER,
        HDR.OPERATORID,
        HDR.BATCHNUMBER,
        MAX(CASE TEQ.DEVICENAME WHEN 'SEM' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "SEM",
        MAX(CASE TEQ.DEVICENAME WHEN 'PLmapper' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "PLmapper",
        MAX(CASE TEQ.DEVICENAME WHEN 'XRD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "XRD",
        MAX(CASE TEQ.DEVICENAME WHEN 'MOCVD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "MOCVD",
        OPS.OPERATIONSTEPNAME || '_' || PARAMETERNAME AS TESTSTEPS,
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
        AND HDR.OPERATION               in ('EA-EML_F1_Format1', 'EA-EML_F7_Format1')
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
            'Thickness_Thickness_Cap' AS "Thickness_Cap",
            'Thickness_Thickness_Core' AS "Thickness_Core",
            'Thickness_Thickness_Total' AS "Thickness_Total",
            'XRayDiffraction_Xray_Thickness' AS "Xray_Thickness",
            'XRayDiffraction_Xray_Strain' AS "Xray_Strain",
            'InPBoardLot_InPBoardLot_CC' AS "InPBoardLot_CC",
            'InPBoardLot_InPBoardLot_EPD' AS "InPBoardLot_EPD",
            'SORTED_DATA_STARTTIME_SORTED' AS "SORTED_DATA_STARTTIME_SORTED",
            'SORTED_DATA_SORTNUMBER' AS "SORTED_DATA_SORTNUMBER"

        )
    )Format1_Int_Data
    LEFT JOIN
    (
        SELECT
            HDR.TESTHEADERID,
            MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_LotNumber_5' THEN VALUESTRING END) AS "FIVE_SERIALNUMBER",
            MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'SORTED_DATA_LotNumber_9' THEN VALUESTRING END) AS "NINE_SERIALNUMBER",
            MAX(CASE OPERATIONSTEPNAME || '_' || PARAMETERNAME WHEN 'InPBoardLot_InPBoardLot_No' THEN VALUESTRING END) AS "InPBoardLot_No"
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
            AND HDR.OPERATION               IN ('EA-EML_F1_Format1', 'EA-EML_F7_Format1')
        GROUP BY
            HDR.TESTHEADERID
    )Format1_String_Data
    ON Format1_Int_Data.TESTHEADERID = Format1_String_Data.TESTHEADERID
    FULL JOIN
    (
        SELECT
            HDR.TESTHEADERID,
            HDR.STARTTIME,
            HDR.PARTNUMBER,
            HDR.SERIALNUMBER,
            HDR.OPERATORID,
            HDR.BATCHNUMBER,
            MAX(CASE TEQ.DEVICENAME WHEN 'PLmapper' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "PLmapper",
            MAX(CASE TEQ.DEVICENAME WHEN 'MOCVD' THEN TEQ.DEVICENAME || TEQ.DEVICESERIALNUMBER END) AS "MOCVD",
            OPS.OPERATIONSTEPNAME || '_' || PARAMETERNAME AS TESTSTEPS,
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
            AND HDR.OPERATION               IN ('EA-EML_F1_Format2', 'EA-EML_F7_Format2')
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
                'Mapper_Average_Mapper_PL_Lambda' AS "Mapper_Average_Mapper_PL_Lambda",
                'Mapper_Average_Mapper_PL_Intensity' AS "Mapper_Average_Mapper_PL_Intensity",
                'Mapper_Average_Mapper_PL_FWHM' AS "Mapper_Average_Mapper_PL_FWHM",
                'Mapper_Average_Mapper_PL_AreaStandards_B2' AS "Mapper_PL_AreaStandards_B2",
                'Mapper_Adjust_Mapper_TargetWavelength' AS "Mapper_TargetWavelength",
                'A_Tails' AS "A_Tails",
                'B_Tails' AS "B_Tails",
                'C_Tails' AS "C_Tails",
                'D_Tails' AS "D_Tails",
                'Measurements_AreaStandards_1' AS "AreaStandards_1",
                'Measurements_AreaStandards_2' AS "AreaStandards_2"
            )
    )Format2
    ON Format1_Int_Data.SERIALNUMBER = Format2.SERIALNUMBER

ORDER BY
    Format1_Int_Data."SORTED_DATA_STARTTIME_SORTED",
    STARTTIME

