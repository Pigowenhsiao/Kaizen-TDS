SELECT
    "STARTTIME",
    "PARTNUMBER",
    "SERIALNUMBER",
    "NINE_SERIALNUMBER",
    "OPERATORID",
    "BATCHNUMBER",
    "SORTED_DATA_STARTTIME_SORTED",
    "SORTED_DATA_SORTNUMBER",
    "MOCVD",
    "SEM",
    "PLmapper",
    "XRD",
    "Thickness_Cap",
    "Thickness_p-Q_115",
    "Thickness_Core",
    "Xray_Thickness",
    "Xray_Strain",
    "Format2_STARTTIME",
    "Format2_SERIALNUMBER",
    "Center_Lambda",
    "AB_Lambda",
    "AC_Lambda",
    "BC_Lambda",
    "CD_Lambda",
    "Mapper_Average_Mapper_PL_Center-ABCD_Diff-Max",
    "Mapper_Average_Mapper_PL_Center-ABCD_Diff-Min",
    CASE
        WHEN ABS("Mapper_Average_Mapper_PL_Center-ABCD_Diff-Max") > ABS("Mapper_Average_Mapper_PL_Center-ABCD_Diff-Min") THEN "Mapper_Average_Mapper_PL_Center-ABCD_Diff-Max"
        ELSE "Mapper_Average_Mapper_PL_Center-ABCD_Diff-Min"
    END AS "Mapper_Average_Mapper_PL_Center-ABCD_Diff",
    "Mapper_Average_Mapper_PL_Lambda_Diff",
    "Mapper_Average_Mapper_PL_Intensity",
    "Mapper_Average_Mapper_PL_FWHM",
    "Mapper_Average_Mapper_PL_Tails"
    
FROM(
    SELECT 
        Format1_Int_Data.STARTTIME AS "STARTTIME",
        Format1_Int_Data.PARTNUMBER AS "PARTNUMBER",
        Format1_Int_Data.SERIALNUMBER AS "SERIALNUMBER",
        Format1_String_Data."NINE_SERIALNUMBER" AS "NINE_SERIALNUMBER",
        Format1_Int_Data.OPERATORID AS "OPERATORID",
        Format1_Int_Data.BATCHNUMBER AS "BATCHNUMBER",
        Format1_Int_Data."SORTED_DATA_STARTTIME_SORTED" as "SORTED_DATA_STARTTIME_SORTED",
        Format1_Int_Data."SORTED_DATA_SORTNUMBER" as "SORTED_DATA_SORTNUMBER",
        Format1_Int_Data."MOCVD",
        Format1_Int_Data."SEM",
        Format1_Int_Data."PLmapper",
        Format1_Int_Data."XRD",
        Format1_Int_Data."Thickness_Cap",
        Format1_Int_Data."Thickness_p-Q_115",
        Format1_Int_Data."Thickness_InP_Spacer1",
        Format1_Int_Data."Thickness_Core",
        Format1_Int_Data."Xray_Thickness",
        Format1_Int_Data."Xray_Strain",
        Format2.STARTTIME AS "Format2_STARTTIME",
        Format2.SERIALNUMBER AS "Format2_SERIALNUMBER",
        Format2."Center_Lambda" AS "Center_Lambda",
        Format2."AB_Lambda" AS "AB_Lambda",
        Format2."AC_Lambda" AS "AC_Lambda",
        Format2."BC_Lambda" AS "BC_Lambda",
        Format2."CD_Lambda" AS "CD_Lambda",
        GREATEST(Format2."AB_DeltaLambda", Format2."AC_DeltaLambda", Format2."BC_DeltaLambda", Format2."CD_DeltaLambda") AS "Mapper_Average_Mapper_PL_Center-ABCD_Diff-Max",
        LEAST(Format2."AB_DeltaLambda", Format2."AC_DeltaLambda", Format2."BC_DeltaLambda", Format2."CD_DeltaLambda") AS "Mapper_Average_Mapper_PL_Center-ABCD_Diff-Min",
        Format2."Mapper_Average_Mapper_PL_Lambda"-Format2."Mapper_TargetWavelength" AS "Mapper_Average_Mapper_PL_Lambda_Diff",
        Format2."Mapper_Average_Mapper_PL_Intensity",
        Format2."Mapper_Average_Mapper_PL_FWHM",
        (Format2."A_Tails"+Format2."B_Tails"+Format2."C_Tails"+Format2."D_Tails") / 4 as "Mapper_Average_Mapper_PL_Tails"
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
            AND HDR.OPERATION               IN ('LD-EML_F1_Format1', 'LD-EML_F2_Format1', 'LD-EML_F6_Format1')
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
                'Thickness_Thickness_p-Q_115' AS "Thickness_p-Q_115",
                'Thickness_Thickness_InP_Spacer1' AS "Thickness_InP_Spacer1",
                'Thickness_Thickness_Core' AS "Thickness_Core",
                'XRayDiffraction_Xray_Thickness' AS "Xray_Thickness",
                'XRayDiffraction_Xray_Strain' AS "Xray_Strain",
                'SORTED_DATA_STARTTIME_SORTED' AS "SORTED_DATA_STARTTIME_SORTED",
                'SORTED_DATA_SORTNUMBER' AS "SORTED_DATA_SORTNUMBER"
            )
        )Format1_Int_Data
        LEFT JOIN
        (
            SELECT
                HDR.TESTHEADERID,
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
                AND HDR.OPERATION               = 'LD-EML_F6_Format1'
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
                AND HDR.OPERATION               IN ('LD-EML_F1_Format2', 'LD-EML_F2_Format2', 'LD-EML_F6_Format2')
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
                    'Mapper_Adjust_Mapper_TargetWavelength' AS "Mapper_TargetWavelength",
                    'Center_Lambda' AS "Center_Lambda",
                    'AB_Lambda' AS "AB_Lambda",
                    'AC_Lambda' AS "AC_Lambda",
                    'BC_Lambda' AS "BC_Lambda",
                    'CD_Lambda' AS "CD_Lambda",
                    'AB_DeltaLambda' AS "AB_DeltaLambda",
                    'AC_DeltaLambda' AS "AC_DeltaLambda",
                    'BC_DeltaLambda' AS "BC_DeltaLambda",
                    'CD_DeltaLambda' AS "CD_DeltaLambda",
                    'A_Tails' AS "A_Tails",
                    'B_Tails' AS "B_Tails",
                    'C_Tails' AS "C_Tails",
                    'D_Tails' AS "D_Tails"
                )
        )Format2
        ON Format1_Int_Data.SERIALNUMBER = Format2.SERIALNUMBER
)
ORDER BY
    "SORTED_DATA_STARTTIME_SORTED",
    STARTTIME
