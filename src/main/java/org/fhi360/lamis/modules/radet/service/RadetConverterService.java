package org.fhi360.lamis.modules.radet.service;

import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.fhi360.lamis.modules.radet.util.PatientEntry;
import org.fhi360.lamis.modules.radet.util.RadetEntry;
import org.fhi360.lamis.modules.radet.util.RegimenIntrospector;
import org.lamisplus.modules.base.config.ApplicationProperties;
import org.lamisplus.modules.lamis.legacy.domain.entities.enumerations.ClientStatus;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.messaging.simp.SimpMessageSendingOperations;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.*;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.TemporalAdjusters;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicInteger;

@Service
@RequiredArgsConstructor
@Slf4j
public class RadetConverterService {
    private final JdbcTemplate jdbcTemplate;
    private final ApplicationProperties applicationProperties;
    private final SimpMessageSendingOperations messagingTemplate;
    private final static String BASE_DIR = "/radet/";

    public void convertExcel(Long facilityId, LocalDate cohortBegin, LocalDate cohortEnd, LocalDate reportingPeriod, boolean today) {

        LocalDate reportingDateEnd = reportingPeriod.with(TemporalAdjusters.lastDayOfMonth());
        LocalDate reportingDateBegin = reportingPeriod.with(TemporalAdjusters.firstDayOfMonth());
        if (today) {
            reportingDateEnd = LocalDate.now();
        }

        cohortBegin = cohortBegin.with(TemporalAdjusters.firstDayOfMonth());
        cohortEnd = cohortEnd.with(TemporalAdjusters.lastDayOfMonth());

        Workbook workbook = new SXSSFWorkbook(100);  // turn off auto-flushing and accumulate all rows in memory
        Sheet sheet = workbook.createSheet();

        //Create a new font
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());

        //Create a style and set the font into it
        CellStyle style = getCellStyle(workbook);

        CellStyle numericStyle = workbook.createCellStyle();
        numericStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0"));
        sheet.setDefaultColumnStyle(0, numericStyle);
        sheet.setDefaultColumnStyle(10, numericStyle);
        sheet.setDefaultColumnStyle(22, numericStyle);
        sheet.setDefaultColumnStyle(23, numericStyle);
        sheet.setDefaultColumnStyle(26, numericStyle);

        int rowNum = 0;
        int cellNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(cellNum++);
        cell.setCellValue("S/No.");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Patient Id");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Hospital Num");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Household Unique No");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Received OVC Service?");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Sex");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Weight (Kg)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Birth (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("ART Start Date (yyyy-mm-dd");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Last Pickup Date (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Months of ARV Refill");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("TPT in the Last 2 years");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("If Yes to TPT, date of TPT Start (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("TPT Completion date (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Regimen Line at ART Start");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Regimen at ART Start");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Regimen Line");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current ART Regimen");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Regimen Switch");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Pregnancy Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Full Disclosure?");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Enrolled on OTZ?");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Number of Support Group (OTZ Club) meeting attended");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Number of OTZ Modules completed");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date Enrolled on OTZ (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Viral Load Sample Collection (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Viral Load (c/ml)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Current Viral Load (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Viral Load Indication");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current ART Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Current ART Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("If Dead, Cause of Dead");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("If Transferred out, new Facility");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("ART Enrollment Setting");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Client Receiving DMOC Service?");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date Commenced DMOC (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Type of DMOC");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Return of DMOC Client to Facility (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Enhanced Adherence Counselling (EAC) Commenced?");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Commencement of EAC (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Number of EAC Sessions Completed");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Repeat Viral Load - Post EAC VL Sample Collected?");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Repeat Viral Load Result Received - Post EAC (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("IPT Type");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Screening for Chronic Conditions");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Co-morbidities");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Referred for Further Care");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Cervical Cancer Screening (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Screened for Cervical Cancer");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Cervical Cancer Screening Type");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Cervical Cancer Screening Method");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Result of Cervical Cancer Screening");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Precancerous Lesions Treatment Methods");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum);
        cell.setCellValue("Case Manager");
        cell.setCellStyle(style);

        //Create a date format for date columns
        CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));
        //style.setAlignment(CellStyle.ALIGN_RIGHT);
        File file = new File(applicationProperties.getTempDir() + BASE_DIR + getFacility(facilityId).trim());
        FileUtils.deleteQuietly(file);

        List<PatientEntry> entries = new ArrayList<>();
        AtomicInteger atomicInteger = new AtomicInteger();
        AtomicInteger sn = new AtomicInteger();
        List<RadetEntry> radetEntries = new ArrayList<>();
        List<RadetEntry> finalEntries = new ArrayList<>();
        String query = "SELECT DISTINCT id, hospital_num, unique_id, enrollment_setting, gender, date_birth, extra->'ovc'->>'householdUniqueNo' household_unique_no," +
                "   jsonb_array_length(extra->'ovc'->'servicesProvided') > 0 services_provided, " +
                "   date_started, DATEDIFF('YEAR', date_birth, CURRENT_DATE) age, status_at_registration FROM patient pl WHERE " +
                "   facility_id = ? AND date_started BETWEEN ? AND ? and archived = false and " +
                "   cast(extra->>'art' as boolean) = true ";
        jdbcTemplate.query(query, resultSet -> {
            while (resultSet.next()) {
                PatientEntry patientEntry = new PatientEntry();
                Long patientId = resultSet.getLong("id");
                String uniqueId = StringUtils.trimToEmpty(resultSet.getString("unique_id"));
                String hospitalNum = resultSet.getString("hospital_num");
                String gender = resultSet.getString("gender");
                String enrollmentSetting = StringUtils.trimToEmpty(resultSet.getString("enrollment_setting"));
                LocalDate dateBirth = resultSet.getObject("date_birth", LocalDate.class);
                LocalDate dateStarted = resultSet.getObject("date_started", LocalDate.class);
                int age = resultSet.getInt("age");

                boolean servicesProvided = resultSet.getBoolean("services_provided");
                patientEntry.setHouseholdUniqueNo(resultSet.getString("household_unique_no"));
                if (StringUtils.isNotEmpty(resultSet.getString("household_unique_no"))) {
                    if (servicesProvided) {
                        patientEntry.setServicesProvided("Yes");
                    } else {
                        patientEntry.setServicesProvided("No");
                    }
                }
                patientEntry.setPatientId(patientId);
                patientEntry.setUniqueId(uniqueId);
                patientEntry.setHospitalNum(hospitalNum);
                patientEntry.setSex(gender);
                patientEntry.setArtEnrollmentSetting(enrollmentSetting);
                patientEntry.setDob(dateBirth);
                patientEntry.setDateStarted(dateStarted);
                patientEntry.setAge(age);
                patientEntry.setStatusAtRegistration(resultSet.getString("status_at_registration"));
                entries.add(patientEntry);

                RadetEntry entry = new RadetEntry();
                entry.setSn(sn.incrementAndGet());
                entry.setPatientId(patientId.toString());
                entry.setUniqueId(uniqueId);
                entry.setHospitalNum(hospitalNum);
                entry.setSex(StringUtils.equals(gender, "MALE") ? "Male" : "Female");
                entry.setDob(convertToDateViaInstant(dateBirth));
                if (dateStarted != null) {
                    entry.setArtStartDate(convertToDateViaInstant(dateStarted));
                }
                entry.setStatusAtRegistration(patientEntry.getStatusAtRegistration());
                entry.setArtEnrollmentSetting(enrollmentSetting);
                radetEntries.add(entry);
            }
            return null;
        }, facilityId, cohortBegin, cohortEnd);
        ExecutorService executorService = Executors.newFixedThreadPool(3);
        LocalDate finalReportingDateEnd = reportingDateEnd;
        entries.forEach(patientEntry -> executorService.submit(() ->
                finalEntries.add(updateEntry(atomicInteger, sn, patientEntry, radetEntries, reportingDateBegin,
                        finalReportingDateEnd, entries.size()))));
        executorService.shutdown();
        while (!executorService.isTerminated()) {

        }
        messagingTemplate.convertAndSend("/topic/radet/status", "Building sheet...");
        buildSheet(sheet, dateStyle, finalEntries);
        messagingTemplate.convertAndSend("/topic/radet/status", "Writing out file...");
        try {
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private RadetEntry updateEntry(AtomicInteger atomicInteger, AtomicInteger sn, PatientEntry patientEntry,
                                   List<RadetEntry> entries, LocalDate reportingDateBegin, LocalDate reportingDateEnd, int total) {
        RadetEntry entry = entries.stream()
                .filter(e -> e.getPatientId().equals(patientEntry.getPatientId().toString()))
                .findFirst().get();
        messagingTemplate.convertAndSend("/topic/radet/status", String.format("Analysing patient %s of %s",
                atomicInteger.incrementAndGet(), total));
        try {
            entry.setHouseholdUniqueNo(patientEntry.getHouseholdUniqueNo());
            entry.setReceivedOvcService(patientEntry.getServicesProvided());
            jdbcTemplate.query("SELECT body_weight FROM clinic WHERE patient_id = ? AND " +
                    "date_visit <= ? AND body_weight > 0 and archived = false ORDER BY date_visit DESC LIMIT 1", rs -> {
                entry.setWeight(rs.getDouble("body_weight"));
            }, patientEntry.getPatientId(), reportingDateEnd);

            jdbcTemplate.query("SELECT cast(extra->'otz'->>'fullDisclosure' as boolean) full_disclosure," +
                    "cast(extra->'otz'->>'dateEnrolledOnOTZ' as date) date_enrolled, " +
                    "cast(extra->'otz'->>'attendedLastOTZMeeting' as boolean) attended_meeting," +
                    "cast(extra->'otz'->>'modulesCompleted' as integer) modules_completed FROM clinic WHERE patient_id = ? AND " +
                    "cast(extra->'otz'->>'dateEnrolledOnOTZ' as date) <= ? AND extra->'otz' is not null and archived = false ORDER BY date_visit DESC ", rs -> {
                if (rs.getBoolean("full_disclosure")) {
                    entry.setFullDisclosure("Yes");
                } else {
                    entry.setFullDisclosure("No");
                }
                if (rs.getDate("date_enrolled") != null) {
                    entry.setDateEnrolledOnOTZ(rs.getDate("date_enrolled"));
                    entry.setEnrolledOnOTZ("Yes");
                }
                int modulesCompleted = rs.getInt("modules_completed");
                Integer modules = entry.getOtzModulesCompleted();
                if (modules == null) {
                    entry.setOtzModulesCompleted(modulesCompleted);
                } else if (modulesCompleted > modules) {
                    entry.setOtzModulesCompleted(modulesCompleted);
                }
                if (rs.getBoolean("attended_meeting")) {
                    Integer meetings = entry.getNumberOfOTZMeetings();
                    if (meetings != null) {
                        entry.setNumberOfOTZMeetings(++meetings);
                    } else {
                        entry.setNumberOfOTZMeetings(1);
                    }
                }
            }, patientEntry.getPatientId(), reportingDateEnd);
            jdbcTemplate.query("SELECT date, data->'cervicalCancerScreening'->>'screeningMethod' screening_method, " +
                    "data->'cervicalCancerScreening'->>'screeningType' screening_type, data->'cervicalCancerScreening'->>'screeningResult' screening_result, " +
                    "data->'cervicalCancerScreening'->>'precancerousLesionsTreatmentMethod' treatment_method FROM observation WHERE patient_id = ? AND " +
                    "type = 'CERVICAL_CANCER_SCREENING' and date <= ? AND archived = false ORDER BY date DESC ", rs -> {
                entry.setScreenedForCervicalCancer("Yes");
                entry.setCancerScreeningDate(rs.getDate("date"));
                String screeningMethod = rs.getString("screening_method");
                switch (screeningMethod) {
                    case "VIA":
                        screeningMethod = "Visual Inspection with Acetric Acid (VIA)";
                        break;
                    case "VILI":
                        screeningMethod = "Visual Inspection with Lugos Iodine (VILI)";
                        break;
                    case "PAP_SMEAR":
                        screeningMethod = "PAP Smear";
                        break;
                }
                entry.setCervicalCancerScreeningMethod(screeningMethod);
                String screeningResult = rs.getString("screening_result");
                switch (screeningResult) {
                    case "NEGATIVE":
                        screeningResult = "Negative";
                        break;
                    case "POSITIVE":
                        screeningResult = "Positive";
                        break;
                    case "SUSPICIOUS":
                        screeningResult = "Suspicious Cancerous Lesions";
                        break;
                }
                entry.setResultOfCervicalCancerScreening(screeningResult);
                String screeningType = rs.getString("screening_type");
                switch (screeningType) {
                    case "FIRST_TIME":
                        screeningType = "First Time";
                        break;
                    case "FOLLOWUP":
                        screeningType = "Followup after previous negative result or suspected cancer";
                        break;
                    case "POST_TREATMENT_FOLLOWUP":
                        screeningType = "Post-treatment Followup";
                        break;
                }
                entry.setCervicalCancerScreeningType(screeningType);
                String treatmentMethod = rs.getString("treatment_method");
                switch (treatmentMethod) {
                    case "CRYOTHERAPY":
                        treatmentMethod = "Cryotherapy";
                        break;
                    case "THERMAL_ABLATION":
                        treatmentMethod = "Thermal Ablation/ Thermocoagulation";
                        break;
                    case "LEETZ_LEEP":
                        treatmentMethod = "LEETZ/ LEEP";
                        break;
                    case "CONIZATION":
                        treatmentMethod = "Conization Knifer/ Lagor";
                        break;
                }
                entry.setPrecancerousScreeningMethods(treatmentMethod);
            }, patientEntry.getPatientId(), reportingDateEnd);

            //Current status on or before the reporting date
            String query = "SELECT status, cause_of_death, date_status, extra->>'facilityTransferredTo' to_facility FROM status_history WHERE " +
                    "patient_id = ? AND date_status <= ? and archived = false ORDER BY date_status DESC, id DESC LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                entry.setCurrentArtStatus(StringUtils.trimToEmpty(rs.getString("status")));
                entry.setCurrentArtStatusDate(rs.getDate("date_status"));
                entry.setCauseOfDeath(rs.getString("cause_of_death"));
                entry.setTransferOutFacility(rs.getString("to_facility"));
            }, patientEntry.getPatientId(), reportingDateEnd);

            // pickup before the reporting date
            int age = patientEntry.getAge();
            final Date[] ltfDate = {null};
            query = "select date_visit, cast(jsonb_extract_path_text(l,'duration') as integer) duration, date_visit + " +
                    "cast(jsonb_extract_path_text(l,'duration') as integer) + INTERVAL '28 DAYS' < ? as ltfu, " +
                    "date(date_visit + cast(jsonb_extract_path_text(l,'duration') as integer) + INTERVAL '29 DAYS') ltfu_date, r.description regimen, " +
                    "t.description as regimen_type from pharmacy p, jsonb_array_elements(lines) with ordinality a(l) " +
                    "join regimen r on r.id = cast(jsonb_extract_path_text(l,'regimen_id') as integer)  " +
                    "join regimen_type t on t.id = r.regimen_type_id where t.id in (1,2,3,4,14) and p.patient_id = ? and " +
                    "date_visit <= ? and p.archived = false ORDER BY p.date_visit DESC, duration DESC LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                Date dlr = rs.getDate("date_visit");
                ltfDate[0] = rs.getDate("ltfu_date");
                entry.setLastPickupDate(dlr);
                int duration = rs.getInt("duration");
                entry.setLtfu(rs.getBoolean("ltfu"));
                int monthRefill = duration / 30;
                if (monthRefill <= 0) {
                    monthRefill = 1;
                }
                entry.setMonthsOfRefill(monthRefill);
                //regimen at last pickup
                String regimenType = StringUtils.trimToEmpty(rs.getString("regimen_type"));
                if (regimenType.contains("ART First Line Adult")) {
                    regimenType = "Adult.1st.Line";
                } else if (regimenType.contains("ART Second Line Adult")) {
                    regimenType = "Adult.2nd.Line";
                } else if (regimenType.contains("ART First Line Children")) {
                    regimenType = "Peds.1st.Line";
                } else if (regimenType.contains("ART Second Line Children")) {
                    regimenType = "Peds.2nd.Line";
                } else if (regimenType.contains("Third Line")) {
                    if (age < 5) {
                        regimenType = "Peds.3rd.Line";
                    } else {
                        regimenType = "Adult.3rd.Line";
                    }
                } else {
                    regimenType = "";
                }
                entry.setCurrentRegimenLine(regimenType);
                if (StringUtils.isNotBlank(regimenType)) {
                    String regimen = RegimenIntrospector.resolveRegimen(StringUtils.trimToEmpty(rs.getString("regimen")));
                    entry.setCurrentArtRegimen(regimen);
                }

            }, reportingDateEnd, patientEntry.getPatientId(), reportingDateEnd);

            String currentStatus = entry.getCurrentArtStatus();
            if (entry.getLastPickupDate() != null) {
                //If the last refill date plus refill duration plus 28 days (or 30 days) is before the last day of the reporting month this patient is LTFU
                if (entry.isLtfu()) {
                    currentStatus = "LTFU";
                    if (entry.getCurrentArtStatus().equalsIgnoreCase(ClientStatus.ART_TRANSFER_OUT.name())) {
                        currentStatus = "Transferred Out";
                    } else if (entry.getCurrentArtStatus().equalsIgnoreCase(ClientStatus.STOPPED_TREATMENT.name())) {
                        currentStatus = "Stopped";
                    } else if (entry.getCurrentArtStatus().equalsIgnoreCase(ClientStatus.KNOWN_DEATH.name())) {
                        currentStatus = "Dead";
                    }
                    if (currentStatus.equals("LTFU")) {
                        entry.setCurrentArtStatusDate(ltfDate[0]);
                    }
                } else {
                    if (currentStatus.equals(ClientStatus.ART_TRANSFER_IN.name())) {
                        currentStatus = "Active-Transfer In";
                    } else if (currentStatus.equals(ClientStatus.ART_RESTART.name())) {
                        currentStatus = "Active-Restart";
                    } else if (currentStatus.equalsIgnoreCase(ClientStatus.ART_TRANSFER_OUT.name())) {
                        currentStatus = "Transferred Out";
                    } else if (currentStatus.equalsIgnoreCase(ClientStatus.STOPPED_TREATMENT.name())) {
                        currentStatus = "Stopped";
                    } else if (currentStatus.equalsIgnoreCase(ClientStatus.KNOWN_DEATH.name())) {
                        currentStatus = "Dead";
                    } else {
                        currentStatus = "Active";
                    }
                }
            } /*else {
                if (entry.getCurrentArtStatus().equalsIgnoreCase(ClientStatus.ART_TRANSFER_OUT.name()) ||
                        (entry.getCurrentArtStatus().equalsIgnoreCase(ClientStatus.KNOWN_DEATH.name())
                                || (entry.getCurrentArtStatus().equalsIgnoreCase(ClientStatus.STOPPED_TREATMENT.name())))) {
                    currentStatus = entry.getCurrentArtStatus();
                    switch (currentStatus) {
                        case "ART_TRANSFER_OUT":
                            currentStatus = "Transferred Out";
                            break;
                        case "KNOWN_DEATH":
                            currentStatus = "Dead";
                            break;
                        case "STOPPED_TREATMENT":
                            currentStatus = "Stopped";
                    }
                } else {
                    currentStatus = "LTFU";
                }
            }*/
            entry.setCurrentArtStatus(currentStatus);
            if (currentStatus.equals("Active") || currentStatus.equals("Active-Transfer In")) {
                entry.setCurrentArtStatusDate(entry.getLastPickupDate());
            }
            //TPT Start
            query = "SELECT date_visit, p.extra->'ipt'->>'type' ipt_type, cast(p.extra->'ipt'->>'dateCompleted' as date)  date_completed FROM pharmacy p, " +
                    "jsonb_array_elements(lines) with ordinality a(l) WHERE p.patient_id = ? AND date_visit BETWEEN (? + INTERVAL '-24 MONTHS') " +
                    "AND ? AND cast(jsonb_extract_path_text(l,'regimen_type_id') as integer ) = 15 and p.archived = false  and p.extra->'ipt'->>'type' " +
                    "is not null ORDER BY date_visit";
            jdbcTemplate.query(query, rs -> {
                String type = rs.getString("ipt_type");
                if (type != null) {
                    entry.setIptInLast2Years("Yes");
                }
                if (StringUtils.equals(type, "START_INITIATION") || StringUtils.equals(type, "FOLLOWUP_INITIATION")) {
                    entry.setIptStartDate(rs.getDate("date_visit"));
                }
                Date completion = rs.getDate("date_completed");
                if (completion != null) {
                    entry.setIptCompletionsDate(completion);
                }
            }, patientEntry.getPatientId(), reportingDateBegin, reportingDateEnd);

            if (!StringUtils.equals("Yes", entry.getIptInLast2Years())) {
                query = "SELECT date_visit FROM pharmacy p, jsonb_array_elements(lines) with ordinality a(l) WHERE p.patient_id = ? AND " +
                        "date_visit BETWEEN (? + INTERVAL '-24 MONTHS') AND ? AND cast(jsonb_extract_path_text(l,'regimen_type_id') as integer ) = 15 " +
                        "and p.archived = false ORDER BY date_visit LIMIT 1";
                this.jdbcTemplate.query(query, rs -> {
                    entry.setIptInLast2Years("Yes");
                    entry.setIptStartDate(rs.getDate("date_visit"));
                }, patientEntry.getPatientId(), reportingDateBegin, reportingDateEnd);
            }

            //Regimen at start of ART
            query = "SELECT t.description regimen_type, r.description regimen FROM clinic JOIN regimen r ON r.id = regimen_id " +
                    "JOIN regimen_type t ON t.id = r.regimen_type_id WHERE patient_id = ? AND date_visit <= ? AND commence = true " +
                    "and archived = false ORDER BY date_visit LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                String regimenTypeStart = StringUtils.trimToEmpty(rs.getString("regimen_type"));
                if (regimenTypeStart.contains("ART First Line Adult")) {
                    regimenTypeStart = "Adult.1st.Line";
                } else if (regimenTypeStart.contains("ART Second Line Adult")) {
                    regimenTypeStart = "Adult.2nd.Line";
                } else if (regimenTypeStart.contains("ART First Line Children")) {
                    regimenTypeStart = "Peds.1st.Line";
                } else if (regimenTypeStart.contains("ART Second Line Children")) {
                    regimenTypeStart = "Peds.2nd.Line";
                } else if (regimenTypeStart.contains("Third Line")) {
                    if (age < 5) {
                        regimenTypeStart = "Peds.3rd.Line";
                    } else {
                        regimenTypeStart = "Adult.3rd.Line";
                    }
                } else {
                    regimenTypeStart = "";
                }
                entry.setRegimenLineAtStart(regimenTypeStart);
                if (!StringUtils.isEmpty(regimenTypeStart)) {
                    String regimenStart = RegimenIntrospector.resolveRegimen(StringUtils.trimToEmpty(rs.getString("regimen")));
                    entry.setRegimenAtStart(regimenStart);
                }
            }, patientEntry.getPatientId(), reportingDateEnd);
            //Check Date Last Regimen Switch
            jdbcTemplate.query("SELECT date_visit FROM (SELECT date_visit, cast(jsonb_extract_path_text(l,'regimen_type_id') as integer) curr," +
                    " LAG(cast(jsonb_extract_path_text(l,'regimen_type_id') as integer)) " +
                    "OVER(ORDER BY date_visit) prev FROM pharmacy p, jsonb_array_elements(lines) with ordinality a(l) WHERE " +
                    "p.patient_id = ? AND date_visit <= ? AND p.archived = false AND cast(jsonb_extract_path_text(l,'regimen_type_id') as integer) " +
                    "IN (1, 2, 3, 4, 14)) l WHERE ((curr - prev = 1) " +
                    "OR (curr - prev = 10) OR (curr - prev = 12)) AND curr IN (2, 4, 14) ORDER BY 1 DESC LIMIT 1", rs -> {
                entry.setRegimenSwitchDate(rs.getDate("date_visit"));
            }, patientEntry.getPatientId(), reportingDateEnd);
            //Pregnancy status as at the last clinic visit before the reporting date
            String gender = entry.getSex();
            if (StringUtils.equals(gender, "Female")) {
                entry.setPregnancyStatus("Not Pregnant");
                query = "SELECT pregnant, breastfeeding FROM clinic WHERE patient_id = ? AND date_visit "
                        + "BETWEEN (? + INTERVAL '-12 MONTH') "
                        + "AND ? and archived ORDER BY date_visit DESC";
                jdbcTemplate.query(query, rs -> {
                    while (rs.next()) {
                        if (rs.getBoolean("pregnant")) {
                            entry.setPregnancyStatus("Pregnant");
                        }
                        if (rs.getBoolean("breastfeeding")) {
                            entry.setPregnancyStatus("Breastfeeding");
                        }
                    }
                    return null;
                }, patientEntry.getPatientId(), reportingDateBegin, reportingDateEnd);
            }
            //Last viral load test value on or before the end of reporting date
            LocalDate dateStarted = patientEntry.getDateStarted();
            query = "SELECT jsonb_extract_path_text(l,'result') result, date_result_received, date_sample_collected, " +
                    "jsonb_extract_path_text(l,'indication') indication FROM laboratory, " +
                    "jsonb_array_elements(lines) with ordinality a(l) WHERE patient_id = ? AND date_sample_collected <= ? AND " +
                    "cast(jsonb_extract_path_text(l,'lab_test_id') as integer) = 16 and archived = false ORDER BY date_sample_collected DESC LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                try {
                    entry.setCurrentViralLoad(Double.parseDouble(StringUtils.trimToEmpty(rs.getString("result"))));
                } catch (Exception ignored) {
                }
                entry.setCurrentViralLoadDate(rs.getDate("date_result_received"));
                entry.setViralLoadIndication(StringUtils.trimToEmpty(rs.getString("indication")));
                entry.setViralLoadSampleCollectedDate(rs.getDate("date_sample_collected"));
                if (entry.getCurrentViralLoadDate() == null) {
                    entry.setCurrentViralLoad(null);
                }
            }, patientEntry.getPatientId(), reportingDateEnd);

            //Devolvement Information
            query = "SELECT * FROM devolve WHERE patient_id = ? and archived = false and date_devolved <= ? ORDER BY date_devolved DESC LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                if (rs.next()) {
                    entry.setDmocServiceCommenceDate(rs.getDate("date_devolved"));
                    entry.setDateReturnedToFacility(rs.getDate("date_returned_to_facility"));
                    String type = rs.getString("dmoc_type");
                    type = StringUtils.equals(type, "F_CARG") ? "F-CARG" : StringUtils.equals(type, "FAST_TRACK") ? "Fast Track" :
                            StringUtils.equals(type, "S_CARG") ? "S-CARG" : StringUtils.equals(type, "ARC") ? "Adolescent Refill Club" : type;
                    entry.setDmocType(type);
                    entry.setReceivingDmocService("Yes");

                    if (entry.getDateReturnedToFacility() != null) {
                        entry.setDmocType("");
                        entry.setReceivingDmocService("No");
                        entry.setDmocServiceCommenceDate(null);
                    }

                } else {
                    if (dateStarted != null && dateStarted.plusMonths(12).isBefore(reportingDateBegin)) {
                        entry.setReceivingDmocService("N/A");
                    } else {
                        entry.setReceivingDmocService("No");
                    }
                }
                return null;
            }, patientEntry.getPatientId(), reportingDateEnd);
            //Viral Load Monitoring/Enhanced Adherence Counseling
            boolean unsuppressed = true;
            try {
                Double vl = entry.getCurrentViralLoad();
                if (vl != null && vl < 1000) {
                    unsuppressed = false;
                }
            } catch (Exception ignored) {
            }
            if (unsuppressed) {
                entry.setRepeatViralLoadPostEacCollected("No");
                query = "SELECT * FROM eac WHERE patient_id = ? and archived = false ORDER BY date_eac1 DESC LIMIT 1";
                jdbcTemplate.query(query, rs -> {
                    entry.setEacCommenced("Yes");
                    entry.setEacCommencementDate(rs.getDate("date_eac1"));
                    if (rs.getDate("date_eac1") != null) {
                        entry.setEacSessions(1);
                        jdbcTemplate.query("SELECT description FROM pharmacy p, jsonb_array_elements(lines) with ordinality a(l)  " +
                                "JOIN regimen r ON r.id = cast(jsonb_extract_path_text(l,'regimen_id') as integer) WHERE p.patient_id = ? AND date_visit <= ? AND " +
                                "r.regimen_type_id  = 15 AND p.archived ORDER BY date_visit DESC LIMIT 1", rse1 -> {
                            entry.setIptType(rse1.getString(1));
                        }, patientEntry.getPatientId(), rs.getDate("date_eac1"));
                    }
                    if (rs.getDate("date_eac2") != null) {
                        entry.setEacSessions(2);
                    }
                    if (rs.getDate("date_eac3") != null) {
                        entry.setEacSessions(3);
                    }
                    Date dateSampleCollected = rs.getDate("date_sample_collected");
                    if (dateSampleCollected != null) {
                        entry.setRepeatViralLoadPostEacCollected("Yes");

                        jdbcTemplate.query("SELECT date_result_received FROM laboratory , jsonb_array_elements(lines) with ordinality a(l) " +
                                "WHERE patient_id = ? AND cast(jsonb_extract_path_text(l,'lab_test_id') as integer ) = 16 AND date_result_received " +
                                "BETWEEN ? AND ? AND archived = false ORDER BY date_result_received LIMIT 1", rse -> {
                            entry.setRepeatViralLoadCollectionDate(rse.getDate("date_result_received"));
                        }, patientEntry.getPatientId(), dateSampleCollected, reportingDateEnd);
                    }
                }, patientEntry.getPatientId());
            }
            entry.setChronicCareScreening("No");
            jdbcTemplate.query("select date_visit, tb_treatment, hypertensive, tb_referred, diabetic, bp_referred, dm_referred " +
                    "from chronic_care where patient_id = ? and archived = false", rs -> {
                entry.setChronicCareScreening("Yes");
                boolean tb = rs.getBoolean("tb_treatment");
                boolean hypertensive = rs.getBoolean("hypertensive");
                boolean diabetic = rs.getBoolean("diabetic");
                tb = false;
                String coMorbidities = hypertensive && diabetic && tb ? "TB/Hypertensive/Diabetic" :
                        hypertensive && tb ? "TB/Hypertensive" : tb && diabetic ? "TB/Diabetic" : hypertensive && diabetic ?
                                "Hypertensive/Diabetic" : tb ? "TB" : diabetic ? "Diabetic" : hypertensive ? "Hypertensive" : "";
                entry.setComorbidities(coMorbidities);
                String referred = "No";
                Boolean bp = rs.getBoolean("bp_referred");
                Boolean dm = rs.getBoolean("dm_referred");
                if (bp || dm || rs.getBoolean("tb_referred")) {
                    referred = "Yes";
                }
                entry.setReferred(referred);
                entry.setDateOfScreening(rs.getDate("date_visit"));
            }, patientEntry.getPatientId());
            jdbcTemplate.query("select name from case_manager where id = (select case_manager_id " +
                    "from patient where id = ?)", rs -> {
                entry.setCaseManager(rs.getString(1));
            }, patientEntry.getPatientId());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return entry;
    }

    private CellStyle getCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        style.setFillPattern(FillPatternType.FINE_DOTS);
        style.setFont(font);
        return style;
    }

    public List<Map<String, Object>> listFacilities() {
        return jdbcTemplate.query("select distinct name, facility_id from patient join facility f ON f.id = facility_id order by 1", rs -> {
            List<Map<String, Object>> facilities = new ArrayList<>();
            while (rs.next()) {
                Map<String, Object> facility = new HashMap<>();
                facility.put("name", rs.getString(1));
                facility.put("id", rs.getString(2));
                facilities.add(facility);
            }
            return facilities;
        });
    }

    private String getFacility(Long id) {
        return jdbcTemplate.queryForObject("select name from facility where id = ?", String.class, id);
    }

    @SneakyThrows
    public Set<String> listFiles() {
        String folder = applicationProperties.getTempDir() + BASE_DIR;
        return listFilesUsingDirectoryStream(folder);
    }

    @SneakyThrows
    public ByteArrayOutputStream downloadFile(String file) {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        String folder = applicationProperties.getTempDir() + BASE_DIR;
        Optional<String> fileToDownload = listFilesUsingDirectoryStream(folder).stream()
                .filter(f -> f.equals(file))
                .findFirst();
        fileToDownload.ifPresent(s -> {
            try (InputStream is = new FileInputStream(folder + s)) {
                IOUtils.copy(is, baos);
            } catch (IOException ignored) {
            }
        });
        return baos;
    }

    private Set<String> listFilesUsingDirectoryStream(String dir) throws IOException {
        Set<String> fileList = new HashSet<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(dir))) {
            for (Path path : stream) {
                if (!Files.isDirectory(path)) {
                    fileList.add(path.getFileName()
                            .toString());
                }
            }
        }
        return fileList;
    }

    public Date convertToDateViaInstant(LocalDate dateToConvert) {
        return java.util.Date.from(dateToConvert.atStartOfDay()
                .atZone(ZoneId.systemDefault())
                .toInstant());
    }

    @SneakyThrows
    @PostConstruct
    public void init() {
        String folder = applicationProperties.getTempDir() + BASE_DIR;
        new File(folder).mkdirs();
        FileUtils.cleanDirectory(new File(folder));
    }

    private void buildSheet(Sheet sheet, CellStyle dateStyle, List<RadetEntry> entries) {
        AtomicInteger atomicInteger = new AtomicInteger();
        entries.forEach(entry -> {
            int cellNum = 0;
            Row row = sheet.createRow(atomicInteger.incrementAndGet());
            Cell cell = row.createCell(cellNum++);
            cell.setCellValue(atomicInteger.get());
            cell = row.createCell(cellNum++);
            if (entry.getUniqueId() != null) {
                cell.setCellValue(entry.getUniqueId());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getHospitalNum());
            cell = row.createCell(cellNum++);
            if (entry.getHouseholdUniqueNo() != null) {
                cell.setCellValue(entry.getHouseholdUniqueNo());
            }
            cell = row.createCell(cellNum++);
            if (entry.getReceivedOvcService() != null) {
                cell.setCellValue(entry.getReceivedOvcService());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getSex());
            cell = row.createCell(cellNum++);
            if (entry.getWeight() != null) {
                cell.setCellValue(entry.getWeight());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDob());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getArtStartDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getLastPickupDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            if (entry.getMonthsOfRefill() != null) {
                cell.setCellValue(entry.getMonthsOfRefill());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getIptInLast2Years());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getIptStartDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getIptCompletionsDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRegimenLineAtStart());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRegimenAtStart());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentRegimenLine());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentArtRegimen());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRegimenSwitchDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getPregnancyStatus());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getFullDisclosure());
            cell = row.createCell(cellNum++);
            if (entry.getEnrolledOnOTZ() != null) {
                cell.setCellValue(entry.getEnrolledOnOTZ());
            }
            cell = row.createCell(cellNum++);
            if (entry.getNumberOfOTZMeetings() != null) {
                cell.setCellValue(entry.getNumberOfOTZMeetings());
            }
            cell = row.createCell(cellNum++);
            if (entry.getOtzModulesCompleted() != null && entry.getOtzModulesCompleted() != 0) {
                cell.setCellValue(entry.getOtzModulesCompleted());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDateEnrolledOnOTZ());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getViralLoadSampleCollectedDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            try {
                cell.setCellValue(entry.getCurrentViralLoad());
            } catch (Exception ignored) {
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentViralLoadDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getViralLoadIndication());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentArtStatus());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentArtStatusDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCauseOfDeath());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getTransferOutFacility());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getArtEnrollmentSetting());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getReceivingDmocService());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDmocServiceCommenceDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDmocType());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDateReturnedToFacility());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getEacCommenced());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getEacCommencementDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            if (entry.getEacSessions() != null) {
                cell.setCellValue(entry.getEacSessions());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRepeatViralLoadPostEacCollected());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRepeatViralLoadCollectionDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRepeatViralLoadReceivedDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getIptType());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getChronicCareScreening());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getComorbidities());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getReferred());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCancerScreeningDate());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getScreenedForCervicalCancer());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCervicalCancerScreeningType());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCervicalCancerScreeningMethod());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getResultOfCervicalCancerScreening());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getPrecancerousScreeningMethods());
            cell = row.createCell(cellNum);
            cell.setCellValue(entry.getCaseManager());
        });
    }

    private void writeCSV() {
    }
}
