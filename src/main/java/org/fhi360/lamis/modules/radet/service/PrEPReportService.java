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
import org.fhi360.lamis.modules.radet.service.vm.PrepEntry;
import org.fhi360.lamis.modules.radet.util.RegimenIntrospector;
import org.lamisplus.modules.base.config.ApplicationProperties;
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
public class PrEPReportService {
    private final JdbcTemplate jdbcTemplate;
    private final ApplicationProperties applicationProperties;
    private final SimpMessageSendingOperations messagingTemplate;
    private final static String BASE_DIR = "prep/";

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
        cell.setCellValue("State");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("LGA");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Facility ID");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Facility Name");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Patient ID");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("PrEP ID");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Surname");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Other Names");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of Birth (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Age");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Sex");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Marital Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Education");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Occupation");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("State of Residence");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("LGA of Residence");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Address");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Phone");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Pregnancy Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Indication for PrEP");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Systolic BP");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Diastolic BP");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Weight (Kg)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Height (cm)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("HIV Status at PrEP Initiation");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Urinalysis");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Creatinine Clearance (mL/min)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Hepatitis B Screening");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Baseline Hepatitis C Screening");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Date of PrEP Initiation (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Initiation Setting");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Regimen at PrEP Initiation");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Systolic BP");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Diastolic BP");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Weight (kg)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Height (cm)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Regimen");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Last Refill Date (yyyy-mm-dd)");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Months of PrEP Refill");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current HIV Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Linked to ART");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Unique ID");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum++);
        cell.setCellValue("Current Status");
        cell.setCellStyle(style);
        cell = row.createCell(cellNum);
        cell.setCellValue("Reasons for discontinuation / Stopped");
        cell.setCellStyle(style);

        //Create a date format for date columns
        CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));
        //style.setAlignment(CellStyle.ALIGN_RIGHT);
        File file = new File(applicationProperties.getTempDir() + BASE_DIR + getFacility(facilityId).trim());
        FileUtils.deleteQuietly(file);

        AtomicInteger atomicInteger = new AtomicInteger();
        AtomicInteger sn = new AtomicInteger();
        List<PrepEntry> prepEntries = new ArrayList<>();
        List<PrepEntry> finalEntries = new ArrayList<>();
        String query = "" +
                "SELECT DISTINCT pl.uuid, pl.id, pl.extra->'prep'->>'prepId' prep_id, pl.facility_id, f.name facility, " +
                "   hospital_num, unique_id, surname, other_names, marital_status, education, occupation, rs.name r_state," +
                "   rl.name r_lga, s.name state, l.name lga, address, phone, enrollment_setting, gender, date_birth, " +
                "   DATEDIFF('YEAR', date_birth, CURRENT_DATE) age, pl.extra->'prep'->>'onDemandIndication' on_demand, " +
                "   pl.extra->'prep'->>'indicationForPrep' indication FROM patient pl JOIN clinic c ON pl.id = c.patient_id " +
                "   JOIN facility f ON f.id = pl.facility_id JOIN state s ON s.id = f.state_id JOIN lga l ON l.id = f.lga_id " +
                "   LEFT JOIN lga rl ON rl.id = pl.lga_id LEFT JOIN state rs ON rs.id = rl.state_id WHERE pl.facility_id = ? " +
                "   AND date_visit BETWEEN ? AND ? and pl.archived = false and c.archived = false AND pl.extra->'prep'->>'registered' = 'true' ";
        jdbcTemplate.query(query, rs -> {
            while (rs.next()) {
                PrepEntry entry = new PrepEntry();
                Long patientId = rs.getLong("id");
                String uniqueId = StringUtils.trimToEmpty(rs.getString("unique_id"));
                String hospitalNum = rs.getString("hospital_num");
                String gender = rs.getString("gender");
                String prepID = rs.getString("prep_id");
                Long facilityID = rs.getLong("facility_id");
                String facility = rs.getString("facility");
                String state = rs.getString("state");
                String lga = rs.getString("lga");
                String enrollmentSetting = StringUtils.trimToEmpty(rs.getString("enrollment_setting"));
                Date dateBirth = rs.getDate("date_birth");
                int age = rs.getInt("age");
                String indication = rs.getString("indication");
                entry.setIndicationForPrep(indication);
                String onDemand = rs.getString("on_demand");
                if (!StringUtils.isEmpty(onDemand)) {
                    entry.setIndicationForPrep(onDemand);
                }
                entry.setPatientId(patientId);
                entry.setUniqueId(uniqueId);
                entry.setHospitalNum(hospitalNum);
                entry.setInitiationSetting(enrollmentSetting);
                entry.setDob(dateBirth);
                entry.setAge(age);
                entry.setPrepId(prepID);
                entry.setFacilityId(facilityID);
                entry.setFacility(facility);
                entry.setState(state);
                entry.setLga(lga);
                entry.setPuuid(rs.getString("uuid"));
                entry.setLgaOfResidence(rs.getString("r_lga"));
                entry.setStateOfResidence(rs.getString("r_state"));
                entry.setMaritalStatus(rs.getString("marital_status"));
                entry.setEducation(rs.getString("education"));
                entry.setOccupation(rs.getString("occupation"));
                entry.setAddress(rs.getString("address"));
                entry.setPhone(rs.getString("phone"));
                entry.setSurname(rs.getString("surname"));
                entry.setOtherNames(rs.getString("other_names"));
                entry.setSn(sn.incrementAndGet());
                entry.setPatientId(patientId);
                entry.setUniqueId(uniqueId);
                entry.setHospitalNum(hospitalNum);
                entry.setSex(StringUtils.equals(gender, "MALE") ? "Male" : "Female");

                prepEntries.add(entry);
            }
            return null;
        }, facilityId, cohortBegin, cohortEnd);
        ExecutorService executorService = Executors.newFixedThreadPool(3);
        LocalDate finalReportingDateEnd = reportingDateEnd;
        prepEntries.forEach(patientEntry -> executorService.submit(() ->
                finalEntries.add(updateEntry(atomicInteger, patientEntry, reportingDateBegin, finalReportingDateEnd, prepEntries.size()))));
        executorService.shutdown();
        while (!executorService.isTerminated()) {

        }
        messagingTemplate.convertAndSend("/topic/prep/status", "Building sheet...");
        buildSheet(sheet, dateStyle, finalEntries);
        messagingTemplate.convertAndSend("/topic/prep/status", "Writing out file...");
        try {
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private PrepEntry updateEntry(AtomicInteger atomicInteger, PrepEntry entry,
                                  LocalDate reportingDateBegin, LocalDate reportingDateEnd, int total) {
        messagingTemplate.convertAndSend("/topic/prep/status", String.format("Analysing patient %s of %s",
                atomicInteger.incrementAndGet(), total));
        try {
            // pickup before the reporting date
            String query = "select date_visit, cast(jsonb_extract_path_text(l,'duration') as integer) duration, date_visit + " +
                    "cast(jsonb_extract_path_text(l,'duration') as integer) + INTERVAL '28 DAYS' < ? as ltfu, " +
                    "date(date_visit + cast(jsonb_extract_path_text(l,'duration') as integer) + INTERVAL '29 DAYS') ltfu_date, r.description regimen, " +
                    "t.description as regimen_type from pharmacy p, jsonb_array_elements(lines) with ordinality a(l) " +
                    "join regimen r on r.id = cast(jsonb_extract_path_text(l,'regimen_id') as integer)  " +
                    "join regimen_type t on t.id = r.regimen_type_id where t.id in (1,2,3,4,14) and p.patient_id = ? and " +
                    "date_visit <= ? and p.archived = false ORDER BY p.date_visit DESC, duration DESC LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                Date dlr = rs.getDate("date_visit");
                entry.setDateOfLastRefill(dlr);
                int duration = rs.getInt("duration");
                int monthRefill = duration / 30;
                if (monthRefill <= 0) {
                    monthRefill = 1;
                }
                entry.setMonthsOfRefill(monthRefill);
                String regimen = RegimenIntrospector.resolveRegimen(StringUtils.trimToEmpty(rs.getString("regimen")));
                entry.setRegimen(StringUtils.isNoneBlank(regimen) ? regimen : rs.getString("regimen"));

            }, reportingDateEnd, entry.getPatientId(), reportingDateEnd);
            //Baseline values
            query = "SELECT bp, body_weight, height, c.extra->'prep'->>'hivTestResult' hiv_status, c.extra->'prep'->>'urinalysis' urinalysis," +
                    "c.extra->'prep'->>'hepatitisB' hepatitis_b, c.extra->'prep'->>'hepatitisC' hepatitis_c, c.extra->'prep'->>'creatinineClearance' " +
                    "creatinine_clearance, date_visit initiation_date, r.description regimen FROM clinic c JOIN regimen r ON r.id = regimen_id " +
                    "WHERE patient_id = ? AND date_visit <= ? AND commence = true " +
                    "AND c.archived = false ORDER BY date_visit LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                String regimenStart = RegimenIntrospector.resolveRegimen(StringUtils.trimToEmpty(rs.getString("regimen")));
                entry.setRegimenAtInitiation(StringUtils.isNoneBlank(regimenStart) ? regimenStart : rs.getString("regimen"));
                String bp = rs.getString("bp");
                if (!StringUtils.isEmpty(bp)) {
                    String[] parts = bp.split("/");
                    if (parts.length >= 2) {
                        entry.setBaselineSystolic(parts[0]);
                        entry.setBaselineDiastolic(parts[1]);
                    }
                }
                double weight = rs.getDouble("body_weight");
                if (weight != 0) {
                    entry.setBaselineWeight(weight);
                }
                int height = rs.getInt("height");
                if (height != 0) {
                    entry.setBaselineHeight(height);
                }
                entry.setHivStatusAtInitiation(rs.getString("hiv_status"));
                entry.setBaselineUrinalysis(rs.getString("urinalysis"));
                entry.setBaselineCreatinineClearance(rs.getString("creatinine_clearance"));
                entry.setBaselineHepatitisB(rs.getString("hepatitis_b"));
                entry.setBaselineHepatitisC(rs.getString("hepatitis_c"));
                entry.setDateOfPrepInitiation(rs.getDate("initiation_date"));
                entry.setRegimenAtInitiation(rs.getString("regimen"));
            }, entry.getPatientId(), reportingDateEnd);
            //Current values
            query = "SELECT bp, body_weight, height, extra->'prep'->>'hivTestResult' hiv_status, extra->'prep'->>'urinalysis' urinalysis," +
                    "extra->'prep'->>'hepatitisB' hepatitis_b, extra->'prep'->>'hepatitisC' hepatitis_c,extra->'prep'->>'creatinineClearance' " +
                    "creatinine_clearance, date_visit initiation_date FROM clinic WHERE patient_id = ? AND date_visit <= ? " +
                    "and archived = false ORDER BY date_visit DESC LIMIT 1";
            jdbcTemplate.query(query, rs -> {
                String bp = rs.getString("bp");
                if (!StringUtils.isEmpty(bp)) {
                    String[] parts = bp.split("/");
                    if (parts.length >= 2) {
                        entry.setSystolic(parts[0]);
                        entry.setDiastolic(parts[1]);
                    }
                }
                double weight = rs.getDouble("body_weight");
                if (weight != 0) {
                    entry.setWeight(weight);
                }
                int height = rs.getInt("height");
                if (height != 0) {
                    entry.setHeight(height);
                }
                entry.setCurrentHivStatus(rs.getString("hiv_status"));
            }, entry.getPatientId(), reportingDateEnd);

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
                }, entry.getPatientId(), reportingDateBegin, reportingDateEnd);
            }

            //Linked to ART
            if (!StringUtils.equals(entry.getCurrentHivStatus(), "Positive")) {
                entry.setLinkedToArt(null);
                entry.setUniqueId(null);
            } else {
                entry.setLinkedToArt("No");
            }
            jdbcTemplate.query("select id from pharmacy, jsonb_array_elements(lines) with ordinality a(l) where " +
                    "cast(jsonb_extract_path_text(l,'regimen_type_id') as integer) in (1, 2, 3, 4, 14) and patient_id = ? " +
                    "and date_visit <= ? and archived = false", rs -> {
                entry.setLinkedToArt("Yes");
            }, entry.getPatientId(), reportingDateEnd);
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

    private void buildSheet(Sheet sheet, CellStyle dateStyle, List<PrepEntry> entries) {
        AtomicInteger atomicInteger = new AtomicInteger();
        entries.forEach(entry -> {
            int cellNum = 0;
            Row row = sheet.createRow(atomicInteger.incrementAndGet());
            Cell cell = row.createCell(cellNum++);
            cell.setCellValue(atomicInteger.get());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getState());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getLga());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getFacilityId());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getFacility());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getPuuid());
            cell = row.createCell(cellNum++);
            if (entry.getPrepId() != null) {
                cell.setCellValue(entry.getPrepId());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getSurname());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getOtherNames());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDob());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getAge());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getSex());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getMaritalStatus());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getEducation());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getOccupation());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getStateOfResidence());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getLgaOfResidence());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getAddress());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getPhone());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getPregnancyStatus());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getIndicationForPrep());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getBaselineSystolic());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getBaselineDiastolic());
            cell = row.createCell(cellNum++);
            if (entry.getBaselineWeight() != null) {
                cell.setCellValue(entry.getBaselineWeight());
            }
            cell = row.createCell(cellNum++);
            if (entry.getBaselineHeight() != null) {
                cell.setCellValue(entry.getBaselineHeight());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getHivStatusAtInitiation());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getBaselineUrinalysis());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getBaselineCreatinineClearance());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getBaselineHepatitisB());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getBaselineHepatitisC());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDateOfPrepInitiation());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getInitiationSetting());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRegimenAtInitiation());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getSystolic());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDiastolic());
            cell = row.createCell(cellNum++);
            if (entry.getWeight() != null) {
                cell.setCellValue(entry.getWeight());
            }
            cell = row.createCell(cellNum++);
            if (entry.getHeight() != null) {
                cell.setCellValue(entry.getHeight());
            }
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getRegimen());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getDateOfLastRefill());
            cell.setCellStyle(dateStyle);
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getMonthsOfRefill());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentHivStatus());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getLinkedToArt());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getUniqueId());
            cell = row.createCell(cellNum++);
            cell.setCellValue(entry.getCurrentStatus());
            cell = row.createCell(cellNum);
            cell.setCellValue(entry.getReasonsForDiscontinuation());
        });
    }
}
