package org.fhi360.lamis.modules.radet.util;

import lombok.Data;

import java.time.LocalDate;

@Data
public class PatientEntry {
    Long patientId;
    String uniqueId;
    String hospitalNum;
    String sex;
    float weight;
    LocalDate dob;
    LocalDate dateStarted;
    String statusAtRegistration;
    int age;
    String artEnrollmentSetting;
    String householdUniqueNo;
    String servicesProvided;
}
