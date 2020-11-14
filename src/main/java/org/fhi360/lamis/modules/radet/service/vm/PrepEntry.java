package org.fhi360.lamis.modules.radet.service.vm;

import lombok.Data;

import java.util.Date;

@Data
public class PrepEntry {
    int sn;
    Long patientId;
    String puuid;
    String prepId;
    String state;
    String lga;
    Long facilityId;
    String facility;
    String stateOfResidence;
    String lgaOfResidence;
    String surname;
    String otherNames;
    String uniqueId;
    String hospitalNum;
    String sex;
    Date dob;
    int age;
    String maritalStatus;
    String occupation;
    String education;
    String address;
    String phone;
    String initiationSetting;
    String pregnancyStatus;
    String indicationForPrep;
    String baselineSystolic;
    String baselineDiastolic;
    Double baselineWeight;
    Integer baselineHeight;
    String hivStatusAtInitiation;
    String baselineCreatinineClearance;
    String baselineUrinalysis;
    String baselineHepatitisB;
    String baselineHepatitisC;
    Date dateOfPrepInitiation;
    String regimenAtInitiation;
    String systolic;
    String diastolic;
    Double weight;
    Integer height;
    String regimen;
    Date dateOfLastRefill;
    int monthsOfRefill;
    String currentHivStatus;
    String linkedToArt;
    String currentStatus;
    String reasonsForDiscontinuation;
}
