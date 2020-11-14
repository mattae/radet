package org.fhi360.lamis.modules.radet.web;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.fhi360.lamis.modules.radet.service.PrEPReportService;
import org.fhi360.lamis.modules.radet.service.RadetConverterService;
import org.springframework.messaging.simp.SimpMessageSendingOperations;
import org.springframework.scheduling.annotation.Async;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.Collection;
import java.util.List;
import java.util.Map;

@Slf4j
@RestController
@RequestMapping("/api")
@RequiredArgsConstructor
public class RadetConverterResource {
    private final RadetConverterService radetConverterService;
    private final PrEPReportService prEPReportService;
    private final SimpMessageSendingOperations messagingTemplate;

    @GetMapping("/radet/convert")
    @Async
    public void runRadet(@RequestParam LocalDate cohortStart, @RequestParam LocalDate cohortEnd,
                         @RequestParam LocalDate reportingPeriod, @RequestParam List<Long> ids,
                         @RequestParam(defaultValue = "false") Boolean today) {
        messagingTemplate.convertAndSend("/topic/radet/status", "start");
        ids.forEach(id -> radetConverterService.convertExcel(id, cohortStart, cohortEnd, reportingPeriod, today));
        messagingTemplate.convertAndSend("/topic/radet/status", "end");
    }

    @GetMapping("/radet/list-facilities")
    public List<Map<String, Object>> listFacilities() {
        return radetConverterService.listFacilities();
    }

    @GetMapping("/radet/download/{file}")
    public void downloadRadetFile(@PathVariable String file, HttpServletResponse response) throws IOException {
        ByteArrayOutputStream baos = radetConverterService.downloadFile(file);
        writeStream(baos, response);
    }

    @GetMapping("/radet/list-files")
    public Collection<String> listRadetFiles() {
        return radetConverterService.listFiles();
    }

    @GetMapping("/prep/convert")
    @Async
    public void runPrep(@RequestParam LocalDate cohortStart, @RequestParam LocalDate cohortEnd,
                        @RequestParam LocalDate reportingPeriod, @RequestParam List<Long> ids,
                        @RequestParam(defaultValue = "false") Boolean today) {
        messagingTemplate.convertAndSend("/topic/prep/status", "start");
        ids.forEach(id -> prEPReportService.convertExcel(id, cohortStart, cohortEnd, reportingPeriod, today));
        messagingTemplate.convertAndSend("/topic/prep/status", "end");
    }

    @GetMapping("/prep/download/{file}")
    public void downloadPrepFile(@PathVariable String file, HttpServletResponse response) throws IOException {
        ByteArrayOutputStream baos = prEPReportService.downloadFile(file);
        writeStream(baos, response);
    }

    @GetMapping("/prep/list-files")
    public Collection<String> listPrepFiles() {
        return prEPReportService.listFiles();
    }

    private void writeStream(ByteArrayOutputStream baos, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Type", "application/octet-stream");
        response.setHeader("Content-Length", Integer.valueOf(baos.size()).toString());
        OutputStream outputStream = response.getOutputStream();
        outputStream.write(baos.toByteArray());
        outputStream.close();
        response.flushBuffer();
    }
}
