package com.example.excel.sample;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Controller
@Slf4j
@RequiredArgsConstructor
public class ExcelSampleController {

    private final ExcelSampleService excelSampleService;

    @GetMapping("/page/sample")
    public String excelSamplePage(Model model) {

        System.out.println("접근 테스트");
        return "sample/samplePage";
    }

    @GetMapping("/file/sample")
    public ResponseEntity<byte[]> downloadExcelSampleFile() throws IOException {

        Workbook workbook = excelSampleService.createExcelWorkbook();

        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "sample.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(out.toByteArray());

        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().build();
        }    }


}
