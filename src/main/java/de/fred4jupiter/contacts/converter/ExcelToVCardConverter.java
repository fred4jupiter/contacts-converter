package de.fred4jupiter.contacts.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;
import org.springframework.shell.standard.ShellOption;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@ShellComponent
public class ExcelToVCardConverter {

    @ShellMethod(key = "excel2vcard", value = "Convert Excel file with address data to vCard format")
    public void excel2vcard(
            @ShellOption(value = {"--input", "-i"}, help = "Path to Excel file") String inputFile,
            @ShellOption(value = {"--output", "-o"}, help = "Output directory", defaultValue = ".") String outputDir) {

        Path inputPath = Paths.get(inputFile);

        if (!Files.exists(inputPath)) {
            System.out.println("❌ Error: File '" + inputFile + "' not found");
            return;
        }

        String fileName = inputFile.toLowerCase();
        if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
            System.out.println("❌ Error: Please provide an Excel file (.xlsx or .xls)");
            return;
        }

        try {
            Path outPath = Paths.get(outputDir);
            Files.createDirectories(outPath);

            System.out.println("📖 Reading " + inputFile + "...");
            List<Map<String, String>> contacts = readExcelFile(inputPath);

            if (contacts.isEmpty()) {
                System.out.println("❌ No contacts found in file");
                return;
            }

            int successful = 0;
            int failed = 0;

            for (int idx = 0; idx < contacts.size(); idx++) {
                Map<String, String> contact = contacts.get(idx);
                String vcard = createVCard(contact);

                if (vcard != null) {
                    String name = contact.getOrDefault("name",
                            (contact.getOrDefault("firstname", "") + " " +
                                    contact.getOrDefault("lastname", "")).trim());

                    if (name.isEmpty()) {
                        name = "contact_" + (idx + 1);
                    }

                    String filename = name.replaceAll("[\\s/]", "_") + ".vcf";
                    Path vcardPath = outPath.resolve(filename);

                    try {
                        Files.write(vcardPath, vcard.getBytes(StandardCharsets.UTF_8));
                        System.out.println("✅ Created: " + filename);
                        successful++;
                    } catch (IOException e) {
                        System.out.println("❌ Failed to create " + filename + ": " + e.getMessage());
                        failed++;
                    }
                } else {
                    System.out.println("⏭️  Skipped row " + (idx + 1) + ": No name found");
                    failed++;
                }
            }

            System.out.println("\n✅ Conversion complete: " + successful + " vCard(s) created, " + failed + " failed");
            System.out.println("📁 Output directory: " + outPath.toAbsolutePath());

        } catch (Exception e) {
            System.out.println("❌ Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    @ShellMethod(key = "help-columns", value = "Show supported column names")
    public void helpColumns() {
        System.out.println("\n📋 Supported column names (case-insensitive):");
        System.out.println("   name, firstname, lastname, email, phone, mobile,");
        System.out.println("   street, city, state, zip, country, company, title, website\n");
    }

    private List<Map<String, String>> readExcelFile(Path filePath) throws IOException {
        List<Map<String, String>> contacts = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            if (headerRow == null) {
                System.out.println("❌ Error: No headers found in Excel file");
                return contacts;
            }

            List<String> headers = new ArrayList<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null && cell.getStringCellValue() != null) {
                    headers.add(normalizeHeader(cell.getStringCellValue()));
                } else {
                    headers.add("");
                }
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> contact = new HashMap<>();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null && !headers.get(j).isEmpty()) {
                        String value = getCellValue(cell).trim();
                        if (!value.isEmpty()) {
                            contact.put(headers.get(j), value);
                        }
                    }
                }

                if (!contact.isEmpty()) {
                    contacts.add(contact);
                }
            }
        }

        return contacts;
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    private String normalizeHeader(String header) {
        return header.trim().toLowerCase().replaceAll("\\s+", "");
    }

    private String createVCard(Map<String, String> data) {
        StringBuilder vcard = new StringBuilder();
        vcard.append("BEGIN:VCARD\n");
        vcard.append("VERSION:3.0\n");

        // Required: Full Name
        String name = data.get("name");
        if (name == null || name.isEmpty()) {
            String firstname = data.getOrDefault("firstname", "");
            String lastname = data.getOrDefault("lastname", "");
            name = (firstname + " " + lastname).trim();
        }

        if (name.isEmpty()) {
            return null; // Skip entries without a name
        }

        vcard.append("FN:").append(name).append("\n");

        // Name structure
        if (data.containsKey("lastname") || data.containsKey("firstname")) {
            String lastname = data.getOrDefault("lastname", "");
            String firstname = data.getOrDefault("firstname", "");
            vcard.append("N:").append(lastname).append(";").append(firstname).append(";;;\n");
        }

        // Phone
        if (data.containsKey("phone") && !data.get("phone").isEmpty()) {
            vcard.append("TEL;TYPE=VOICE:").append(data.get("phone")).append("\n");
        }
        if (data.containsKey("mobile") && !data.get("mobile").isEmpty()) {
            vcard.append("TEL;TYPE=CELL:").append(data.get("mobile")).append("\n");
        }

        // Email
        if (data.containsKey("email") && !data.get("email").isEmpty()) {
            vcard.append("EMAIL;TYPE=INTERNET:").append(data.get("email")).append("\n");
        }

        // Address
        String street = data.getOrDefault("street", "");
        String city = data.getOrDefault("city", "");
        String state = data.getOrDefault("state", "");
        String zip = data.getOrDefault("zip", "");
        String country = data.getOrDefault("country", "");

        if (!street.isEmpty() || !city.isEmpty() || !state.isEmpty() || !zip.isEmpty() || !country.isEmpty()) {
            vcard.append("ADR:;;").append(street).append(";").append(city).append(";")
                    .append(state).append(";").append(zip).append(";").append(country).append("\n");
        }

        // Organization
        if (data.containsKey("company") && !data.get("company").isEmpty()) {
            vcard.append("ORG:").append(data.get("company")).append("\n");
        }

        // Title
        if (data.containsKey("title") && !data.get("title").isEmpty()) {
            vcard.append("TITLE:").append(data.get("title")).append("\n");
        }

        // Website
        if (data.containsKey("website") && !data.get("website").isEmpty()) {
            vcard.append("URL:").append(data.get("website")).append("\n");
        }

        vcard.append("END:VCARD\n");
        return vcard.toString();
    }
}