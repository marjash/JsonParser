import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.BooleanNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.SneakyThrows;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

public class JsonToXls {
    private ObjectMapper objectMapper = new ObjectMapper();
    private Sheet sheet = null;
    private JsonNode sheetData = null;

    @SneakyThrows
    public File JsonFileToXlsFile(File json, String xls) {
        try {
            if (!json.getName().endsWith(".json"))
                throw new IllegalArgumentException("file must bi json format");
            else {
                Workbook workbook = null;
                if (xls.equals(".xls"))
                    workbook = new HSSFWorkbook();
                else
                    throw new IllegalArgumentException("file must be xls format");
                JsonNode node = objectMapper.readTree(json);
                ArrayList<String> headers = getHeaders(workbook, node);
                Iterator<JsonNode> iterator = node.iterator();
                int count = 1;
                while (iterator.hasNext()) {
                    createCell(headers, iterator.next(), count);
                    count += sheetData.size();
                }
                String fileName = json.getName();
                fileName = fileName.substring(0, fileName.lastIndexOf(".json")) + xls;
                File targetFile = new File(json.getParent(), fileName);

                FileOutputStream fileOutputStream = new FileOutputStream(targetFile);
                workbook.write(fileOutputStream);

                workbook.close();
                fileOutputStream.close();
                return targetFile;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private void createCell(ArrayList<String> headers, JsonNode next, int count) {
        JsonNode node = next.get("data");
        for (int i = 0; i < node.size(); i++) {
            ObjectNode rowData = (ObjectNode) node.get(i);
            Row row = sheet.createRow(count);
            for (int j = 0; j < headers.size(); j++) {
                StringBuilder value = new StringBuilder();
                if (rowData.get(headers.get(j)) instanceof ArrayNode) {
                    for (JsonNode arr : rowData.get(headers.get(j)))
                        value.append(arr.asText() + "\n");
                } else
                    value = new StringBuilder(rowData.get(headers.get(j)).asText());
                row.createCell(j).setCellValue(value.toString());
            }
            count++;
        }
        for (int i = 0; i < headers.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private ArrayList<String> getHeaders(Workbook workbook, JsonNode node) {
        Iterator<String> iterator = node.get(0).fieldNames();
        ArrayList<String> headers = new ArrayList<>();
        while (iterator.hasNext()) {
            String sheetName = iterator.next();
            sheet = workbook.createSheet(sheetName);
            sheetData = node.get(0).get(sheetName);
            CellStyle cellStyle = getCellStyle(workbook);

            Row header = sheet.createRow(0);
            if (sheetData.size() > 0) {
                Iterator<Map.Entry<String, JsonNode>> it = sheetData.get(0).fields();
                int headerIdx = 0;
                while (it.hasNext()) {
                    String headerName = it.next().getKey();
                    headers.add(headerName);
                    Cell cell = header.createCell(headerIdx++);
                    cell.setCellValue(headerName);
                    cell.setCellStyle(cellStyle);
                }
            } else {
                int i = 0;
//                headers.add(sheetName);
                Cell cell = header.createCell(i);
                cell.setCellValue(sheetName);
                cell.setCellStyle(cellStyle);
                Row row = sheet.createRow(i + 1);
                row.createCell(i).setCellValue(sheetData.asText());
            }
        }
        return headers;
    }

    private static CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }

}
