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
import java.util.*;

public class JsonToXls2 {
    private ObjectMapper objectMapper = new ObjectMapper();

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
                List<Map<String, String>> maps = new ArrayList<>();
                for (JsonNode n : node) {
                    Map<String, String> nodeValue = getNodeValue(n.iterator());
                    maps.add(nodeValue);
                }
//                    String sheetName = iterator.next();
//                    Sheet sheet = workbook.createSheet(sheetName);
//                    JsonNode sheetData = node.get(0).get(sheetName);
//                        createJsonObject(sheetData);
//                    ArrayList<String> headers = new ArrayList<>();
//                    CellStyle cellStyle = getCellStyle(workbook);
//
//                    Row header = sheet.createRow(0);
//                    if (sheetData.size() > 0) {
//                        Iterator<Map.Entry<String, JsonNode>> it = sheetData.get(0).fields();
//                        int headerIdx = 0;
//                        while (it.hasNext()) {
//                            String headerName = it.next().getKey();
//                            headers.add(headerName);
//                            Cell cell = header.createCell(headerIdx++);
//                            cell.setCellValue(headerName);
//                            cell.setCellStyle(cellStyle);
//                        }
//                    } else {
//                        int i = 0;
//                        headers.add(sheetName);
//                        Cell cell = header.createCell(i);
//                        cell.setCellValue(sheetName);
//                        cell.setCellStyle(cellStyle);
//                        Row row = sheet.createRow(i+1);
//                        row.createCell(i).setCellValue(sheetData.asText());
//
//                    }
//                    for (int i = 1; i < sheetData.size(); i++) {
//                        ObjectNode rowData = (ObjectNode) sheetData.get(i);
//                        Row row = sheet.createRow(i);
//                        for (int j = 0; j < headers.size(); j++) {
//                            StringBuilder value = new StringBuilder();
//                            if (rowData.get(headers.get(j)) instanceof ArrayNode){
//                                for (JsonNode arr: rowData.get(headers.get(j)))
//                                    value.append(arr.asText() + "\n");
//                            }
//                            else
//                                value = new StringBuilder(rowData.get(headers.get(j)).asText());
//                            row.createCell(j).setCellValue(value.toString());
//                        }
//                    }
//
//                    for (int i = 0; i < headers.size(); i++) {
//                        sheet.autoSizeColumn(i);
//                    }
//                }

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

    private static Map<String, String> getNodeValue(Iterator<JsonNode> iterator) {
        Boolean success = false;
        ArrayList<String> arrayList;
        Map<String, String> stringMap = new LinkedHashMap<>();
        while (iterator.hasNext()) {
            JsonNode next = iterator.next();
            if (next instanceof BooleanNode) {
                success = next.asBoolean();
            }
            if (next instanceof ArrayNode) {
                getNodeValue(next.iterator());
            }
            if (next instanceof ObjectNode){
                for (Iterator<String> it = next.fieldNames(); it.hasNext(); ) {
                    String field = it.next();
                    String node = next.get(field).asText();
                    stringMap.put(field, node);
                }
            }
        }
        return stringMap;
    }


//            map.put();
//            if (next.)
//                getNodeValue(next.iterator());
//
//            if (next instanceof ArrayNode){
//                Map<JsonNode, JsonNode> map = new LinkedHashMap<>();
//                map.put(next.get(0),next.get(1));
//                map.get(0);
//            }


    private static CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private static void createJsonObject(JsonNode sheetData) {
        JsonObject jsonObject = new JsonObject();
        if (sheetData instanceof BooleanNode)
            jsonObject.setSuccess((sheetData.asBoolean()));
        else {
            Iterator<JsonNode> iterator = sheetData.iterator();
            while (iterator.hasNext()) {
//                Iterator<LinkedHashMap<String, String>> iterator1 = iterator.next().fieldNames();
//                String next = iterator1.next();
            }
        }
    }
}
