import java.io.File;

public class Main {
    public static void main(String[] args) {
        JsonToXls jsonToXml = new JsonToXls();
        File json = new File("src/main/resources/attendees.json");
        File xls = jsonToXml.JsonFileToXlsFile(json, ".xls");
        System.out.println("Sucessfully converted JSON to Excel file at =" + xls.getAbsolutePath());

    }

}
