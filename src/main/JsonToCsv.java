import lombok.SneakyThrows;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import com.github.opendevl.JFlat;


public class JsonToCsv {
    @SneakyThrows
    public static void main(String[] args) {
        String string = new String(Files.readAllBytes(Paths.get("src/main/resources/attendees.json")));
        JFlat flat = new JFlat(string);
        List<Object[]> jsonAsSheet = flat.json2Sheet().getJsonAsSheet();
        flat.write2csv("my.csv");
    }
}
