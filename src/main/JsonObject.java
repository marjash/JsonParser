import com.fasterxml.jackson.databind.JsonNode;
import lombok.Data;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

@Data
public class JsonObject {
    private boolean success;
    private Data data;

    static class Data<K, V> extends LinkedHashMap<K, V> {
    }
}



