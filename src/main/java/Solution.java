import java.io.IOException;
import java.util.ArrayList;

public class Solution {
    public static void main(String[] args) {
        ArrayList<String> filePaths = new ArrayList<>();
        filePaths.add("/media/dmitry/диск/apachePOITest/Копище, ул. А. Микояна, 1.xlsx");
        filePaths.add("/media/dmitry/диск/apachePOITest/Копище, ул. А. Эрхарт, 1.xlsx");
        ExcelFileConverter excelFileConverter = new ExcelFileConverter();
        try {
            excelFileConverter.convert(filePaths);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
