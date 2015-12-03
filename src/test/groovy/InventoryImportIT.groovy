import io.dbmaster.testng.BaseToolTestNGCase;
import static org.testng.Assert.assertTrue;

import org.testng.annotations.Parameters;
import org.testng.annotations.Test

public class InventoryImportIT extends BaseToolTestNGCase {

    @Test
    @Parameters("inventory-import.p_excel_file")
    public void importConnections(String p_excel_file) {
        def parameters = [ "p_excel_file"  :  p_excel_file,
                           "p_object_type" : "Connection",
                           "p_action"      : "Preview" ]
        tools.toolExecutor("inventory-import", parameters).execute()
    }
}
