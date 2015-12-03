import io.dbmaster.testng.BaseToolTestNGCase;
import static org.testng.Assert.assertTrue;

import org.testng.annotations.Parameters;
import org.testng.annotations.Test

public class InventoryImportIT extends BaseToolTestNGCase {

    @Test
    public void testConnectionImport() {
        def parameters = [ "p_excel_file"  :  "import-connections.xslx",
                           "p_object_type" : "Connection",
                           "p_action"      : "Preview" ]
        tools.toolExecutor("inventory-import", parameters).execute()
    }
    
    
    @Test
    public void testIgnoreImportColumns() {
        def parameters = [ "p_excel_file"    : "import-connections.xslx",
                           "p_object_type"   : "Connection",
                           "p_field_mapping" : "InputColumn=",
                           "p_action"        : "Preview" ]
        tools.toolExecutor("inventory-import", parameters).execute()
    }
    
}