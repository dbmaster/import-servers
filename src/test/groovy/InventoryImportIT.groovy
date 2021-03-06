import io.dbmaster.testng.BaseToolTestNGCase
import static org.testng.Assert.assertTrue

import org.testng.annotations.Parameters
import org.testng.annotations.Test

import com.branegy.files.FileService
import com.branegy.files.FileReference
import org.apache.commons.io.IOUtils

public class InventoryImportIT extends BaseToolTestNGCase {

    @Test
    public void testConnectionImport() {
        // TODO implement file import (from testImportApps) and use import-connections-test.xls
        def parameters = [ "p_excel_file"  :  "import-connections.xslx",
                           "p_object_type" : "Connection",
                           "p_action"      : "Preview" ]
        tools.toolExecutor("inventory-import", parameters).execute()
    }
    
/*
    // https://github.com/dbmaster/inventory-import/issues/9
    @Test
    public void testIgnoreImportColumns() {
        def parameters = [ "p_excel_file"    : "import-connections.xslx",
                           "p_object_type"   : "Connection",
                           "p_field_mapping" : "InputColumn=",
                           "p_action"        : "Preview" ]
        tools.toolExecutor("inventory-import", parameters).execute()
    }
*/

    @Test
    public void testImportApplications() {
        def fileService = getInjector().getInstance(FileService.class)
        FileReference file = null
        def filename = "applications.xlsx"
        try {
            file = fileService.getFile(filename)
            // println "File ${p_filename} already exists. Replacing content"
        } catch (Exception e) {
            // this means file does not exists
            file = fileService.createFile(filename, "inventory-import-test")
        }
        def outputStream = file.getOutputStream()
        def input = new java.io.FileInputStream( new java.io.File(getTestResourcesDir(), filename) )
        IOUtils.copy(input, outputStream)
        input.close()
        outputStream.close()

        def parameters = [ "p_excel_file"    : filename,
                           "p_object_type"   : "Application",
                           // "p_field_mapping" : "InputColumn=",
                           "p_action"        : "Preview" ]
        tools.toolExecutor("inventory-import", parameters).execute()
        
        fileService.deleteFile(file)
    }


}