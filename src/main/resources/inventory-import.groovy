import java.util.ArrayList
import java.util.Iterator
import java.util.List
import java.util.Map.Entry

import ExcelSynchronizer.MissingObject

import static com.branegy.persistence.custom.BaseCustomEntity.getDiscriminator;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.lang.reflect.ParameterizedType;
import java.text.SimpleDateFormat;

import org.apache.commons.io.IOUtils

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.usermodel.DateUtil

import org.slf4j.Logger;

import com.branegy.dbmaster.custom.CustomFieldConfig;
import com.branegy.dbmaster.custom.CustomFieldConfig.Type;
import com.branegy.dbmaster.custom.field.server.api.ICustomFieldService;

import com.branegy.inventory.api.InventoryService;
import com.branegy.inventory.model.*;
import com.branegy.persistence.custom.BaseCustomEntity;
import com.branegy.scripting.DbMaster;
import com.branegy.service.core.exception.EntityNotFoundApiException;

import com.branegy.dbmaster.model.NamedObject;
import com.branegy.dbmaster.sync.api.*
import com.branegy.dbmaster.sync.impl.RootObject
import com.branegy.dbmaster.sync.impl.BeanComparer;
import com.branegy.dbmaster.sync.api.SyncPair.ChangeType;
import com.branegy.dbmaster.sync.api.SyncAttributePair.AttributeChangeType;

import com.branegy.service.connection.api.ConnectionService;
import com.branegy.service.connection.model.DatabaseConnection;
import com.branegy.service.core.QueryRequest
import com.branegy.dbmaster.sync.api.SyncSession.SearchTarget


//CHECKSTYLE:OFF
public class ExcelSynchronizer extends SyncSession {

    boolean error;
    int objectKeyColumnIndex = -1;
    List<ColumnInfo> columnConfig;
    Logger logger;
    Set<String> contactNameSet = new TreeSet<String>();
    Set<String> objectNameSet = new TreeSet<String>();

    def processedObjects = [:]
    //Map<String, Contact> processedContact = new HashMap<String,Contact>();
    //List<ContactLink>    processedContactLink = new ArrayList<ContactLink>();
    String contacts=""
    DbMaster dbm
    InventoryService inventoryService
    Sheet sheet;

    def missingMetaValues = [:]

    List dateParsers = [ new SimpleDateFormat("MMM d, yyyy"),
                         new SimpleDateFormat("EEE MMM d, yyyy h:mm a"),
                         new SimpleDateFormat("MMM d,yyyy h:mm a"),
                         new SimpleDateFormat("EEE MMM d h:mm:ss z yyyy"),
                         new SimpleDateFormat("MMM d,yyyy h:mm:ss")]
                            

    private void logError(String error){
        this.error = true;
        logger.error(error);
    }

    private void logInfo(String info) {  logger.info(info); }
    private void logWarn(String warn) {  logger.warn(warn); }

    public static enum MissingObject { CREATE, IGNORE }

    Class targetClass;
    String keyColumnName;

    ExcelSynchronizer(DbMaster dbm, Logger logger, Class targetClass, String keyColumnName, String objectType, String objectFilter) {
        super(new InventoryComparer(objectType, objectFilter))
        setNamer(new InventoryNamer())
        this.dbm = dbm
        this.logger = logger
        this.targetClass = targetClass
        this.keyColumnName = keyColumnName
        inventoryService = dbm.getService(InventoryService.class)
    }

     public void importChanges(SyncPair pair) {
        String objectType = pair.getObjectType();
        if (objectType.equals("Inventory")) {
            pair.getChildren().each { importChanges(it) }
        } else if (objectType.equals("Server") || objectType.equals("Application")) {
            BaseCustomEntity sourceObj = (BaseCustomEntity)pair.getSource()
            BaseCustomEntity targetObj = (BaseCustomEntity)pair.getTarget()

            switch (pair.getChangeType()) {
                case ChangeType.NEW:
                    if (objectType.equals("Server")) {
                        inventoryService.createServer(targetObj)
                    } else if (objectType.equals("Application")) {
                        inventoryService.createApplication(targetObj)
                    }
                    break;
                case ChangeType.CHANGED:
                    for (SyncAttributePair attr : pair.getAttributes()) {
                        if (attr.getChangeType() != AttributeChangeType.EQUALS) {
                            sourceObj.setCustomData( attr.getAttributeName(), attr.getTargetValue())
                        }
                    }
                    if (objectType.equals("Server")) {
                        inventoryService.saveServer(sourceObj)
                    } else if (objectType.equals("Application")) {
                        inventoryService.updateApplication(sourceObj)
                    }
                    break;
                case ChangeType.DELETED:
                    if (objectType.equals("Server")) {
                        inventoryService.deleteServer(sourceObj.getId())
                    } else if (objectType.equals("Application")) {
                        inventoryService.deleteApplication(sourceObj.getId())
                    }
                    break;
                case ChangeType.COPIED:
                case ChangeType.EQUALS:
                    break;
                default:
                    throw new RuntimeException("Unexpected change type ${pair.getChangeType()}")
            }
        } else {
            throw new SyncException("Unexpected object type ${objectType}");
        }
    }

    public void applyChanges() {
        try {
            importChanges(getSyncResult());
        } finally {
            dbm.closeResources()
        }
    }

    public boolean loadAndValidateExcel(Map parameters) throws Exception {
        //MissingObject missingObjectAction = MissingObject.valueOf(parameters.p_objects.toUpperCase())
        //MissingObject contactConfig =  MissingObject.valueOf(parameters.p_contacts.toUpperCase())

        InputStream fis = null
        try {
            logInfo("Loading data from file "+parameters.p_excel_file.getName())

            fis = parameters.p_excel_file.getInputStream()
            Workbook wb = WorkbookFactory.create(fis)

            sheet = wb.getSheetAt(0)

            Row header = sheet.getRow(0);
            Set<String> headerSet = new LinkedHashSet<String>();
            for (Cell cell:header) {
                headerSet.add(cell.getStringCellValue());
            }
            def mapping = parameters.p_field_mapping
            validateHeader(dbm.getService(ICustomFieldService.class), headerSet, mapping);
            if (error) {
                dbm.setRollbackOnly();
            }
            return !error;
        } catch (Exception e) {
            dbm.setRollbackOnly();
            throw e;
        } finally {
            IOUtils.closeQuietly(fis);
        }
    }

    private void validateHeader(ICustomFieldService service, Set<String> headerSet, String fieldMappingStr) {
        if (service.getConfigByName(Contact.class, Contact.NAME)==null){
            logError((Contact.class)+"."+Contact.NAME+" not found");
            return;
        }
        CustomFieldConfig contactLinkRole = service.getConfigByName(ContactLink.class, ContactLink.ROLE);
        if (contactLinkRole == null){
            logError(getDiscriminator(ContactLink.class)+"."+ContactLink.ROLE+" not found");
            return;
        }
        if (contactLinkRole.getTextValues().isEmpty()){
            logError(getDiscriminator(ContactLink.class)+"."+ContactLink.ROLE+" must be multivalue");
            return;
        }
        Set<String> contactLinkRoleSet = new HashSet<String>(contactLinkRole.getTextValues());

        columnConfig = []
        Pattern pattern = Pattern.compile("Contact\\(([^)]+)\\)\\.(.+)");
        Map<String,Integer> role2ContactNameIndex = new HashMap<String, Integer>();
        int i = -1;

        def fieldMapping = [:]
        if (fieldMappingStr!=null) {
            fieldMappingStr.split("\n").each { pair ->
                def key_value = pair.trim().split("=")
                fieldMapping.put(key_value[0], key_value[1])
            }
        }
        logInfo("Field Mappings="+fieldMapping)

        for (String value:headerSet) {
            i++;
            if (fieldMapping[value]!=null) {
                logInfo("Replaced field ${value} with ${fieldMapping[value]}")
                value = fieldMapping[value]
            }

            Matcher matcher = pattern.matcher(value);
            if (matcher.matches()) { // contact field
                String role = matcher.group(1);
                if (!contactLinkRoleSet.contains(role)){
                    logError("${getDiscriminator(ContactLink.class)}.${ContactLink.ROLE} ${role} not in ${contactLinkRoleSet}");
                }
                String fieldName = matcher.group(2);
                CustomFieldConfig config = service.getConfigByName(Contact.class, fieldName);
                if (config == null){
                    logError("Field "+getDiscriminator(Contact.class)+".["+fieldName+"] not found");
                    continue;
                }
                if (fieldName.equals(Contact.NAME)){
                    if (role2ContactNameIndex.put(role, i) != null){
                        logError("Multiple contactName for role "+role);
                    }
                }
                logInfo("Found contact field " + fieldName + " for role ${role} index "+i);
                columnConfig.add(new ColumnInfo(i, role, config));
            } else { // simple field
                String fieldName = value;
                CustomFieldConfig config = service.getConfigByName(targetClass, fieldName);
                if (config == null){
                    logError("Field ${getDiscriminator(targetClass)}.[${fieldName}] not found");
                    continue;
                }
                if (keyColumnName.equals(fieldName)){
                    objectKeyColumnIndex = i;
                    continue;
                }
                // TODO Handle situation when fields we have duplicates in source field names
                columnConfig.add(new ColumnInfo(i, null, config));
            }
        }
        if (objectKeyColumnIndex == -1){
            logError("Key field ${keyColumnName} not found in excel");
        }
        for (ColumnInfo ci:columnConfig){
            if (ci.contactRole!=null){
                ci.contactNameIndex = role2ContactNameIndex.get(ci.contactRole);
                if (ci.contactNameIndex == -1){
                    logError(getDiscriminator(Contact.class)+"."+Contact.NAME+" not set for role "+ci.contactRole);
                }
            }
        }
    }

    private String autoIncrementName(String name) {
        Pattern pattern = Pattern.compile("(.+) (\\d+)");
        Matcher matcher = pattern.matcher(name);
        if (matcher.matches()){
            int i = Integer.valueOf(matcher.group(2))+1;
            return matcher.group(1)+" ("+i+")";
        } else {
            return name+" (1)";
        }
    }

    private List getExcelObjects() {
        for (Row row : sheet) {
            if (row.getRowNum()==0) { // skip header row
                continue;
            }
            String objectName = getStringValue(row, objectKeyColumnIndex);
            if (objectName == null) {
                logWarn("${getDiscriminator(getDiscriminator)}.${keyColumnName} is not set at row ${row.getRowNum()}");
                continue;
            }
            BaseCustomEntity objectToImport;

            objectToImport = processedObjects[objectName];
            while (objectToImport != null) {
                // actually here is a duplicate
                objectName = autoIncrementName(objectName);
                logWarn("Duplicate entity found at row ${row.getRowNum()}");
                objectToImport = processedObjects.get(objectName);
            }
            objectToImport = targetClass.newInstance();
            objectToImport.setCustomData(keyColumnName, objectName);
            logInfo("Processing row ${row.getRowNum()}: ${objectName}");

            processedObjects[objectName] = objectToImport;

            for (ColumnInfo info : columnConfig) {
                // logInfo("Info role "+ info.contactRole+" "+info.contactNameIndex);
                String value = getStringValue(row, info.index);
                if (info.contactRole!=null) {
                    // skip for now
                } else {
                    // logInfo("Set custom field  "+info.field.name+" to value "+ value);
                    setupCustomField(info.field, objectToImport, value, row.getRowNum(), info.index);
                }
            }
        }
        return processedObjects.collect{ key, value -> value }
    }

    protected ContactLink findByRole(String role, List<ContactLink> contactLinks) {
        for (ContactLink link:contactLinks) {
            if (role.equals(link.getCustomData(ContactLink.ROLE))) {
                return link;
            }
        }
        return null;
    }

    protected void setupCustomField(CustomFieldConfig config, BaseCustomEntity entity, String value, int row, int column) {
        Object v;
        if (config.getType() == Type.BOOLEAN) {
            if (value == null || value.isEmpty()) {
                v = null;
            } else if ("Yes".equalsIgnoreCase(value)){
                v = Boolean.TRUE;
            } else if ("No".equalsIgnoreCase(value)){
                v = Boolean.FALSE;
            } else {
                logError("'${value}' is not boolean for field ${config.name} at ${row}:${column}");
                return;
            }
        } else if (config.getType() == Type.STRING || config.getType() == Type.TEXT) {
            v = value;
        } else if (config.getType() == Type.FLOAT) {
            if (value != null) {
                try {
                    v = new Float(value)
                } catch (NumberFormatException e) {
                    logError("'${value}' is not a float for field ${config.name} at ${row}:${column}")
                }
            } else {
                v = null
            }
        } else if (config.getType() == Type.DATE) {
            if (value == null || value.isEmpty()) {
                v = null;
            } else {
                  if (dateParsers.find {
                        try {
                            v = new java.sql.Timestamp(it.parse(value).getTime());
                            return true
                        } catch (java.text.ParseException e) {
                            return false
                        }
                    } == null) {
                        logError("'"+value+"' is not date for field "+"."+config.name+" at "+row+":"+column);
                  }

            }
        } else {
            logError("Unsupported type ${config.getType()} for column [${config.getName()}]"+
                    " at row ${row} and column ${column}");
            return;
        }
        if (v==null && config.isRequired()){
            logError("Value is required for ${config.getName()} for field ${config.getClazz()}.${config.getName()} at ${row}:${column}");
            return;
        }
        List<String> textValues = config.getTextValues();
        if (v!=null && !textValues.isEmpty() && !textValues.contains(v)){
            logError("Value '${value}' not in ${textValues} for field ${config.getClazz()}.${config.getName()} at ${row}:${column}");
            def key = config.getClazz()+"."+config.getName()
            def newValuesPerField = missingMetaValues[key]
            if (newValuesPerField == null) {
                newValuesPerField = [value] as Set
                missingMetaValues[key] = newValuesPerField
            } else {
                newValuesPerField.add(value)
            }
            return;
        }
        //if (v!=null) {
        //    logInfo("Set attribute ${config.getName()}:${config.getType()} value  ${v} of type ${v.getClass().getName()}");
        //}
        entity.setCustomData(config.getName(), v);
    }

    protected String getStringValue(Row row,int columnIndex) {
        Cell cell = row.getCell(columnIndex)
        if (cell!=null) {
            def value;
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    value = cell.getRichStringCellValue().getString()
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = cell.getDateCellValue()
                    } else {
                        value = cell.getNumericCellValue()
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    value = cell.getBooleanCellValue()
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    value = cell.getCellFormula()
                    break;
                case Cell.CELL_TYPE_BLANK: 
                    break;
                default:
                    throw new RuntimeException("Unsupported cell type ${cell.getCellType()}")
                    break
            }
            if (value!=null) {
                value = value.toString().trim();
                if (value.isEmpty()) {
                    return null;
                } else {
                    return value;
                }
            } else {
                return null;
            }
            // String value = cell.getStringCellValue();
        } else {
            return null;
        }
    }

    static class ColumnInfo {
        String contactRole;
        int contactNameIndex = -1;

        int index;
        CustomFieldConfig field;

        public ColumnInfo(int index, String contactRole, CustomFieldConfig field) {
            this.index = index;
            this.contactRole = contactRole;
            this.field = field;
        }
    }
}

class InventoryNamer implements Namer {
        @Override
        public String getName(Object o) {
            if (o instanceof RootObject) {                 return "Inventory";
            } else if (o instanceof Server) {              return ((Server)o).getServerName();
            } else if (o instanceof Application) {         return ((Application)o).getApplicationName();
            } else {
                throw new IllegalArgumentException("Unexpected object class "+o);
            }
        }

        @Override
        public String getType(Object o) {
            if (o instanceof RootObject) {                 return "Inventory";
            } else if (o instanceof Server) {              return "Server";
            } else if (o instanceof Application) {         return "Application";
            } else {
                throw new IllegalArgumentException("Unexpected object class "+o);
            }
        }
}

class InventoryComparer extends BeanComparer {
    def connections
    def inventoryDBs
    String objectFilter, objectType
    
    InventoryComparer(String objectType, String objectFilter) {
        this.objectFilter = objectFilter
        this.objectType   = objectType
    }

    @Override
    public void syncPair(SyncPair pair, SyncSession session) {
        String objectType = pair.getObjectType();
        Namer namer = session.getNamer();
        if (objectType.equals("Inventory")) {
            def request = objectFilter == null ? new QueryRequest() : new QueryRequest(objectFilter)
            def inventoryObjects;
            if (objectType.equals("Server")) {
                inventoryObjects= session.inventoryService.getServerList(request)
            } else if (objectType.equals("Application")) {
                inventoryObjects= session.inventoryService.getApplicationList(request)
            }
            def importedObjects = session.getExcelObjects()

            session.logInfo("Total imported objects ${importedObjects.size()}")
            def childs = mergeCollections(pair, inventoryObjects, importedObjects, namer)

            session.logInfo("Total pairs ${childs.size()}")

            pair.getChildren().addAll(childs);
        } else if (objectType.equals("Server") || objectType.equals("Application")) {
            BaseCustomEntity sourceObject = (BaseCustomEntity)pair.getSource()
            BaseCustomEntity targetObject = (BaseCustomEntity)pair.getTarget()

            try {
                Map sourceAttrs = sourceObject == null ? null : new HashMap()
                Map targetAttrs = targetObject == null ? null : new HashMap()

                // take into consideration only attributes that came from Excel
                session.columnConfig.each { ci ->
                    // println "name="+ci.field.name
                    if (ci.contactRole==null) {
                        if (sourceObject!=null) {
                            sourceAttrs[ci.field.name]=sourceObject.getCustomData(ci.field.name)
                        }
                        if (targetObject!=null) {
                            targetAttrs[ci.field.name]=targetObject.getCustomData(ci.field.name)
                        }
                    }
                }
                pair.setAttributes(mergeAttributes(sourceAttrs, targetAttrs))
            } catch (Exception e) {
                session.logError("error ${e.getMessage()}");
                // TODO show message to user (finally should be shown)
                e.printStackTrace()
                throw e
            }
        } else {
            throw new SyncException("Unexpected object type ${objectType}");
        }
    }
}

if (p_object_type.equals("Server")) {
    synchronizer = new ExcelSynchronizer(dbm, logger, Server.class, Server.SERVER_NAME, p_object_type, p_object_filter)
} else if (p_object_type.equals("Application")) {
    synchronizer = new ExcelSynchronizer(dbm, logger, Application.class, Application.APPLICATION_NAME, p_object_type, p_object_filter)
} else {
    println "Unexpected object type ${p_object_type}. Only Server and Application are supported"
    return
}

if (synchronizer.loadAndValidateExcel(parameters)) {
    logger.info("Parsing excel")
    rootObject = new RootObject("Repository", "Excel")
    synchronizer.syncObjects(rootObject, rootObject)

    // logger.info("synchronizer log:"+synchronizer.getLog())

    if (synchronizer.missingMetaValues.size()>0) {
        println "<p>Missing value(s) in Custom Fields with enumerations. Fix errors and re-run import</p>"
        synchronizer.missingMetaValues.each { key, value ->
            println "<hr size=\"1\" /><p>Missing value(s) for ${key}</p><br/>"
            println value.join("<br/>")
        }
        println "<hr size=\"1\" />"
    } else {
        def syncService = dbm.getService(SyncService.class)
        def sessionHtml = syncService.generateSyncSessionPreviewHtml(synchronizer, false)
        if (p_action.equals("Import")) {
            logger.info("Importing changes")
            synchronizer.applyChanges();
            synchronizer.setParameter("html",  sessionHtml.toString())
            synchronizer.setParameter("title", "Inventory import from excel")
            syncService.saveSession(synchronizer, "Inventory Import (${p_object_type})")
            logger.info("Import completed successfully")
        }
        logger.info("Generating change log")
        println sessionHtml
    }
} else {
    println "Import failed. Check log for errors"
    logger.error("Import failed")
}