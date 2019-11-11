package com.eisgroup.exercise;

import com.eisgroup.annotation.ClassTypeDesc;
import com.eisgroup.annotation.FieldDesc;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * @Description:
 * @Date: 2019/10/31 14:03
 * @author: xfei
 */
public class POIUtil {

    public static final String SHEET_NAME = "table_name";

    public static final String SHEET_FIELD_NAME = "table_field_name";

    public static final String XLSX_SUFFIX = ".xlsx";

    public static final String XLS_SUFFIX = ".xls";

    public static <T> List<T> parseExcel(String filePath, Class<T> tc) {
        // 从 POITest-one.xlsx 文件中读取数据，并保存到 List<ToolDto> 中后打印输出。
        List<T> list = new ArrayList<>();
        Workbook wb = getWorkbook(filePath);
        if (wb != null) {
            // 获得excel的所有的sheet
            List<Sheet> sheets = getExcelSheets(wb);
            Map<String, Object> map = getModelFields(tc);
            sheets.parallelStream().forEach(sheet -> {
                if (!StringUtils.isEmpty(sheet.getSheetName()) && sheet.getSheetName().equals(map.get(SHEET_NAME))) {
                    Map<String, String> fieldMaps = (Map<String, String>)map.get(SHEET_FIELD_NAME);
                    Row titleBarRow = sheet.getRow(0);
                    fieldMaps = convertEntryFieldDescToRowNumber(fieldMaps, titleBarRow);
                    // 4、循环读取表格数据
                    for (Row row : sheet) {
                        // 首行（即表头）不读取
                        if (row.getRowNum() == 0) {
                            continue;
                        }

                        T t = convertRowToMap(fieldMaps, row, tc);
                        list.add(t);
                    }
                }
            });
        }
        if (wb != null) {
            try {
                wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return list;
    }

    /**
     *
     * @param fieldMap 实体字段信息
     * @param row excel行信息
     * @param clazz 实体字节码信息
     * @exception: 将excel的row信息保存到map集合中
     * @date: 2019/11/1 17:16
     * @return: T
     */
    private static <T> T convertRowToMap(Map<String, String> fieldMap, Row row, Class<T> clazz) {
        Set<Map.Entry<String, String>> set = fieldMap.entrySet();
        Map<String, String> EntityMap = new HashMap<>();
        for (Cell cell : row) {
            String cellValue = getCellValue(cell);
            Iterator<Map.Entry<String, String>> iterator = set.iterator();
            while (iterator.hasNext()) {
                Map.Entry<String, String> entry = iterator.next();
                if (entry.getValue().equals(String.valueOf(cell.getColumnIndex()))) {
                    EntityMap.put(entry.getKey(), cellValue);
                    break;
                }
            }
        }
        T t = convertMapToEntity(EntityMap, clazz);
        return t;
    }

    /**
     *
     * @param map 数据集合
     * @param clazz 实体字节码信息
     * @exception: 将map集合转换成实例
     * @date: 2019/11/1 17:19
     * @return: T
     */
    private static <T> T convertMapToEntity(Map<String, String> map, Class<T> clazz) {
        Object obj = null;
        try {
            obj = clazz.newInstance();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        Set<Map.Entry<String, String>> entrySet = map.entrySet();
        Iterator<Map.Entry<String, String>> iterator = entrySet.iterator();
        while (iterator.hasNext()) {
            Map.Entry<String, String> entry = iterator.next();
            //属性名
            String propertyName = entry.getKey();
            String value = entry.getValue();
            String setMethodName = "set"
                    + propertyName.substring(0, 1).toUpperCase()
                    + propertyName.substring(1);
            Field field = getClassField(clazz, propertyName);
            if (field == null) {
                continue;
            }
            Class<?> fieldTypeClass = field.getType();
            value = convertValType(value, fieldTypeClass);
            try{
                clazz.getMethod(setMethodName, field.getType()).invoke(obj, value);
            }catch(NoSuchMethodException e){
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            } catch (InvocationTargetException e) {
                e.printStackTrace();
            }
        }
        return (T)obj;
    }

    /**
     *
     * @param value 字段内容
     * @param clazz 实体字节码信息
     * @exception: 将字段转换成对应的类型(暂时是String类型)
     * @date: 2019/11/4 9:21
     * @return: java.lang.String
     */
    private static <T> String convertValType(String value, Class<T> clazz) {
        Object retVal = null;
        if(Long.class.getName().equals(clazz.getName())
                || long.class.getName().equals(clazz.getName())) {
            retVal = Long.parseLong(value);
        } else if(Integer.class.getName().equals(clazz.getName())
                || int.class.getName().equals(clazz.getName())) {
            retVal = Integer.parseInt(value);
        } else if(Float.class.getName().equals(clazz.getName())
                || float.class.getName().equals(clazz.getName())) {
            retVal = Float.parseFloat(value);
        } else if(Double.class.getName().equals(clazz.getName())
                || double.class.getName().equals(clazz.getName())) {
            retVal = Double.parseDouble(value);
        } else {
            retVal = value;
        }
        return String.valueOf(retVal);
    }

    /**
     *
     * @param clazz 实体类字节码信息
     * @param fieldName 字段名称
     * @exception: 反射得到字段
     * @date: 2019/11/4 9:21
     * @return: java.lang.reflect.Field
     */
    private static <T> Field getClassField(Class<T> clazz, String fieldName) {
        if( Object.class.getName().equals(clazz.getName())) {
            return null;
        }
        Field []declaredFields = clazz.getDeclaredFields();
        for (Field field : declaredFields) {
            if (field.getName().equals(fieldName)) {
                return field;
            }
        }

        Class<?> superClass = clazz.getSuperclass();
        if(superClass != null) {
            // 简单的递归一下
            return getClassField(superClass, fieldName);
        }
        return null;
    }

    /**
     *
     * @param fieldMaps 实体字段map
     * @param titleBarRow excel标题行
     * @exception: 将实体中字段描述转换为excel中所对应的的行号
     * @date: 2019/11/1 15:26
     * @return:
     */
    private static Map<String, String> convertEntryFieldDescToRowNumber(Map<String, String> fieldMaps, Row titleBarRow) {
        // 获取行的迭代器
        Iterator<Cell> i = titleBarRow.cellIterator();
        Set<Map.Entry<String, String>> entrySet = fieldMaps.entrySet();
        while(i.hasNext()) {
            Cell cell = i.next();
            if (fieldMaps.containsValue(String.valueOf(cell).trim())) {
                Iterator<Map.Entry<String, String>> iterator = entrySet.iterator();
                while (iterator.hasNext()) {
                    Map.Entry<String, String> entry = iterator.next();
                    if (entry.getValue().equals(String.valueOf(cell).trim())) {
                        fieldMaps.replace(entry.getKey(), String.valueOf(cell.getColumnIndex()));
                        break;
                    }
                }
            }
        }
        return fieldMaps;
    }

    /**
     *
     * @param tc 实体类
     * @exception: 获取实体类中注解信息，并解析出实体的字段信息
     * @date: 2019/11/1 10:46
     * @return: java.util.Map<java.lang.String,java.lang.String>
     */
    private static <T> Map<String, Object> getModelFields(Class<T> tc) {
        Map<String, Object> map = new HashMap<>();
        boolean hasAnnotation = tc.isAnnotationPresent(ClassTypeDesc.class);
        if (hasAnnotation) {
            ClassTypeDesc ctd = tc.getAnnotation(ClassTypeDesc.class);
            map.put(SHEET_NAME, ctd.value());
        }
        Field[] fields = tc.getDeclaredFields();
        Map<String, String> fieldMaps = new LinkedHashMap<>();
        for (Field field : fields) {
            FieldDesc fieldDesc = field.getAnnotation(FieldDesc.class);
            if (fieldDesc != null) {
                fieldMaps.put(field.getName(), fieldDesc.value());
            }
        }
        map.put(SHEET_FIELD_NAME, fieldMaps);
        return map;
    }

    /**
     *
     * @param cell Cell excel单元格
     * @exception: 获取excel的单元格信息
     * @date: 2019/10/31 17:44
     * @return: java.lang.String
     */
    private static String getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();
        String cellValue = "";

        switch(cellType) {
            case STRING:
                cellValue = cell.getRichStringCellValue().getString();
                cellValue = StringUtils.isEmpty(cellValue) ? "" : cellValue;
                break;
            case NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                    cell.getDateCellValue();
                    try {
                        cellValue = DateUtils.dateFormat(cell.getDateCellValue(), DateUtils.DATE_PATTERN);
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                } else {
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case BLANK:
                break;
        }
        return cellValue;
    }

    /**
     *
     * @param wb excel解析工作簿
     * @exception: 将wb中所有的sheet放入list集合中
     * @date: 2019/10/31 16:18
     * @return: java.util.List<org.apache.poi.ss.usermodel.Sheet>
     */
    private static List<Sheet> getExcelSheets(Workbook wb) {
        List<Sheet> list = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i ++) {
            list.add(wb.getSheetAt(i));
        }
        return list;
    }

    /**
     *
     * @param filePath String 文件路径
     * @exception: 将excel转换成workbook对象
     * @date: 2019/10/31 15:38
     * @return: org.apache.poi.ss.usermodel.Workbook
     */
    private static Workbook getWorkbook(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String suffix = filePath.substring(filePath.lastIndexOf("."));
        InputStream fis = null;
        try {
            fis = new FileInputStream(new File(filePath));
            if (XLS_SUFFIX.equals(suffix)) {
                wb = new HSSFWorkbook(fis);
            } else if (XLSX_SUFFIX.equals(suffix)) {
                wb = new XSSFWorkbook(fis);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return wb;
    }

    /**
     *
     * @param entityList 需要转化的实体集合
     * @param fileName excel文件名称
     * @param clazz 对象实体的字节码信息
     * @param filePath 可选参数，如果填则将生成的excel文件放在改路径下，如果不填，则将文件放在项目路径下
     * @exception: 将实体集合转换成excel
     * @date: 2019/11/4 14:17
     * @return: void
     */
    public static <T> void exportExcel(List<T> entityList, String fileName, Class<T> clazz, String ... filePath) {
        // 创建表格信息
        Workbook wb = createExcel(entityList, fileName, clazz);
        if (wb == null) {
            return;
        }
        // 导出excel信息
        if (filePath != null && filePath.length > 0) {
            String path = filePath[0];
            File file = new File(path + "/" + fileName);
            FileOutputStream fos = null;
            try {
                fos = new FileOutputStream(file);
                wb.write(fos);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                if (fos != null) {
                    try {
                        fos.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if (wb != null) {
                    try {
                        wb.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }



    /**
     *
     * @param entityList 需要转化的实体集合
     * @param fileName excel文件名称
     * @param clazz 对象实体的字节码信息
     * @exception: 将实体集合转换成excel
     * @date: 2019/11/4 10:25
     * @return Workbook
     */
    private static <T> Workbook createExcel(List<T> entityList, String fileName, Class<T> clazz) {
        Workbook wb = createWorkbook(fileName);
        // 获取字节码的中注解信息，用于生成表格
        Map<String, Object> map = getModelFields(clazz);
        Sheet sheet;
        String sheetName = String.valueOf(map.get(SHEET_NAME));
        if (StringUtils.isEmpty(sheetName)) {
            sheet = wb.createSheet("sheet1");
        } else {
            sheet = wb.createSheet(sheetName);
        }
        // 设置表格的格式
        CellStyle cellStyle = wb.createCellStyle();
        // 居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        // 设置字体
        font.setFontName("黑体");
        // 设置字体大小
        font.setFontHeightInPoints((short)16);
        cellStyle.setFont(font);
        // 设置自动换行
        cellStyle.setWrapText(true);

        Row row = sheet.createRow(0);
        // 设置行高
        row.setHeightInPoints((short)22);
        Map<String, String> fieldsMap = (Map<String, String>)map.get(SHEET_FIELD_NAME);
        if (fieldsMap != null) {
            Set<Map.Entry<String, String>> set = fieldsMap.entrySet();
            Iterator<Map.Entry<String, String>> iterator = set.iterator();
            int temp = 0;
            Cell cell = null;
            while(iterator.hasNext()) {
                Map.Entry<String, String> entry = iterator.next();
                cell = row.createCell(temp);
                cell.setCellValue(entry.getValue());
                cell.setCellStyle(cellStyle);
                sheet.setColumnWidth(temp, entry.getValue().getBytes().length * 2 * 256);
                temp ++;
            }


            Field[] fields = clazz.getDeclaredFields();

            // 设置表格的格式
            CellStyle columnCellStyle = wb.createCellStyle();
            columnCellStyle.setAlignment(HorizontalAlignment.CENTER);
            Font columnFont = wb.createFont();
            columnFont.setFontName("宋体");
            columnFont.setFontHeightInPoints((short)13);
            columnCellStyle.setFont(columnFont);
            // 设置自动换行
            cellStyle.setWrapText(true);

            for (int i = 0; i < entityList.size(); i ++) {
                Row entityRow = sheet.createRow(i + 1);
                entityRow.setHeightInPoints((short)17);
                T model = entityList.get(i);

                int line = 0;
                for (Field field : fields) {
                    // 真真的excel字段的遍历
                    for (Map.Entry<String, String> entry: set) {
                        if (field.getName().equals(entry.getKey())) {
                            field.setAccessible(true);
                            Cell cell1 = entityRow.createCell(line);
                            cell1.setCellStyle(columnCellStyle);
                            try {
                                cell1.setCellValue((String)field.get(model));
                            } catch (IllegalAccessException e) {
                                e.printStackTrace();
                            }
                            line ++;
                        }
                    }
                }
            }
        }
        return wb;
    }

    /**
     *
     * @param fileName 需要生成的excel文件名，根据文件名的后缀生成对应的工作簿类型
     * @exception: 创建excel的工作簿
     * @date: 2019/11/4 13:28
     * @return: org.apache.poi.ss.usermodel.Workbook
     */
    private static Workbook createWorkbook(String fileName) {
        String suffix = fileName.substring(fileName.lastIndexOf("."));
        Workbook wb = null;
        if (XLSX_SUFFIX.equals(suffix)) {
            wb = new XSSFWorkbook();
        } else {
            wb = new HSSFWorkbook();
        }
        return wb;
    }
}
