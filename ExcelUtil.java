import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * POI实现excel文件读写(导入/导出)操作工具类
 *
 * @author zhuxiongxian
 * @version 1.0
 * @keywords java excel poi 导入 导出
 * @date at 2016年11月10日 下午2:10:57
 * @modifiedBy idealism
 */
public class ExcelUtil {

    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 用于汇总多个 sheet 的 VO
     *
     * @author zhuxiongxian
     * @version 1.0
     * @date at 2016年12月14日 下午4:25:53
     */
    public static class ExcelSheet<T> {
        /**
         * sheet 的表名
         */
        private String sheetName;

        /**
         * sheet 的表头
         */
        private String[] headers;

        /**
         * sheet 的数据集
         */
        private Collection<T> dataset;

        /**
         * @return 表名
         */
        public String getSheetName() {
            return sheetName;
        }

        /**
         * @param sheetName 表名
         */
        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        /**
         * @return 表头列表
         */
        public String[] getHeaders() {
            return headers;
        }

        /**
         * @param headers 表头列表
         */
        public void setHeaders(String[] headers) {
            this.headers = headers;
        }

        /**
         * @return 数据集
         */
        public Collection<T> getDataset() {
            return dataset;
        }

        /**
         * @param dataset 数据集
         */
        public void setDataset(Collection<T> dataset) {
            this.dataset = dataset;
        }

    }

    /**
     * 对外提供读取excel的方法， 当且仅当只有一个sheet， 默认从第一个 sheet 读取数据
     *
     * @param filePath 文件路径
     * @return 文件第一个 sheet 的所有数据（包含表头）
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath) throws IOException {
        return readExcel(filePath, 0);
    }

    /**
     * 对外提供读取excel的方法， 根据 sheet 下标读取 sheet 数据
     *
     * @param filePath   文件路径
     * @param sheetIndex 表下标（下标从 0 开始）
     * @return 文件第 sheetIndex + 1 个 sheet 的所有数据（包含表头）
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath, int sheetIndex) throws IOException {
        File file = new File(filePath);
        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);
        
        return readExcel(new FileInputStream(file), extension, sheetIndex);
    }

    /**
     * 对外提供读取excel的方法， 根据 sheet 下标索引读取sheet对象，并指定行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始
     *
     * @param filePath         文件路径
     * @param sheetIndex       表下标（下标从 0 开始）
     * @param startRowIndex    起始行下标
     * @param startColumnIndex 起始列下标
     * @return 文件第 sheetIndex + 1 个 sheet 的 行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始的所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath, int sheetIndex, int startRowIndex, int startColumnIndex) throws IOException {
        File file = new File(filePath);
        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

        return readExcel(new FileInputStream(file), extension, sheetIndex, startRowIndex, startColumnIndex);
    }

    /**
     * 对外提供读取excel的方法， 根据sheet下标索引读取sheet对象， 并指定行列区间获取数据[startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex)
     *
     * @param filePath         文件路径
     * @param sheetIndex       表下标（下标从 0 开始）
     * @param startRowIndex    起始行下标
     * @param endRowIndex      结束行下标 + 1
     * @param startColumnIndex 起始列下标
     * @param endColumnIndex   结束列下标 + 1
     * @return 文件第 sheetIndex + 1 个 sheet 的区间 [startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex) 内所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath, int sheetIndex, int startRowIndex, int endRowIndex,
                                               int startColumnIndex, int endColumnIndex) throws IOException {
        File file = new File(filePath);
        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

        return readExcel(new FileInputStream(file), extension, sheetIndex, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
    }

    /**
     * 对外提供读取excel的方法， 根据sheet名称读取sheet数据
     *
     * @param filePath  文件路径
     * @param sheetName 表名
     * @return 文件第 表名为 sheetName 的所有数据（包含表头）
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath, String sheetName) throws IOException {
        File file = new File(filePath);
        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

        return readExcel(new FileInputStream(file), extension, sheetName);
    }

    /**
     * 对外提供读取excel的方法， 根据sheet名称读取sheet对象， 并指定行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始
     *
     * @param filePath         文件路径
     * @param sheetName        表名
     * @param startRowIndex    起始行下标
     * @param startColumnIndex 起始列下标
     * @return 文件表名为 sheetName 的 sheet 的 行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始的所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath, String sheetName, int startRowIndex, int startColumnIndex) throws IOException {
        File file = new File(filePath);
        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

        return readExcel(new FileInputStream(file), extension, sheetName, startRowIndex, startColumnIndex);
    }

    /**
     * 对外提供读取excel的方法， 根据sheet名称读取sheet对象， 并指定行列区间获取数据[startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex)
     *
     * @param filePath         文件路径
     * @param sheetName        表名
     * @param startRowIndex    起始行下标
     * @param endRowIndex      结束行下标 + 1
     * @param startColumnIndex 起始列下标
     * @param endColumnIndex   结束列下标 + 1
     * @return 文件表名为 sheetName 的 sheet 的区间 [startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex) 内所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(String filePath, String sheetName, int startRowIndex, int endRowIndex,
                                               int startColumnIndex, int endColumnIndex) throws IOException {
        File file = new File(filePath);
        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

        return readExcel(new FileInputStream(file), extension, sheetName, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
    }

    /**
     * 对外提供读取excel的方法， 当且仅当只有一个sheet， 默认从第一个 sheet 读取数据
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @return 文件第一个 sheet 的所有数据（包含表头）
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension) throws IOException {
        return readExcel(inputStream, extension, 0);
    }

    /**
     * 对外提供读取excel的方法， 根据 sheet 下标读取 sheet 数据
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @param sheetIndex 表下标（下标从 0 开始）
     * @return 文件第 sheetIndex + 1 个 sheet 的所有数据（包含表头）
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension, int sheetIndex) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        Workbook workbook = getWorkbook(inputStream, extension);
        if (workbook != null) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            list = getSheetData(workbook, sheet);
        }
        return list;
    }

    /**
     * 对外提供读取excel的方法， 根据 sheet 下标索引读取sheet对象，并指定行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @param sheetIndex       表下标（下标从 0 开始）
     * @param startRowIndex    起始行下标
     * @param startColumnIndex 起始列下标
     * @return 文件第 sheetIndex + 1 个 sheet 的 行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始的所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension, int sheetIndex, int startRowIndex, int startColumnIndex) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        Workbook workbook = getWorkbook(inputStream, extension);
        if (workbook != null) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            // 获取总行数
            int rowNum = sheet.getPhysicalNumberOfRows();
            // 获取第一行的总列数
            int colNum = sheet.getRow(0).getPhysicalNumberOfCells();
            list = getSheetData(workbook, sheet, startRowIndex, rowNum, startColumnIndex, colNum);
        }
        return list;
    }

    /**
     * 对外提供读取excel的方法， 根据sheet下标索引读取sheet对象， 并指定行列区间获取数据[startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex)
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @param sheetIndex       表下标（下标从 0 开始）
     * @param startRowIndex    起始行下标
     * @param endRowIndex      结束行下标 + 1
     * @param startColumnIndex 起始列下标
     * @param endColumnIndex   结束列下标 + 1
     * @return 文件第 sheetIndex + 1 个 sheet 的区间 [startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex) 内所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension, int sheetIndex, int startRowIndex, int endRowIndex,
                                               int startColumnIndex, int endColumnIndex) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        Workbook workbook = getWorkbook(inputStream, extension);
        if (workbook != null) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            list = getSheetData(workbook, sheet, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }
        return list;
    }

    /**
     * 对外提供读取excel的方法， 根据sheet名称读取sheet数据
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @param sheetName 表名
     * @return 文件第 表名为 sheetName 的所有数据（包含表头）
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension, String sheetName) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        Workbook workbook = getWorkbook(inputStream, extension);
        if (workbook != null) {
            Sheet sheet = workbook.getSheet(sheetName);
            list = getSheetData(workbook, sheet);
        }
        return list;
    }

    /**
     * 对外提供读取excel的方法， 根据sheet名称读取sheet对象， 并指定行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @param sheetName        表名
     * @param startRowIndex    起始行下标
     * @param startColumnIndex 起始列下标
     * @return 文件表名为 sheetName 的 sheet 的 行下标从 startRowIndex 开始，列下标从 startColumnIndex 开始的所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension, String sheetName, int startRowIndex, int startColumnIndex) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        Workbook workbook = getWorkbook(inputStream, extension);
        if (workbook != null) {
            Sheet sheet = workbook.getSheet(sheetName);
            // 获取总行数
            int rowNum = sheet.getPhysicalNumberOfRows();
            // 获取第一行的总列数
            int colNum = sheet.getRow(0).getPhysicalNumberOfCells();
            list = getSheetData(workbook, sheet, startRowIndex, rowNum, startColumnIndex, colNum);
        }
        return list;
    }

    /**
     * 对外提供读取excel的方法， 根据sheet名称读取sheet对象， 并指定行列区间获取数据[startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex)
     *
     * @param inputStream 文件输入流
     * @param extension 文件后缀
     * @param sheetName        表名
     * @param startRowIndex    起始行下标
     * @param endRowIndex      结束行下标 + 1
     * @param startColumnIndex 起始列下标
     * @param endColumnIndex   结束列下标 + 1
     * @return 文件表名为 sheetName 的 sheet 的区间 [startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex) 内所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> readExcel(InputStream inputStream, String extension, String sheetName, int startRowIndex, int endRowIndex,
                                               int startColumnIndex, int endColumnIndex) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        Workbook workbook = getWorkbook(inputStream, extension);
        if (workbook != null) {
            Sheet sheet = workbook.getSheet(sheetName);
            list = getSheetData(workbook, sheet, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }
        return list;
    }

    /**
     * 获取 workbook 的所有 sheet
     *
     * @param workbook 工作簿
     * @return sheet 列表
     */
    public static List<Sheet> getAllSheets(Workbook workbook) {
        int numOfSheets = workbook.getNumberOfSheets();
        List<Sheet> sheets = new ArrayList<>();
        for (int i = 0; i < numOfSheets; i++) {
            sheets.add(workbook.getSheetAt(i));
        }
        return sheets;
    }

    /**
     * 根据 输入流 和 其文件后缀 来获取 workbook
     *
     * @param inputStream 输入流
     * @param extension   文件后缀
     * @return workbook
     * @throws IOException IO 异常
     */
    public static Workbook getWorkbook(InputStream inputStream, String extension) throws IOException {
        Workbook workbook = null;
        if (inputStream != null) {
            // for office2003
            if ("xls".equals(extension)) {
                workbook = new HSSFWorkbook(inputStream);
            } // for office2007
            else if ("xlsx".equals(extension)) {
                workbook = new XSSFWorkbook(inputStream);
            } else {
                throw new IOException("不支持的文件类型");
            }
        }
        return workbook;
    }

    /**
     * 根据 excel文件 来获取workbook
     *
     * @param file 文件
     * @return workbook
     * @throws IOException IO 异常
     */
    public static Workbook getWorkbook(File file) throws IOException {
        if (file != null && file.exists() && file.isFile()) {
            // 获取文件后缀
            String fileName = file.getName();
            int lastIndex = fileName.lastIndexOf(".");
            String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

            return getWorkbook(new FileInputStream(file), extension);
        }
        return null;
    }

    /**
     * 根据excel文件来获取workbook
     *
     * @param filePath 文件路径
     * @return workbook
     * @throws IOException IO 异常
     */
    public static Workbook getWorkbook(String filePath) throws IOException {
        File file = new File(filePath);
        return getWorkbook(file);
    }

    /**
     * 根据excel文件输出路径来获取对应的workbook
     *
     * @param filePath 文件路径
     * @return workbook
     * @throws IOException IO 异常
     */
    public static Workbook getExportWorkbook(String filePath) throws IOException {
        Workbook workbook;
        File file = new File(filePath);

        // 获取文件后缀
        String fileName = file.getName();
        int lastIndex = fileName.lastIndexOf(".");
        String extension = lastIndex == -1 ? "" : fileName.substring(lastIndex + 1);

        // for 少量数据
        if ("xls".equals(extension)) {
            workbook = new HSSFWorkbook();
        } // for 大量数据
        else if ("xlsx".equals(extension)) {
            // 定义内存里一次只留5000行
            workbook = new SXSSFWorkbook(5000);
        } else {
            throw new IOException("不支持的文件类型");
        }
        return workbook;
    }

    /**
     * 获取 sheet 的所有数据
     *
     * @param workbook 工作簿
     * @param sheet    表
     * @return sheet 的所有数据
     */
    public static List<List<Object>> getSheetData(Workbook workbook, Sheet sheet) {
        List<List<Object>> list = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            // 整行都空，就跳过
            if (isBlankRow(workbook, row)) {
                continue;
            }
            List<Object> rowData = getRowData(workbook, row);
            list.add(rowData);
        }
        return list;
    }


    /**
     * 获取该 sheet 的指定行列的数据[startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex)
     *
     * @param workbook         工作簿
     * @param sheet            表
     * @param startRowIndex    开始行下标
     * @param endRowIndex      结束行下标 + 1
     * @param startColumnIndex 开始列下标
     * @param endColumnIndex   结束列下标 + 1
     * @return sheet 的区间 [startRowIndex, endRowIndex), [startColumnIndex, endColumnIndex) 内所有数据
     * @throws IOException IO 异常
     */
    public static List<List<Object>> getSheetData(Workbook workbook, Sheet sheet, int startRowIndex, int endRowIndex,
                                                  int startColumnIndex, int endColumnIndex) throws IOException {
        List<List<Object>> list = new ArrayList<>();
        if (startRowIndex > endRowIndex || startColumnIndex > endColumnIndex) {
            return list;
        }

        // 获取总行数
        int rowNum = sheet.getPhysicalNumberOfRows();
        // 第一行总列数
        int colNum = sheet.getRow(0).getPhysicalNumberOfCells();

        if (endRowIndex > rowNum) {
            throw new IOException("行的最大下标索引超过了该sheet实际总行数(包括标题行)" + rowNum);
        }
        if (endColumnIndex > colNum) {
            throw new IOException("列的最大下标索引超过了实际标题总列数" + colNum);
        }
        for (int i = startRowIndex; i < endRowIndex; i++) {
            Row row = sheet.getRow(i);
            // 整行都空，就跳过
            if (isBlankRow(workbook, row)) {
                continue;
            }
            List<Object> rowData = getRowData(workbook, row, startColumnIndex, endColumnIndex);
            list.add(rowData);
        }
        return list;
    }

    /**
     * 根据指定列区间获取行的数据
     *
     * @param workbook         工作簿
     * @param row              行
     * @param startColumnIndex 开始列下标
     * @param endColumnIndex   结束列下标 + 1
     * @return row 行 [startColumnIndex, endColumnIndex) 内所有数据
     */
    public static List<Object> getRowData(Workbook workbook, Row row, int startColumnIndex, int endColumnIndex) {
        List<Object> rowData = new ArrayList<>();
        for (int j = startColumnIndex; j < endColumnIndex; j++) {
            Cell cell = row.getCell(j);
            rowData.add(getCellValue(cell));
        }
        return rowData;
    }

    /**
     * 判断整行是不是都为空
     *
     * @param row 行
     * @return true：全为空；false：不全为空
     */
    public static boolean isBlankRow(Workbook workbook, Row row) {
        boolean allRowIsBlank = true;
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Object cellValue = getCellValue(cellIterator.next());
            if (cellValue != null && !"".equals(cellValue)) {
                allRowIsBlank = false;
                break;
            }
        }
        return allRowIsBlank;
    }

    /**
     * 获取行的数据
     *
     * @param workbook 工作簿
     * @param row 行
     * @return row 行的所有数据
     */
    public static List<Object> getRowData(Workbook workbook, Row row) {
        List<Object> rowData = new ArrayList<>();
        /**
         * 不建议用row.cellIterator(), 因为空列会被跳过， 后面的列会前移， 建议用for循环， row.getLastCellNum()是获取最后一个不为空的列是第几个
         * 结论：空行可以跳过， 空列最好不要跳过
         */
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            Object cellValue = getCellValue(cell);
            rowData.add(cellValue);
        }
        return rowData;
    }

    /**
     * 获取单元格值
     *
     * @param cell 单元格
     * @return 单元格值对应的 java对象
     */
    public static Object getCellValue(Cell cell) {
        if (cell == null || (CellType.STRING.equals(cell.getCellType()) && StringUtils.isBlank(cell.getStringCellValue()))) {
            return null;
        }

        // 格式化数字
        DecimalFormat decimalFormat = new DecimalFormat("0");

        switch (cell.getCellType()) {
            case BLANK:
                return null;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return cell.getErrorCellValue();
            case FORMULA:
                return cell.getNumericCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                    String value = decimalFormat.format(cell.getNumericCellValue());
                    if (StringUtils.isBlank(value)) {
                        return null;
                    }
                    return value;
                } else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                    String value = decimalFormat.format(cell.getNumericCellValue());
                    if (StringUtils.isBlank(value)) {
                        return null;
                    }
                    return value;
                } else {
                    return cell.getNumericCellValue();
                }
            case STRING:
                String value = cell.getStringCellValue();
                if (StringUtils.isBlank(value)) {
                    return null;
                } else {
                    return value;
                }
            default:
                return null;
        }
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于单个sheet
     *
     * @param <T>      数据类型
     * @param headers  表格属性列名数组
     * @param dataset  需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
     *                 javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param filePath excel文件输出路径
     */
    public static <T> void exportExcel(String[] headers, Collection<T> dataset, String filePath) {
        exportExcel(headers, dataset, filePath, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于单个sheet
     *
     * @param <T>
     * @param headers  表格属性列名数组
     * @param dataset  需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
     *                 javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param filePath excel文件输出路径
     * @param pattern  如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    public static <T> void exportExcel(String[] headers, Collection<T> dataset, String filePath, String pattern) {
        try {
            // 声明一个工作薄
            Workbook workbook = getExportWorkbook(filePath);
            if (workbook != null) {
                // 生成一个表格
                Sheet sheet = workbook.createSheet();

                write2Sheet(sheet, headers, dataset, pattern);
                OutputStream out = new FileOutputStream(new File(filePath));
                workbook.write(out);
                out.close();
            }
        } catch (IOException e) {
            LOGGER.error(e.toString(), e);
        }
    }

    /**
     * 导出数据到Excel文件
     *
     * @param dataList 要输出到Excel文件的数据集
     * @param filePath excel文件输出路径
     */
    public static void exportExcel(String[][] dataList, String filePath) {
        try {
            // 声明一个工作薄
            Workbook workbook = getExportWorkbook(filePath);
            if (workbook != null) {
                // 生成一个表格
                Sheet sheet = workbook.createSheet();

                for (int i = 0; i < dataList.length; i++) {
                    String[] r = dataList[i];
                    Row row = sheet.createRow(i);
                    for (int j = 0; j < r.length; j++) {
                        Cell cell = row.createCell(j);
                        // cell max length 32767
                        if (r[j].length() > 32767) {
                            LOGGER.warn("异常处理", "--此字段过长(超过32767),已被截断--" + r[j]);
                            r[j] = r[j].substring(0, 32766);
                        }
                        cell.setCellValue(r[j]);
                    }
                }
                // 自动列宽
                if (dataList.length > 0) {
                    int colCount = dataList[0].length;
                    for (int i = 0; i < colCount; i++) {
                        sheet.autoSizeColumn(i);
                    }
                }
                OutputStream out = new FileOutputStream(new File(filePath));
                workbook.write(out);
                out.close();
            }
        } catch (IOException e) {
            LOGGER.error("#exportExcel error.", e);
        }
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于多个sheet
     *
     * @param sheets   ExcelSheet的集体
     * @param filePath excel文件路径
     */
    public static <T> void exportExcel(List<ExcelSheet<T>> sheets, String filePath) {
        exportExcel(sheets, filePath, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于多个sheet
     *
     * @param sheets   ExcelSheet的集合
     * @param filePath excel文件输出路径
     * @param pattern  如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    public static <T> void exportExcel(List<ExcelSheet<T>> sheets, String filePath, String pattern) {
        if (CollectionUtils.isEmpty(sheets)) {
            return;
        }
        try {
            // 声明一个工作薄
            Workbook workbook = getExportWorkbook(filePath);
            if (workbook != null) {
                for (ExcelSheet<T> sheetInfo : sheets) {
                    // 生成一个表格
                    Sheet sheet = workbook.createSheet(sheetInfo.getSheetName());
                    write2Sheet(sheet, sheetInfo.getHeaders(), sheetInfo.getDataset(), pattern);
                }
                OutputStream out = new FileOutputStream(new File(filePath));
                workbook.write(out);
                out.close();
            }
        } catch (IOException e) {
            LOGGER.error("#exportExcel error.", e);
        }
    }

    /**
     * 每个sheet的写入
     *
     * @param sheet   页签
     * @param headers 表头
     * @param dataset 数据集合
     * @param pattern 日期格式
     */
    public static <T> void write2Sheet(Sheet sheet, String[] headers, Collection<T> dataset, String pattern) {
        // 产生表格标题行
        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
        }
        // 遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = 0;
        while (it.hasNext()) {
            index++;
            row = sheet.createRow(index);
            T t = it.next();
            // row data is map
            if (t instanceof Map) {
                @SuppressWarnings("unchecked")
                Map<String, Object> map = (Map<String, Object>) t;
                int cellNum = 0;
                for (String k : headers) {
                    if (!map.containsKey(k)) {
                        LOGGER.error("Map 中 不存在 key [" + k + "]");
                        continue;
                    }
                    Cell cell = row.createCell(cellNum);
                    Object value = map.get(k);
                    if (value == null) {
                        cell.setCellValue(StringUtils.EMPTY);
                    } else {
                        cell.setCellValue(String.valueOf(value));
                    }
                    cellNum++;
                }
            } // row data is Object[]
            else if (t instanceof Object[]) {
                Object[] tObjArr = (Object[]) t;
                for (int i = 0; i < tObjArr.length; i++) {
                    Cell cell = row.createCell(i);
                    Object value = tObjArr[i];
                    if (value == null) {
                        cell.setCellValue(StringUtils.EMPTY);
                    } else {
                        cell.setCellValue(String.valueOf(value));
                    }
                }
            } // row data is List
            else if (t instanceof List<?>) {
                List<?> rowData = (List<?>) t;
                for (int i = 0; i < rowData.size(); i++) {
                    Cell cell = row.createCell(i);
                    Object value = rowData.get(i);
                    if (value == null) {
                        cell.setCellValue(StringUtils.EMPTY);
                    } else {
                        cell.setCellValue(String.valueOf(value));
                    }
                }
            } // row data is vo
            else {
                // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
                Field[] fields = t.getClass().getDeclaredFields();
                for (int i = 0; i < fields.length; i++) {
                    Cell cell = row.createCell(i);
                    Field field = fields[i];
                    String fieldName = field.getName();
                    String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);

                    try {
                        Class<?> tClazz = t.getClass();
                        Method getMethod = tClazz.getMethod(getMethodName);
                        Object value = getMethod.invoke(t);
                        String textValue = null;
                        if (value instanceof Integer) {
                            int intValue = (Integer) value;
                            cell.setCellValue(intValue);
                        } else if (value instanceof Float) {
                            float fValue = (Float) value;
                            cell.setCellValue(fValue);
                        } else if (value instanceof Double) {
                            double dValue = (Double) value;
                            cell.setCellValue(dValue);
                        } else if (value instanceof Long) {
                            long longValue = (Long) value;
                            cell.setCellValue(longValue);
                        } else if (value instanceof Boolean) {
                            boolean bValue = (Boolean) value;
                            cell.setCellValue(bValue);
                        } else if (value instanceof Date) {
                            Date date = (Date) value;
                            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                            textValue = sdf.format(date);
                        } else {
                            // 其它数据类型都当作字符串简单处理
                            textValue = value.toString();
                        }
                        if (textValue != null) {
                            cell.setCellValue(textValue);
                        } else {
                            cell.setCellValue(StringUtils.EMPTY);
                        }
                    } catch (Exception e) {
                        LOGGER.error("#write2Sheet error.", e);
                    }

                }
            }
        }
        // 设定自动宽度
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * EXCEL文件下载
     *
     * @param filePath 文件路径
     * @param response 响应
     */
    public static void download(String filePath, HttpServletResponse response) {
        try {
            File file = new File(filePath);
            // 取得文件名
            String filename = file.getName();
            // 以流的形式下载文件
            InputStream fis = new BufferedInputStream(new FileInputStream(filePath));
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            fis.close();
            // 清空response
            response.reset();
            // 设置response的Header
            response.addHeader("Content-Disposition", "attachment;filename=" + new String(filename.getBytes()));
            response.addHeader("Content-Length", "" + file.length());
            OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
            response.setContentType("application/vnd.ms-excel;charset=gb2312");
            toClient.write(buffer);
            toClient.flush();
            toClient.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * 分别测试以下示例并成功通过：
     * 1. 单个sheet， 数据集类型是List<List<Object>>
     * 2. 单个sheet， 数据集类型是List<Object[]>
     * 3. 多个sheet， 数据集类型是List<ExcelSheet<List<Object>>>
     * 4. 多个sheet， 数据集类型是List<ExcelSheet<List<Object>>>
     * 5. 多个sheet， 数据集类型是List<ExcelSheet<List<Object>>>, 支持大数据量
     *
     * @param args 参数列表
     */
    public static void main(String[] args) {
        List<List<Object>> list = new ArrayList<>();

        try {

            list = readExcel("D:/test.xlsx");
             // 导入
            list = readExcel("D:/test.xlsx", 1);
            list = readExcel("D:/test.xlsx", 1);

            List<Object[]> dataList = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                Object[] objArr = new Object[list.get(i).size()];
                List<Object> objList = list.get(i);
                for (int j = 0; j < objList.size(); j++) {
                    objArr[j] = objList.get(j);
                }
                dataList.add(objArr);
            }
            for (int i = 0; i < dataList.size(); i++) {
            	System.out.println(Arrays.toString(dataList.get(i)));
            }

            String[] headers = { "代理商ID", "代理商编码", "系统内代理商名称", "贷款代理商名称", "入网时长", "佣金账期", "佣金类型", "金额" };

            ExcelSheet<List<Object>> sheetList = new ExcelSheet<>();
            sheetList.setHeaders(headers);
            sheetList.setSheetName("按入网时间提取佣金数");
            sheetList.setDataset(list);
            List<ExcelSheet<List<Object>>> sheetsList = new ArrayList<>();
            sheetsList.add(sheetList);

            ExcelSheet<Object[]> sheetArray = new ExcelSheet<>();
            sheetArray.setHeaders(headers);
            sheetArray.setSheetName("按入网时间提取佣金数");
            sheetArray.setDataset(dataList);
            List<ExcelSheet<Object[]>> sheetsArray = new ArrayList<>();
            sheetsArray.add(sheetArray);
            // 导出
            exportExcel(headers, list, "d://out_" + System.currentTimeMillis() + ".xlsx");
            exportExcel(headers, dataList, "d://out_" + System.currentTimeMillis() + ".xlsx");
            exportExcel(sheetsList, "d://out_" + System.currentTimeMillis() + ".xlsx");
            exportExcel(sheetsArray, "d://out_" + System.currentTimeMillis() + ".xlsx");
            list = readExcel("D:/test.xlsx", "按入网时间提取佣金数");
            list = readExcel("D:/test.xlsx", 0, 1, 85, 0, 6);
            list = readExcel("D:/test.xlsx", "按入网时间提取佣金数", 1061, 1062, 0, 8);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 有3个sheet的数据， 每个sheet数据为50万行， 共150万行数据输出到Excel文件, 性能测试。
        List<ExcelSheet<List<Object>>> sheetsData = new ArrayList<>();

        int sheetRowNum = 50000;

        for (int i = 0; i < 3; i++) {
            ExcelSheet<List<Object>> sheetData = new ExcelSheet<>();
            String[] headers = {"姓名", "手机号码", "性别", "身份证号码", "家庭住址"};
            String sheetName = "第" + (i + 1) + "个sheet";

            List<List<Object>> sheetDataList = new ArrayList<>();
            for (int j = 0; j < sheetRowNum; j++) {
                List<Object> rowData = new ArrayList<>();
                rowData.add("小明");
                rowData.add("18888888888");
                rowData.add("男");
                rowData.add("123123123123123123");
                rowData.add("广州市");
                sheetDataList.add(rowData);
            }
            sheetData.setSheetName(sheetName);
            sheetData.setHeaders(headers);
            sheetData.setDataset(sheetDataList);

            sheetsData.add(sheetData);
        }
        String filePath = "d://out_" + System.currentTimeMillis() + ".xlsx";
        exportExcel(sheetsData, filePath);
        System.out.println("-----end-----");
        for (int i = 0; i < list.size(); i++) {
        	System.out.println(list.get(i));
        }
    }
}
