package com.annotation.excel.utils;

import com.annotation.excel.annotation.Excel;
import com.annotation.excel.constant.FileConstants;
import com.annotation.excel.po.ExcelExportPo;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel 工具类
 *
 * @author : tdl
 * @date : 2019/3/25 下午6:42
 */
public class ExcelUtil {

    private Workbook workbook;

    private Workbook getWorkbook() {
        return workbook;
    }

    private OutputStream os;

    private String pattern;

    private OutputStream getOs() {
        return os;
    }

    private void setOs(OutputStream os) {
        this.os = os;
    }

    private String getPattern() {
        return pattern;
    }

    private void setPattern(String pattern) {
        this.pattern = pattern;
    }

    private ExcelUtil(Workbook workboook) {
        this.workbook = workboook;
    }

    private void createSheet() {
        workbook.createSheet();
    }

    /**
     * 设置sheet名称，长度为1-31，不能包含后面任一字符: ：\ / ? * [ ]
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param name       sheet名称
     * @return 操作结果
     * @throws IOException
     */
    private boolean setSheetName(int sheetIndex, String name) throws IOException {
        workbook.setSheetName(sheetIndex, name);
        return true;
    }

    /**
     * 获取 sheet名称
     *
     * @param sheetIx 指定 Sheet 页，从 0 开始
     * @return
     * @throws IOException
     */
    public String getSheetName(int sheetIx) throws IOException {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        return sheet.getSheetName();
    }

    /**
     * 关闭流
     *
     * @throws IOException
     */
    private void close() throws IOException {
        if (os != null) {
            os.close();
        }
        workbook.close();
    }

    /**
     * 返回sheet 中的行数
     *
     * @param sheetIx 指定 Sheet 页，从 0 开始
     * @return
     */
    public int getRowCount(int sheetIx) {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return 0;
        }
        return sheet.getLastRowNum() + 1;

    }

    private boolean write(List<List<Object>> rowData, List<String> cellTypes) throws IOException {
        return write(0, rowData, 0, cellTypes);
    }

    private boolean write(int sheetIx, List<List<Object>> rowData, int startRow, List<String> cellTypes) throws IOException {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        sheet.createFreezePane(0, 1, 0, 1);
        int dataSize = rowData.size();
        // 如果小于等于0，则一行都不存在
        if (getRowCount(sheetIx) > 0) {
            sheet.shiftRows(startRow, getRowCount(sheetIx), dataSize);
        }

        CellStyle cellStyle = getWorkbook().createCellStyle();
        for (int i = 0; i < dataSize; i++) {
            Row row = sheet.createRow(i + startRow);
            for (int j = 0; j < rowData.get(i).size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellType(CellType.STRING);
                if (StringUtil.isBlank(StringUtil.getObjectToString(rowData.get(i).get(j)))) {
                    cell.setCellType(CellType.BLANK);
                } else {
                    // 根据传入的单元格类型进行设置
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    if (i > 0 && "money".equals(cellTypes.get(j))) {
                        HSSFWorkbook workbook = new HSSFWorkbook();
                        HSSFDataFormat df = workbook.createDataFormat();
                        cellStyle.setDataFormat(df.getFormat("0.00"));
                        // 过滤空值
                        if (rowData.get(i).get(j) == null || rowData.get(i).get(j) == "null") {
                            cell.setCellStyle(cellStyle);
                            cell.setCellValue(0);
                            continue;
                        }
                        try {
                            Double number = Double.parseDouble((String) rowData.get(i).get(j));
                            cell.setCellStyle(cellStyle);
                            cell.setCellValue(number);
                        } catch (NumberFormatException e) {
                            cell.setCellStyle(cellStyle);
                            cell.setCellValue(0);
                        }
                    } else {
                        cell.setCellStyle(cellStyle);
                        cell.setCellValue(rowData.get(i).get(j) + "");
                    }
                }
            }
        }
        return true;
    }

    /**
     * 将数据写入到 Excel 指定 Sheet 页指定开始行中,指定行后面数据向后移动
     *
     * @param rowData  数据
     * @param sheetIx  指定 Sheet 页，从 0 开始
     * @param startRow 指定开始行，从 0 开始
     * @return true | false
     * @throws IOException
     */
    public boolean write(int sheetIx, List<List<Object>> rowData, int startRow) throws IOException {
        Sheet sheet = workbook.getSheetAt(sheetIx);
        int dataSize = rowData.size();
        // 如果小于等于0，则一行都不存在
        if (getRowCount(sheetIx) > 0) {
            sheet.shiftRows(startRow, getRowCount(sheetIx), dataSize);
        }

        for (int i = 0; i < dataSize; i++) {
            Row row = sheet.createRow(i + startRow);
            for (int j = 0; j < rowData.get(i).size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellType(CellType.STRING);
                if (StringUtil.isBlank(StringUtil.getObjectToString(rowData.get(i).get(j)))) {
                    cell.setCellType(CellType.BLANK);
                } else {
                    cell.setCellValue(rowData.get(i).get(j) + "");
                }
            }
        }
        return true;
    }

    private boolean createExcel(Workbook workbook, OutputStream outputStream) {
        try {
            workbook.write(outputStream);
            outputStream.flush();
            close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return true;
    }

    /**
     * 创建目录
     *
     * @param filePath 目录名
     */
    private static void checkDirExists(String filePath) {
        File file = new File(filePath);
        // 创建目录
        if (!file.exists()) {
            file.mkdir();
        }
    }

    /**
     * 导出excel
     *
     * @param excelExportPo excelExportPo
     * @throws Exception
     * @throws IOException
     */
    public static void exportExcel(ExcelExportPo excelExportPo) throws Exception, IOException {
        HttpServletRequest request = excelExportPo.getRequest();
        HttpServletResponse response = excelExportPo.getResponse();
        String excelName = excelExportPo.getExcelName();
        List<Object> list = excelExportPo.getList();

        String filePath = request.getSession().getServletContext().getRealPath("/") + FileConstants.EXPORT_CACHE_FILE_PATH;
        String fileName = excelName + FileConstants.EXPORT_FORMAT;
        try {
            checkDirExists(filePath);
            filePath = filePath + "/" + fileName;
            File file = new File(filePath);
            if (!file.exists()) {
                // 创建excel文档
                Workbook workbook = new XSSFWorkbook();
                ExcelUtil excelUtil = new ExcelUtil(workbook);
                excelUtil.createSheet();
                excelUtil.setSheetName(0, excelName);

                List<List<Object>> rowData = new ArrayList<>();
                // 单元格宽度
                Sheet sheet = workbook.getSheetAt(0);
                // 单元格数据类型
                List<String> cellTypeList = new ArrayList<>();
                // 解析excel数据
                for (Integer i = 0; i < list.size(); i++) {
                    Object object = list.get(0);
                    Class<?> cls = object.getClass();
                    Field[] fieldArray = cls.getDeclaredFields();
                    if (i == 0) {
                        List<Object> titleList = new ArrayList<>();
                        for (Integer index = 0; index < fieldArray.length; index++) {
                            Field field = fieldArray[index];
                            Excel column = field.getAnnotation(Excel.class);
                            if (column != null) {
                                String columnTitle = column.title();
                                String columnValueType = column.valueType();
                                Integer columnWidth = column.width();
                                titleList.add(columnTitle);
                                // 设置单元格宽度
                                sheet.setColumnWidth(index, columnWidth);
                                // 设置单元格数据类型
                                cellTypeList.add(columnValueType);
                            }
                        }
                        // 插入标题栏数据
                        rowData.add(titleList);
                    }

                    List<Object> valueList = new ArrayList<>();
                    for (Field field : fieldArray) {
                        Excel column = field.getAnnotation(Excel.class);
                        field.setAccessible(true);
                        if (column != null) {
                            valueList.add(field.get(object));
                        }
                    }
                    rowData.add(valueList);
                }

                // 写入数据
                excelUtil.write(rowData, cellTypeList);
                // 设置文件流
                excelUtil.setOs(new FileOutputStream(filePath));
                // 创建Excel文件
                excelUtil.createExcel(workbook, excelUtil.getOs());
            }

            // 执行下载
            doDownload(response, filePath, fileName, file);
        } catch (IOException ie) {
            throw ie;
        } catch (Exception e) {
            throw e;
        }
    }

    private static void doDownload(HttpServletResponse response, String filePath, String fileName, File file) {
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(filePath);
            // 响应下载请求
            DownloadUtil.downloadFile(response, inputStream, fileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (file != null && file.exists()) {
                file.delete();
            }
        }
    }
}