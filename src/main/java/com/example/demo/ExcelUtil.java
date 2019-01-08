package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 操作excel的工具类
 * 所需jar包
 * 	poi-excelant-3.8-20120326.jar
 *	xmlbeans-2.3.0.jar
 *	stax-api-1.0.1.jar
 *	dom4j-1.6.1.jar
 *	poi-3.8-20120326.jar
 *	poi-ooxml-3.8-20120326.jar
 *	poi-ooxml-schemas-3.8-20120326.jar
 */
public class ExcelUtil {

	/**
	 * 导出Excel
	 * @param filePath	文件路径
	 * @param fileName	文件名称
	 * @param sheetName	sheet名称
	 * @param headName	标题名称
	 * @param colName	列名称
	 * @param colValue	列值
	 * @return
	 * @throws Exception
	 */
	public static Workbook createExcel(String filePath,String fileName,String sheetName,String headName,String[] colName,List<String[]> colValue) throws Exception {
		String outfile =  filePath+"/"+fileName+".xlsx";
		FileOutputStream fileOutputStream =  new FileOutputStream(outfile);
		Workbook wb = null; //创建一个工作区
		
		try {
			//生成xlsx
			wb = new XSSFWorkbook();
			//生成xls
			//wb = new HSSFWorkbook();
			
			//创建工作表
            Sheet sheet = wb.createSheet(sheetName);
            sheet.setHorizontallyCenter(true);	//设置水平居中
            sheet.setDefaultColumnWidth(15);	// 设置表格默认列宽度为15个字节

            /*******************第一行设置标题**************************/
            //表头样式
            XSSFCellStyle headStyle = headStyle(wb);
            //表头：合并单元格
            sheet.addMergedRegion(new CellRangeAddress(0,0,0,colName.length-1)); //合并单元格
            
            //索引为0的地方创建标题行
            XSSFRow headRow = (XSSFRow) sheet.createRow(0);
            headRow.setHeightInPoints(20);	//设置行高
            XSSFCell headCell = headRow.createCell(0);
            headCell.setCellType(XSSFCell.CELL_TYPE_STRING);
            headCell.setCellValue(headName);
            headCell.setCellStyle(headStyle);
            
            /*******************第二行设置字段名称**************************/
            XSSFRow nameRow = (XSSFRow) sheet.createRow(1);
            //正文样式
            XSSFCellStyle nameStyle = nameStyle(wb);
            for (int i=0;i<colName.length;i++) {
            	XSSFCell nameCell = nameRow.createCell(i);
            	nameCell.setCellValue(colName[i]);
            	nameCell.setCellStyle(nameStyle);
            }
            
            /*******************第二行设置字段值**************************/ 
            //正文样式
            XSSFCellStyle valueStyle = valueStyle(wb);
            for (int j=0;j<colValue.size();j++) {
            	XSSFRow valueRow = (XSSFRow) sheet.createRow(2+j);
            	for (int i=0;i<colValue.get(j).length;i++) {
                	XSSFCell valueCell = valueRow.createCell(i);
                	valueCell.setCellValue(colValue.get(j)[i]);
                	valueCell.setCellStyle(valueStyle);
                }
                
            }
			
			wb.write(fileOutputStream);
		} catch (Exception e) {
			// TODO: handle exception
		} finally {
			fileOutputStream.flush();
			fileOutputStream.close();
		}
		return wb;
	}
	
	/**
	 * 设置表头的字体样式
	 * @param wb
	 * @return
	 */
	public static XSSFCellStyle headStyle(Workbook wb) {
		XSSFCellStyle headStyle = (XSSFCellStyle) wb.createCellStyle();	//设置表头样式
        headStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //设置字体
        XSSFFont font = (XSSFFont)wb.createFont();
        font.setFontName("新宋体");
        font.setFontHeightInPoints((short)16);	//设置字体高度
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD); //粗体
        headStyle.setFont(font);
        headStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);	//设置垂直居中
        return headStyle;
	}
	
	/**
	 * 设置name的字体样式
	 * @param wb
	 * @return
	 */
	public static XSSFCellStyle nameStyle(Workbook wb) {
		XSSFCellStyle headStyle = (XSSFCellStyle) wb.createCellStyle();	//设置表头样式
        headStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //设置字体
        XSSFFont font = (XSSFFont)wb.createFont();
        font.setFontName("新宋体");
        font.setFontHeightInPoints((short)12);	//设置字体高度
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD); //粗体
        headStyle.setFont(font);
        headStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);	//设置垂直居中
        return headStyle;
	}
	
	/**
	 * 设置value的字体样式
	 * @param wb
	 * @return
	 */
	public static XSSFCellStyle valueStyle(Workbook wb) {
		XSSFCellStyle headStyle = (XSSFCellStyle) wb.createCellStyle();	//设置表头样式
        headStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //设置字体
        XSSFFont font = (XSSFFont)wb.createFont();
        font.setFontName("新宋体");
        font.setFontHeightInPoints((short)12);	//设置字体高度
        headStyle.setFont(font);
        headStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);	//设置垂直居中
        return headStyle;
	}
	
	/**
	 * 读取【导入】excel信息
	 * @param file  读取的文件
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("resource")
	public static List<String[]> readExcel(File file) throws Exception {
		List<String[]> str = new ArrayList<String[]>();
		String fileName = file.getName();
		//获取文件后缀名并转成小写
		String ext=fileName.substring(fileName.lastIndexOf(".")+1).toLowerCase();
		//读入数据
		FileInputStream fs = new FileInputStream(file);
		List<String[]> dataList = new ArrayList<String[]>();
		if (ext.equals("xls")) {
			//读取2003
			HSSFWorkbook wb = new HSSFWorkbook(fs); // 获得工作
			HSSFSheet sheet = wb.getSheetAt(0); // 拿到第一个sheet页
			Iterator<Row> rows = sheet.rowIterator(); // 拿到第一行
			while (rows.hasNext()) { // 如果有值
				HSSFRow row = (HSSFRow) rows.next(); // 拿到当前行
				if(row.getRowNum() != 0) {  //不读取标题(可以根据自己需求进行修改)
					String[] data = new String[row.getLastCellNum()];
					Iterator<Cell> cells = row.cellIterator(); // 拿到当前行所有列的集合
					while (cells.hasNext()) {
						HSSFCell cell = (HSSFCell) cells.next(); // 拿到列值
						String cellValue = ""; // 存放单元格的值
						switch (cell.getCellType()) { // 判断单元格的类型，取出单元格的值
						case HSSFCell.CELL_TYPE_NUMERIC:
							// 处理数字类型 去掉科学计数法格式
							double strCell = cell.getNumericCellValue();
							DecimalFormat formatCell = (DecimalFormat) NumberFormat.getPercentInstance();
							formatCell.applyPattern("0");
							String value = formatCell.format(strCell);
							if (Double.parseDouble(value) != strCell) {
								formatCell.applyPattern(Double.toString(strCell));
								value = formatCell.format(strCell);
							}
							cellValue = value;
							break;
						case HSSFCell.CELL_TYPE_STRING:
							cellValue = cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_BOOLEAN:
							cellValue = String.valueOf(cell.getBooleanCellValue());
							break;
						case HSSFCell.CELL_TYPE_FORMULA:
							cellValue = cell.getCellFormula();
							break;
						default:
							break;
						}
						data[cell.getColumnIndex()] = cellValue;
					}
					dataList.add(data);
				}
			}
		} else if (ext.equals("xlsx")) {
			//读取2007、2010
			XSSFWorkbook wb = new XSSFWorkbook(fs); // 获得工作
			XSSFSheet sheet = wb.getSheetAt(0); // 拿到第一个sheet页
			Iterator<Row> rows = sheet.rowIterator(); // 拿到第一行
			while (rows.hasNext()) { // 如果有值
				XSSFRow row = (XSSFRow) rows.next(); // 拿到当前行
				if(row.getRowNum() != 0) {  //不读取标题(可以根据自己需求进行修改)
					String[] data = new String[row.getLastCellNum()];
					Iterator<Cell> cells = row.cellIterator(); // 拿到当前行所有列的集合
					while (cells.hasNext()) {
						XSSFCell cell = (XSSFCell) cells.next(); // 拿到列值
						String cellValue = ""; // 存放单元格的值
						switch (cell.getCellType()) { // 判断单元格的类型，取出单元格的值
						case XSSFCell.CELL_TYPE_NUMERIC:
							// 处理数字类型 去掉科学计数法格式
							double strCell = cell.getNumericCellValue();
							BigDecimal b = new BigDecimal(Double.toString(strCell));
							DecimalFormat formatCell = (DecimalFormat) NumberFormat.getPercentInstance();
							formatCell.applyPattern("0");
							String value = formatCell.format(strCell);
							if(cell.getCellStyle().getDataFormatString().indexOf("%")!=-1){
								value = cell.getNumericCellValue()*100+"%";
							}else{
								if (Double.parseDouble(value) != strCell) {
									//formatCell.applyPattern(Double.toString(strCell));
									value = b.toString();
								}
							}
							cellValue = value;
							break;
						case XSSFCell.CELL_TYPE_STRING:
							cellValue = cell.getStringCellValue();
							break;
						case XSSFCell.CELL_TYPE_BOOLEAN:
							cellValue = String.valueOf(cell.getBooleanCellValue());
							break;
						case XSSFCell.CELL_TYPE_FORMULA:
							cellValue = cell.getCellFormula();
							break;
						default:
							break;
						}
						data[cell.getColumnIndex()] = cellValue;
					}
					dataList.add(data);
				}
			}
		} else {
			throw new Exception("什么破文件连后缀名都没有！");
		}
		return dataList;
	}
	public static List<String> readExcelSheetName(File file) throws Exception {
		String fileName = file.getName();
		//获取文件后缀名并转成小写
		String ext=fileName.substring(fileName.lastIndexOf(".")+1).toLowerCase();
		//读入数据
		FileInputStream fs = new FileInputStream(file);
		List<String> dataList = new ArrayList<String>();
		if (ext.equals("xls")) {
			HSSFWorkbook wb = new HSSFWorkbook(fs); // 获得工作
			HSSFSheet sheet = wb.getSheetAt(0); // 拿到第一个sheet页
			String data = sheet.getSheetName();
			dataList.add(data);
		}else if (ext.equals("xlsx")) {
			//读取2007、2010
			XSSFWorkbook wb = new XSSFWorkbook(fs); // 获得工作
			XSSFSheet sheet = wb.getSheetAt(0); // 拿到第一个sheet页
			String data = sheet.getSheetName();
			dataList.add(data);
		}
		return dataList;
	}
	
	
	public static void main(String[] args) {
		try {
			//文件导出
			/*String[] colName = {"列1name","列2name","列3name","列4name","列5name","列6name","列7name"};
			String[] colValue = {"1Value","2Value","3Value","4Value","5Value","6Value","7Value"};
			String[] colValue1 = {"19Value","29Value","39Value","49Value","59Value","69Value","列79Value"};
			List<String[]> colValueList = new ArrayList<String[]>();
			colValueList.add(colValue);
			colValueList.add(colValue1);
			colValueList.add(colValue);
			colValueList.add(colValue);
			colValueList.add(colValue);
			colValueList.add(colValue1);
			colValueList.add(colValue);
			colValueList.add(colValue);
			ExcelUtil.createExcel("C:/TEMP", "测试", "测试1", "标题测试", colName, colValueList);
			*/
			//文件读取
			File file = new File("C:/TEMP/导入2003.xlsx");
			List<String[]> list = ExcelUtil.readExcel(file);
			for (int i=0;i<list.size();i++) {
				for (int j=0;j<list.get(i).length;j++) {
					System.out.print(list.get(i)[j]+",");
				}
				System.out.println();
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}