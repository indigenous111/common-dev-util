package in.indigenous.common.util.io.excel;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import in.indigenous.common.util.io.FileReader;

@Component
public class XLSXFileReader implements FileReader {

	@Override
	public Map<Object, List<Object>> read(String dir, String fileName) {
		String filePath = dir + File.separator + fileName;
		File file = new File(filePath);
		try (FileInputStream fis = new FileInputStream(file)) {
			XSSFWorkbook workBook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workBook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			List<Object> headers = new ArrayList<>();
			Map<Object, List<Object>> data = new HashMap<>();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
					case STRING:
						if (cell.getRowIndex() == 0) {
							headers.add(cell.getStringCellValue());
						} else {
							Object header = headers.get(cell.getColumnIndex());
							Object cellData = cell.getStringCellValue();
							handleData(data, header, cellData);
						}
						break;
					case NUMERIC:
						if (cell.getRowIndex() == 0) {
							headers.add(cell.getNumericCellValue());
						} else {
							Object header = headers.get(cell.getColumnIndex());
							Object cellData = cell.getNumericCellValue();
							handleData(data, header, cellData);
						}
						break;
					case BOOLEAN:
						if (cell.getRowIndex() == 0) {
							headers.add(cell.getBooleanCellValue());
						} else {
							Object header = headers.get(cell.getColumnIndex());
							Object cellData = cell.getBooleanCellValue();
							handleData(data, header, cellData);
						}
						break;
					case BLANK:
						if (cell.getRowIndex() == 0) {
							headers.add(null);
						} else {
							Object header = headers.get(cell.getColumnIndex());
							handleData(data, header, null);
						}
						break;
					default:
					}
				}
			}
			return data;
		} catch (Throwable t) {
			throw new RuntimeException(t);
		}
	}

	private void handleData(Map<Object, List<Object>> dataMap, Object header, Object data) {
		List<Object> list = dataMap.get(header);
		if (list == null) {
			list = new ArrayList<>();
			dataMap.put(header, list);
		}
		list.add(data);
	}
}
