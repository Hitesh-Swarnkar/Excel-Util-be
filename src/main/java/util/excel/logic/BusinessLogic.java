package util.excel.logic;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import util.excel.model.SourceSheetData;

public class BusinessLogic {

	public void copySheetData(String sourceFile, String destFile) throws Exception {
		List<SourceSheetData> sourceWBList = new ArrayList<SourceSheetData>();
		//Read source file
		Workbook sourceWB = new XSSFWorkbook(new FileInputStream(new File(sourceFile)));

		Sheet sourceSheet = sourceWB.getSheetAt(0);

		for(Row row : sourceSheet) {
			if(row.getRowNum()==0)
				continue;
			SourceSheetData sheetModel = new SourceSheetData();
			sheetModel.setServiceName(row.getCell(0).toString());
			sheetModel.setCostCenter(row.getCell(1).toString());
			sheetModel.setIccCode(row.getCell(2).toString());
			sheetModel.setCost((Double)row.getCell(5).getNumericCellValue());
			sheetModel.setServiceOwner(row.getCell(6).toString());
			sourceWBList.add(sheetModel);
		}
		//close the file 
		sourceWB.close();
		
		for(SourceSheetData data : sourceWBList) {
			//check for the condition 1) check if Cost centre starts with A else throw an exception
			if(!data.getCostCenter().startsWith("A")) throw new Exception("Cost Center in SOURCE file doesn't start with A"); //check 1
		}
			

		int positionToUpdate =6;
		
		for(SourceSheetData data :  sourceWBList) {
			//read destination file
			FileInputStream destFileStream =  new FileInputStream(new File(destFile));
			Workbook destWB = new XSSFWorkbook(destFileStream);
			Sheet destSheet =  destWB.getSheetAt(0);
			
			for(Row row : destSheet) {
				if(row.getRowNum()<positionToUpdate)
					continue;
				
				if(row.getCell(0).getStringCellValue().startsWith("D")) {//check 2 for Action
					row.getCell(1).setCellValue(data.getIccCode());
					row.getCell(2).setCellValue(data.getServiceName());
					row.getCell(4).setCellValue(data.getCostCenter());
					row.getCell(5).setCellValue(data.getServiceName()+" for #");
					row.getCell(6).setCellValue(data.getCost());
					row.getCell(9).setCellValue(data.getServiceOwner());
					row.getCell(10).setCellValue(data.getServiceOwner());
				}else if(row.getCell(0).getStringCellValue().startsWith( "C")) {
					row.getCell(5).setCellValue(data.getServiceName()+" for #");
					row.getCell(6).setCellFormula(row.getCell(6).getCellFormula()); 
				}
			}
			//copy to destination file
			FileOutputStream destFileOutStream =  new FileOutputStream(new File(destFile));
			destWB.write(destFileOutStream);
			//close the file
			destFileOutStream.close();
			destWB.close();
			destFileStream.close();
			positionToUpdate=positionToUpdate+2;
		}	
	}
}
