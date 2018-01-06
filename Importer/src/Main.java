import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;


public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // XSSFWorkbook, File
        OPCPackage pkg = OPCPackage.open(new File("C:\\Users\\kimwa\\OneDrive\\Documents\\Codes\\Matlab\\Data US\\US_Data_P1_1.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(pkg);
        Sheet sheet = wb.getSheetAt(0);
        CompanyData comp = new CompanyData();
        GlobalData globalData = new GlobalData();


        //Daten Import
        //========================================================================================================================================
        comp.nameConstructor(sheet.getRow(0).getCell(0).getStringCellValue());

        for(int i=2;i<sheet.getLastRowNum()+1;i++) {
            comp.dateAdd(sheet.getRow(i).getCell(0).getDateCellValue());
            globalData.dateAdd(sheet.getRow(i).getCell(0).getDateCellValue());
        }

        //Add Company specific Data
        comp.addReturns(sheet, "Mid");
        comp.addReturns(sheet, "Bid");
        comp.addReturns(sheet, "Ask");
        comp.addMarketData(sheet,"MktCap");
        comp.addMarketData(sheet, "Volume");

        //Global Data
        globalData.addMarketData(sheet, "S&P");
        globalData.addMarketData(sheet, "Russel");
        globalData.addMarketData(sheet, "RiskFree");

        //comp.returnsOut();
        comp.dataOut();
    }
}

