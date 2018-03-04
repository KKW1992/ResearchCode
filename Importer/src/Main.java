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
        CompanyData comp = new CompanyData(sheet);
        GlobalData globalData = new GlobalData(sheet);


        //Daten Import
        //========================================================================================================================================

        for(int i=1;i<sheet.getLastRowNum()+1;i++) {
            comp.dateAdd(sheet.getRow(i).getCell(0).getDateCellValue());
            globalData.dateAdd(sheet.getRow(i).getCell(0).getDateCellValue());
        }

        //Add Company specific Returns
        comp.addReturns("Mid");
        comp.addReturns( "Bid");
        comp.addReturns( "Ask");

        //Add Company specific Market Data
        comp.addMarketData("MktCap");
        comp.addMarketData( "Volume");

        //Global Data
        globalData.addMarketData("S&P");
        globalData.addMarketData( "Russel");
        globalData.addMarketData( "RiskFree");

        //Company Data
        comp.addCompanyData( "PB");
        comp.addCompanyData( "PE");
        comp.addCompanyData( "PS");


        //Lagged Returns
        comp.addLaggedReturns( "oneWeek");
        comp.addLaggedReturns( "oneMonth");
        comp.addLaggedReturns( "threeMonth");
        comp.addLaggedReturns( "sixMonth");

        //Volatility
        comp.addVolatility( "oneWeek");
        comp.addVolatility( "oneMonth");
        comp.addVolatility( "threeMonth");
        comp.addVolatility( "sixMonth");

        comp.writeData();
    }
}

