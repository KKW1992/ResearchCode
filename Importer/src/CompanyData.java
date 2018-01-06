import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

public class CompanyData {
    //Returns
    ArrayList<String> oneDay_midReturn = new ArrayList<>();
    ArrayList<String> midPrice = new ArrayList<>();
    ArrayList<String> oneDay_bidReturn = new ArrayList<>();
    ArrayList<String> oneDay_askReturn = new ArrayList<>();

    //Market Data
    ArrayList<String> mktCap = new ArrayList<>();
    ArrayList<String> volume = new ArrayList<>();

    //Company specific Data
    ArrayList<String> PB = new ArrayList<>();
    ArrayList<String> PE = new ArrayList<>();
    ArrayList<String> PS = new ArrayList<>();

    ArrayList<LocalDate> datum = new ArrayList<>();
    String name;

    void nameConstructor(String compName){
        name = compName;
    }

    void addReturns(Sheet sheet, String returnType){
        double priceOne = 0;
        double priceTwo = 0;
        boolean tester;
        int x = 1;

        switch (returnType){
            case "Mid":
                tester = (sheet.getRow(1).getCell(1).getCellTypeEnum()== CellType.NUMERIC && sheet.getRow(1).getCell(1).getNumericCellValue()!=0);
                if(tester){
                    priceOne = sheet.getRow(1).getCell(1).getNumericCellValue();
                }else{
                    do {
                        midPrice.add("na");
                        oneDay_midReturn.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(1).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(1).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(1).getNumericCellValue()==0)));
                    priceOne = sheet.getRow(x).getCell(1).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(1).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(1).getNumericCellValue() != 0);

                    if (tester) {
                        priceTwo = sheet.getRow(i).getCell(1).getNumericCellValue();

                        this.midReturnCalc(priceOne, priceTwo);
                        midPrice.add(Double.toString(priceOne));

                        priceOne = sheet.getRow(i).getCell(1).getNumericCellValue();
                    }else{
                        oneDay_midReturn.add("na");
                    }
                }
                midPrice.add(Double.toString(priceOne));
                break;
            case "Bid":
                tester = (sheet.getRow(1).getCell(2).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(2).getNumericCellValue()!=0);
                if(tester){
                    priceOne = sheet.getRow(1).getCell(2).getNumericCellValue();
                }else{
                    do {
                        oneDay_bidReturn.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(2).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(2).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(2).getNumericCellValue()==0)));
                    priceOne = sheet.getRow(x).getCell(2).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(2).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(2).getNumericCellValue() != 0);

                    if (tester) {
                        priceTwo = sheet.getRow(i).getCell(2).getNumericCellValue();

                        this.bidReturnCalc(priceOne, priceTwo);

                        priceOne = sheet.getRow(i).getCell(2).getNumericCellValue();
                    }else{
                        oneDay_bidReturn.add("na");
                    }
                }
                break;
            case "Ask":
                tester = (sheet.getRow(1).getCell(3).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(3).getNumericCellValue()!=0);
                if(tester){
                    priceOne = sheet.getRow(1).getCell(3).getNumericCellValue();
                }else{
                    do {
                        oneDay_bidReturn.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(3).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(3).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(3).getNumericCellValue()==0)));
                    priceOne = sheet.getRow(x).getCell(3).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(3).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(3).getNumericCellValue() != 0);

                    if (tester) {
                        priceTwo = sheet.getRow(i).getCell(3).getNumericCellValue();

                        this.askReturnCalc(priceOne, priceTwo);

                        priceOne = sheet.getRow(i).getCell(3).getNumericCellValue();
                    }else{
                        oneDay_askReturn.add("na");
                    }
                }
                break;
            default:
                System.out.println("Please specify the Type of Return (Bid, Mid, Ask)");
                break;
        }
    }

    private void midReturnCalc(double firstPrice, double secondPrice){
        oneDay_midReturn.add(Double.toString(Math.log(secondPrice/firstPrice)));
    }

    private void bidReturnCalc(double firstPrice, double secondPrice){
        oneDay_bidReturn.add(Double.toString(Math.log(secondPrice/firstPrice)));
    }

    private void askReturnCalc(double firstPrice, double secondPrice){
        oneDay_askReturn.add(Double.toString(Math.log(secondPrice/firstPrice)));
    }

    void addMarketData(Sheet sheet, String dataType){
        double valueOne = 0;
        double valueTwo = 0;
        boolean tester;
        int x = 1;

        switch (dataType){
            case "MktCap":
                tester = (sheet.getRow(1).getCell(4).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(4).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(4).getNumericCellValue();
                }else{
                    do {
                        mktCap.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(4).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(4).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(4).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(4).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(4).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(4).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(4).getNumericCellValue();

                        mktCap.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(4).getNumericCellValue();
                    }else{
                        mktCap.add("na");
                    }
                }
                break;
            case "Volume":
                tester = (sheet.getRow(1).getCell(11).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(11).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(11).getNumericCellValue();
                }else{
                    do {
                        volume.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(11).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(11).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(11).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(11).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=2;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(11).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(11).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(11).getNumericCellValue();

                        volume.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(11).getNumericCellValue();
                    }else{
                        mktCap.add("na");
                    }
                }
                break;
            default:
                System.out.println("Please specify the Type of Return (MktCap, Volume)");
                break;
        }
    }

    void addCompanyData(Sheet sheet, String dataType){
        double valueOne = 0;
        double valueTwo = 0;
        boolean tester;
        int x = 1;

        switch (dataType){
            case "PB":
                tester = (sheet.getRow(1).getCell(8).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(8).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(8).getNumericCellValue();
                }else{
                    do {
                        PB.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(8).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(8).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(8).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(8).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(8).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(8).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(8).getNumericCellValue();

                        PB.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(8).getNumericCellValue();
                    }else{
                        PB.add("na");
                    }
                }
                break;
            case "PE":
                tester = (sheet.getRow(1).getCell(6).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(6).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(6).getNumericCellValue();
                }else{
                    do {
                        PE.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(6).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(6).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(6).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(6).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(6).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(6).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(6).getNumericCellValue();

                        PE.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(6).getNumericCellValue();
                    }else{
                        PE.add("na");
                    }
                }
                break;
            case "PS":
                tester = (sheet.getRow(1).getCell(5).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(5).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(5).getNumericCellValue();
                }else{
                    do {
                        PS.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(5).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(5).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(5).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(5).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(5).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(5).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(5).getNumericCellValue();

                        PS.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(5).getNumericCellValue();
                    }else {
                        PS.add("na");
                    }
                }
                break;
            default:
                System.out.println("Please specify the Type of Return (PB, PE, PS)");
                break;
        }
    }

    void dateAdd (Date sampleDate){
        datum.add(sampleDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
    }

    void returnsOut(){
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");

        System.out.printf("%s%20s%14s%14s%n", "Date", "Mid-Return", "Bid-Return", "Ask-Return");

        for(int i=0;i<(datum.size()-1);i++){
            System.out.printf("%s%14s%14s%14s%n", datum.get(i+1).format(formatter), oneDay_midReturn.get(i), oneDay_bidReturn.get(i), oneDay_askReturn.get(i));
        }
    }

    void dataOut(){
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");

        System.out.printf("%s%20s%14s%n", "Date", "Market Cap", "Volume");
        for(int i=0;i<(datum.size()-1);i++){
            System.out.printf("%s%14s%14s%n",datum.get(i+1).format(formatter), mktCap.get(i), volume.get(i));
        }
    }
}
