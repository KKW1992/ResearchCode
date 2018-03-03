import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;

public class GlobalData {
    Sheet sheet = null;
    int lastRow = 0;

    ArrayList<String> russelIndex = new ArrayList<>(lastRow);
    ArrayList<String> russelAbs = new ArrayList<>(lastRow);
    ArrayList<String> spIndex = new ArrayList<>(lastRow);
    ArrayList<String> spAbs = new ArrayList<>(lastRow);
    ArrayList<String> riskFree = new ArrayList<>(lastRow);
    ArrayList<String> riskFreeAbs = new ArrayList<>(lastRow);

    ArrayList<LocalDate> datum = new ArrayList<>();

    GlobalData(Sheet sheet){
        this.sheet = sheet;
        lastRow = sheet.getLastRowNum();
    }

    void addMarketData(String dataType){
        double valueOne = 0;
        double valueTwo = 0;
        boolean tester;
        int x = 1;

        switch (dataType){
            case "S&P":
                tester = (sheet.getRow(1).getCell(15).getCellTypeEnum()== CellType.NUMERIC && sheet.getRow(1).getCell(15).getNumericCellValue()!=0); //Test if first value is numeric
                if(tester){
                    valueOne = sheet.getRow(1).getCell(15).getNumericCellValue(); //If true save value
                }else{
                    do {
                        spIndex.add("na");
                        spAbs.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(15).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(15).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(15).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(15).getNumericCellValue(); //find first numeric value
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(15).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(15).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(15).getNumericCellValue(); //find second numeric value

                        spIndex.add(Double.toString(Math.log(valueTwo/valueOne))); //calculate difference
                        spAbs.add(Double.toString(valueOne)); //set absolute Value

                        valueOne = sheet.getRow(i).getCell(15).getNumericCellValue(); //use second value as first value
                    }else{
                        spIndex.add("na");
                        spAbs.add("na");
                    }
                }
                spAbs.add(Double.toString(valueOne));
                spAbs.trimToSize();
                spIndex.trimToSize();
                break;
            case "Russel":
                tester = (sheet.getRow(1).getCell(14).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(14).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(14).getNumericCellValue();
                }else{
                    do {
                        russelAbs.add("na");
                        russelIndex.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(14).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(14).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(14).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(14).getNumericCellValue(); //find first numeric value
                }

                //Loop through all rows of a sheet
                for(int i=2;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(14).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(14).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(14).getNumericCellValue();

                        russelIndex.add(Double.toString(Math.log(valueTwo/valueOne)));
                        russelAbs.add(Double.toString(valueOne)); //set absolute Value

                        valueOne = sheet.getRow(i).getCell(14).getNumericCellValue();
                    }else{
                        russelIndex.add("na");
                        russelAbs.add("na");
                    }
                }
                russelAbs.add(Double.toString(valueOne));
                russelAbs.trimToSize();
                russelIndex.trimToSize();
                break;
            case "RiskFree":
                tester = (sheet.getRow(1).getCell(16).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(16).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(16).getNumericCellValue();
                }else{
                    do {
                        riskFreeAbs.add("na");
                        riskFree.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(16).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(16).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(16).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(16).getNumericCellValue(); //find first numeric value
                }

                //Loop through all rows of a sheet
                for(int i=2;i<sheet.getLastRowNum()+1;i++) {
                    tester = (sheet.getRow(i).getCell(16).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(16).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(16).getNumericCellValue();

                        riskFree.add(Double.toString(Math.log(valueTwo/valueOne)));
                        riskFreeAbs.add(Double.toString(valueOne)); //set absolute Value

                        valueOne = sheet.getRow(i).getCell(16).getNumericCellValue();
                    }else{
                        riskFree.add("na");
                        riskFreeAbs.add("na");
                    }
                }
                riskFreeAbs.add(Double.toString(valueOne));
                riskFreeAbs.trimToSize();
                riskFree.trimToSize();
                break;
            default:
                System.out.println("Please specify the Type of Return (S&P, Russel, RiskFree)");
                break;
        }
    }

    void dateAdd (Date sampleDate){
        datum.add(sampleDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
    }

}
