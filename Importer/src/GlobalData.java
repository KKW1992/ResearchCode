import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;

public class GlobalData {
    ArrayList<String> russelIndex = new ArrayList<>();
    ArrayList<String> russelAbs = new ArrayList<>();
    ArrayList<String> spIndex = new ArrayList<>();
    ArrayList<String> spAbs = new ArrayList<>();
    ArrayList<String> riskFree = new ArrayList<>();
    ArrayList<String> riskFreeAbs = new ArrayList<>();

    ArrayList<LocalDate> datum = new ArrayList<>();

    void addMarketData(Sheet sheet, String dataType){
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
                        valueOne = sheet.getRow(x).getCell(15).getNumericCellValue(); //find first numeric value
                        x++;
                    }while ((sheet.getRow(x).getCell(15).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(15).getNumericCellValue()!=0));
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
                break;
            case "Russel":
                tester = (sheet.getRow(1).getCell(14).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(14).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(14).getNumericCellValue();
                }else{
                    System.out.println("Here comes a routine if the first value is not numeric");
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
                break;
            case "RiskFree":
                tester = (sheet.getRow(1).getCell(16).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(16).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(16).getNumericCellValue();
                }else{
                    do {
                        valueOne = sheet.getRow(x).getCell(16).getNumericCellValue(); //find first numeric value
                        x++;
                    }while ((sheet.getRow(x).getCell(16).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(16).getNumericCellValue()!=0));
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
