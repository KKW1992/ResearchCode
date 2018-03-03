import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

public class CompanyData {
    //Object wide Variables for functions
    Sheet sheet = null;
    int lastRow = 0;
    String name;

    //Returns
    ArrayList<String> oneDay_midReturn = new ArrayList<>(lastRow);
    ArrayList<String> midPrice = new ArrayList<>(lastRow);
    ArrayList<String> oneDay_bidReturn = new ArrayList<>(lastRow);
    ArrayList<String> bidPrice = new ArrayList<>(lastRow);
    ArrayList<String> oneDay_askReturn = new ArrayList<>(lastRow);
    ArrayList<String> askPrice = new ArrayList<>(lastRow);

    //Lagged Returns
    private ArrayList<String> completeReturns = new ArrayList<>(lastRow);
    ArrayList<String> oneWeek_midRetrun = new ArrayList<>(lastRow);
    ArrayList<String> oneMonth_midReturn = new ArrayList<>(lastRow);
    ArrayList<String> threeMonth_midReturn = new ArrayList<>(lastRow);
    ArrayList<String> sixMonth_midReturn = new ArrayList<>(lastRow);

    //Market Data
    ArrayList<String> mktCap = new ArrayList<>(lastRow);
    ArrayList<String> volume = new ArrayList<>(lastRow);

    //Company specific Data
    ArrayList<String> PB = new ArrayList<>(lastRow);
    ArrayList<String> PE = new ArrayList<>(lastRow);
    ArrayList<String> PS = new ArrayList<>(lastRow);

    //Volatility
    ArrayList<String> oneWeek_Volatility = new ArrayList<>(lastRow);
    ArrayList<String> oneMonth_Volatility = new ArrayList<>(lastRow);
    ArrayList<String> threeMonth_Volatility = new ArrayList<>(lastRow);
    ArrayList<String> sixMonth_Volatility = new ArrayList<>(lastRow);

    ArrayList<LocalDate> datum = new ArrayList<>(lastRow);


    CompanyData(Sheet sheet){
        this.sheet = sheet;
        lastRow = sheet.getLastRowNum();
        name = sheet.getRow(0).getCell(0).getStringCellValue();
    }

    void addReturns(String returnType){
        double priceOne = 0;
        double priceTwo = 0;
        boolean tester;
        int x = 1;

        switch (returnType){
            case "Mid":
                tester = (sheet.getRow(1).getCell(1).getCellTypeEnum()== CellType.NUMERIC && sheet.getRow(1).getCell(1).getNumericCellValue()!=0);
                if(tester){
                    priceOne = sheet.getRow(1).getCell(1).getNumericCellValue();
                    oneDay_midReturn.add("na");
                }else{
                    do {
                        midPrice.add("na");
                        oneDay_midReturn.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(1).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(1).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(1).getNumericCellValue()==0)));
                    priceOne = sheet.getRow(x).getCell(1).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(1).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(1).getNumericCellValue() != 0);

                    if (tester) {
                        priceTwo = sheet.getRow(i).getCell(1).getNumericCellValue();

                        midReturnCalc(priceOne, priceTwo);
                        midPrice.add(Double.toString(priceOne));

                        priceOne = sheet.getRow(i).getCell(1).getNumericCellValue();
                    }else{
                        oneDay_midReturn.add("na");
                        midPrice.add("na");
                    }
                }
                midPrice.add(Double.toString(priceOne));
                midPrice.trimToSize();
                oneDay_midReturn.trimToSize();
                break;
            case "Bid":
                tester = (sheet.getRow(1).getCell(2).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(2).getNumericCellValue()!=0);
                if(tester){
                    priceOne = sheet.getRow(1).getCell(2).getNumericCellValue();
                    oneDay_bidReturn.add("na");
                }else{
                    do {
                        bidPrice.add("na");
                        oneDay_bidReturn.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(2).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(2).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(2).getNumericCellValue()==0)));
                    priceOne = sheet.getRow(x).getCell(2).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(2).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(2).getNumericCellValue() != 0);

                    if (tester) {
                        priceTwo = sheet.getRow(i).getCell(2).getNumericCellValue();

                        bidReturnCalc(priceOne, priceTwo);
                        bidPrice.add(Double.toString(priceOne));

                        priceOne = sheet.getRow(i).getCell(2).getNumericCellValue();
                    }else{
                        oneDay_bidReturn.add("na");
                        bidPrice.add("na");
                    }
                }
                bidPrice.add(Double.toString(priceOne));
                bidPrice.trimToSize();
                oneDay_bidReturn.trimToSize();
                break;
            case "Ask":
                tester = (sheet.getRow(1).getCell(3).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(3).getNumericCellValue()!=0);
                if(tester){
                    priceOne = sheet.getRow(1).getCell(3).getNumericCellValue();
                    oneDay_askReturn.add("na");
                }else{
                    do {
                        askPrice.add("na");
                        oneDay_askReturn.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(3).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(3).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(3).getNumericCellValue()==0)));
                    priceOne = sheet.getRow(x).getCell(3).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(3).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(3).getNumericCellValue() != 0);

                    if (tester) {
                        priceTwo = sheet.getRow(i).getCell(3).getNumericCellValue();

                        askReturnCalc(priceOne, priceTwo);
                        askPrice.add(Double.toString(priceOne));

                        priceOne = sheet.getRow(i).getCell(3).getNumericCellValue();
                    }else{
                        oneDay_askReturn.add("na");
                        askPrice.add("na");
                    }
                }
                askPrice.add(Double.toString(priceOne));
                askPrice.trimToSize();
                oneDay_askReturn.trimToSize();
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

    void addMarketData(String dataType){
        double valueOne = 0;
        double valueTwo = 0;
        boolean tester;
        int x = 1;

        switch (dataType){
            case "MktCap":
                tester = (sheet.getRow(1).getCell(4).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(4).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(4).getNumericCellValue();
                    mktCap.add("na");
                }else{
                    do {
                        mktCap.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(4).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(4).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(4).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(4).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(4).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(4).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(4).getNumericCellValue();

                        mktCap.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(4).getNumericCellValue();
                    }else{
                        mktCap.add("na");
                    }
                }
                mktCap.trimToSize();
                break;
            case "Volume":
                tester = (sheet.getRow(1).getCell(11).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(11).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(11).getNumericCellValue();
                    volume.add("na");
                }else{
                    do {
                        volume.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(11).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(11).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(11).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(11).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=2;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(11).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(11).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(11).getNumericCellValue();

                        volume.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(11).getNumericCellValue();
                    }else{
                        volume.add("na");
                    }
                }
                volume.trimToSize();
                break;
            default:
                System.out.println("Please specify the Type of Return (MktCap, Volume)");
                break;
        }
    }

    void addCompanyData(String dataType){
        double valueOne = 0;
        double valueTwo = 0;
        boolean tester;
        int x = 1;

        switch (dataType){
            case "PB":
                tester = (sheet.getRow(1).getCell(8).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(8).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(8).getNumericCellValue();
                    PB.add("na");
                }else{
                    do {
                        PB.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(8).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(8).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(8).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(8).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(8).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(8).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(8).getNumericCellValue();

                        PB.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(8).getNumericCellValue();
                    }else{
                        PB.add("na");
                    }
                }
                PB.trimToSize();
                break;
            case "PE":
                tester = (sheet.getRow(1).getCell(6).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(6).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(6).getNumericCellValue();
                    PE.add("na");
                }else{
                    do {
                        PE.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(6).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(6).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(6).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(6).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(6).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(6).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(6).getNumericCellValue();

                        PE.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(6).getNumericCellValue();
                    }else{
                        PE.add("na");
                    }
                }
                PE.trimToSize();
                break;
            case "PS":
                tester = (sheet.getRow(1).getCell(5).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(1).getCell(5).getNumericCellValue()!=0);
                if(tester){
                    valueOne = sheet.getRow(1).getCell(5).getNumericCellValue();
                    PS.add("na");
                }else{
                    do {
                        PS.add("na");
                        x++;
                    }while ((sheet.getRow(x).getCell(5).getCellTypeEnum()!=CellType.NUMERIC || (sheet.getRow(x).getCell(5).getCellTypeEnum()==CellType.NUMERIC && sheet.getRow(x).getCell(5).getNumericCellValue()==0)));
                    valueOne = sheet.getRow(x).getCell(5).getNumericCellValue();
                }

                //Loop through all rows of a sheet
                for(int i=x+1;i<lastRow+1;i++) {
                    tester = (sheet.getRow(i).getCell(5).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(5).getNumericCellValue() != 0);

                    if (tester) {
                        valueTwo = sheet.getRow(i).getCell(5).getNumericCellValue();

                        PS.add(Double.toString(Math.log(valueTwo/valueOne)));

                        valueOne = sheet.getRow(i).getCell(5).getNumericCellValue();
                    }else {
                        PS.add("na");
                    }
                }
                PS.trimToSize();
                break;
            default:
                System.out.println("Please specify the Type of Return (PB, PE, PS)");
                break;
        }
    }

    void dateAdd (Date sampleDate){
        datum.add(sampleDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
        datum.trimToSize();
    }

    private void completeReturn(){
        boolean tester;
        int x = 1;

        tester = (sheet.getRow(1).getCell(1).getCellTypeEnum()== CellType.NUMERIC && sheet.getRow(1).getCell(1).getNumericCellValue()!=0);
        if(tester){
            completeReturns.add(Double.toString(sheet.getRow(1).getCell(1).getNumericCellValue()));
            x++;
        }else {
            do {
                completeReturns.add("na");
                x++;
            }
            while ((sheet.getRow(x).getCell(1).getCellTypeEnum() != CellType.NUMERIC || (sheet.getRow(x).getCell(1).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(x).getCell(1).getNumericCellValue() == 0)));
        }

        //Loop through all rows of a sheet
        for(int i=x;i<lastRow+1;i++) {
            tester = (sheet.getRow(i).getCell(1).getCellTypeEnum() == CellType.NUMERIC && sheet.getRow(i).getCell(1).getNumericCellValue() != 0);

            if (tester) {
                completeReturns.add(Double.toString(sheet.getRow(i).getCell(1).getNumericCellValue()));
            } else {
                completeReturns.add(Double.toString(sheet.getRow(i - 1).getCell(1).getNumericCellValue()));
            }
        }
        completeReturns.trimToSize();
    }

    void addLaggedReturns( String returnType){
        boolean tester;
        int x = 0;

        this.completeReturn();

        switch (returnType){
            case "oneWeek":
                //find first row with numeric value
                tester = (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                if(tester){
                    do {
                        oneWeek_midRetrun.add("na");
                        x++;
                    }while (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                }

                for(int i = x;i<x+4;i++){
                    oneWeek_midRetrun.add("na");
                }

                //Loop through all rows of a sheet
                for(int i=x+4;i<completeReturns.size();i++) {
                    oneWeek_midRetrun.add(Double.toString(Math.log(new Double(completeReturns.get(i))/new Double(completeReturns.get(i-4)))));
                }
                oneWeek_midRetrun.trimToSize();
                break;
            case "oneMonth":
                //find first row with numeric value
                tester = (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                if(tester){
                    do {
                        oneMonth_midReturn.add("na");
                        x++;
                    }while (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                }

                for(int i = x;i<x+20;i++){
                    oneMonth_midReturn.add("na");
                }

                //Loop through all rows of a sheet
                for(int i=x+20;i<completeReturns.size();i++) {
                    oneMonth_midReturn.add(Double.toString(Math.log(new Double(completeReturns.get(i))/new Double(completeReturns.get(i-20)))));
                }
                oneMonth_midReturn.trimToSize();
                break;
            case "threeMonth":
                //find first row with numeric value
                tester = (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                if(tester){
                    do {
                        threeMonth_midReturn.add("na");
                        x++;
                    }while (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                }

                for(int i = x;i<x+62;i++){
                    threeMonth_midReturn.add("na");
                }

                //Loop through all rows of a sheet
                for(int i=x+62;i<completeReturns.size();i++) {
                    threeMonth_midReturn.add(Double.toString(Math.log(new Double(completeReturns.get(i))/new Double(completeReturns.get(i-62)))));
                }
                threeMonth_midReturn.trimToSize();
                break;
            case "sixMonth":
                //find first row with numeric value
                tester = (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                if(tester){
                    do {
                        sixMonth_midReturn.add("na");
                        x++;
                    }while (completeReturns.get(x).equals("na") || completeReturns.get(x).equals("0"));
                }

                for(int i = x;i<x+124;i++){
                    sixMonth_midReturn.add("na");
                }

                //Loop through all rows of a sheet
                for(int i=x+124;i<completeReturns.size();i++) {
                    sixMonth_midReturn.add(Double.toString(Math.log(new Double(completeReturns.get(i))/new Double(completeReturns.get(i-124)))));
                }
                sixMonth_midReturn.trimToSize();
                break;
            default:
                System.out.println("Please specify the Type of Return (oneWeek, oneMonth, threeMonth, sixMonth)");
                break;
        }
    }

    void addVolatility(String duration){
        this.completeReturn();

        switch (duration){
            case "oneWeek":
                //Loop through all rows of a sheet
                for(int i=0;i<4;i++){
                    oneWeek_Volatility.add("na");
                }

                for(int i=4;i<completeReturns.size();i++) {
                    oneWeek_Volatility.add(String.valueOf(Statistics.Variance(completeReturns.subList(i-4,i))));
                }
                oneWeek_Volatility.trimToSize();
                break;
            case "oneMonth":
                //Loop through all rows of a sheet
                for(int i=0;i<21;i++){
                    oneMonth_Volatility.add("na");
                }

                for(int i=21;i<completeReturns.size();i++) {
                    oneMonth_Volatility.add(String.valueOf(Statistics.Variance(completeReturns.subList(i-21,i))));
                }
                oneMonth_Volatility.trimToSize();
                break;
            case "threeMonth":
                //Loop through all rows of a sheet
                for(int i=0;i<63;i++){
                    threeMonth_Volatility.add("na");
                }

                for(int i=63;i<completeReturns.size();i++) {
                    threeMonth_Volatility.add(String.valueOf(Statistics.Variance(completeReturns.subList(i-63,i))));
                }
                threeMonth_Volatility.trimToSize();
                break;
            case "sixMonth":
                //Loop through all rows of a sheet
                for(int i=0;i<125;i++){
                    sixMonth_Volatility.add("na");
                }

                for(int i=125;i<completeReturns.size();i++) {
                    sixMonth_Volatility.add(String.valueOf(Statistics.Variance(completeReturns.subList(i-125,i))));
                }
                sixMonth_Volatility.trimToSize();
                break;
            default:
                System.out.println("Please specify the Type of Return (oneWeek, oneMonth, threeMonth, sixMonth)");
                break;
        }
    }

    void writeData() throws FileNotFoundException {
        StringBuffer sb = new StringBuffer(1000);

    PrintWriter pw = new PrintWriter(new File("C:\\Users\\kimwa\\OneDrive\\Documents\\Codes\\Java\\US-OutPut\\test.csv"));

        sb.append(name).append(";");
        sb.append("Mid_Price").append(";");
        sb.append("Bid_Price").append(";");
        sb.append("Ask_Price").append(";");

        sb.append("Mid_Return").append(";");
        sb.append("Bid_Return").append(";");
        sb.append("Ask_Return").append(";");

        sb.append("OneWeek_MidReturn").append(";");
        sb.append("OneMonth_MidReturn").append(";");
        sb.append("ThreeMonth_MidReturn").append(";");
        sb.append("SixMonth_MidReturn").append(";");

        sb.append("OneWeek_Volatility").append(";");
        sb.append("OneMonth_Volatility").append(";");
        sb.append("ThreeMonth_Volatility").append(";");
        sb.append("SixMonth_Volatility").append(";");

        sb.append("Mkt_Cap").append(";");
        sb.append("Volume").append(";");

        sb.append("PB").append(";");
        sb.append("PE").append(";");
        sb.append("PS");

        for(int i=0;i<datum.size();i++){
            sb.append("\n").append(datum.get(i)).append(";");
            sb.append(midPrice.get(i)).append(";");
            sb.append(bidPrice.get(i)).append(";");
            sb.append(askPrice.get(i)).append(";");

            sb.append(oneDay_midReturn.get(i)).append(";");
            sb.append(oneDay_bidReturn.get(i)).append(";");
            sb.append(oneDay_askReturn.get(i)).append(";");

            sb.append(oneWeek_midRetrun.get(i)).append(";");
            sb.append(oneMonth_Volatility.get(i)).append(";");
            sb.append(threeMonth_midReturn.get(i)).append(";");
            sb.append(sixMonth_midReturn.get(i)).append(";");

            sb.append(oneWeek_Volatility.get(i)).append(";");
            sb.append(oneMonth_Volatility.get(i)).append(";");
            sb.append(threeMonth_Volatility.get(i)).append(";");
            sb.append(sixMonth_Volatility.get(i)).append(";");

            sb.append(mktCap.get(i)).append(";");
            sb.append(volume.get(i)).append(";");

            sb.append(PB.get(i)).append(";");
            sb.append(PE.get(i)).append(";");
            sb.append(PS.get(i)).append(";");
        }

        pw.write(sb.toString());
        pw.close();
        System.out.println("Done Writing");
    }
}
