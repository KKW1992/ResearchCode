import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        PeriodData firstPeriod = new PeriodData(1,3000);
        System.out.println("Max-Row: " + firstPeriod.maxRow);
        System.out.println("First Date: " + firstPeriod.firstDate);
    }
}

