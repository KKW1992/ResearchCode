import java.util.List;

public class Statistics {
    static double Mean(List<String> data){
        double sum = 0;
        double len = 0;

        for(String d:data){
            if(!d.equals("na")){
                sum += Double.parseDouble(d);
                len++;
            }
        }
        return sum/len;
    }

    static double Mean(Double[] data){
        double sum = 0;
        for(double d:data){
            sum += d;
        }
        return sum/data.length;
    }

    static double Variance(Double[] data){
        double mean = Mean(data);
        double temp = 0;

        for(double d:data){
            temp+=(d-mean)*(d-mean);
        }
        return temp/data.length;
    }

    static double Variance(List<String> data){
        double mean = Mean(data);
        double temp = 0;
        double len = 0;

        for(String d:data){
            if(!d.equals("na")){
                temp+=(Double.parseDouble(d)-mean)*(Double.parseDouble(d)-mean);
                len++;
            }
        }
        return temp/len;
    }

    double StdDeviation(double Variance){
        return Math.sqrt(Variance);
    }
}
