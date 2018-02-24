
public class Statistics {
    double Mean(double data[]){
        double sum = 0;
        for(double d:data){
            sum += d;
        }
        return sum/data.length;
    }

    double Variance(double data[]){
        double mean = this.Mean(data);
        double temp = 0;

        for(double d:data){
            temp+=(d-mean)*(d-mean);
        }
        return temp/data.length;
    }

    double StdDeviation(double Variance){
        return Math.sqrt(Variance);
    }
}
