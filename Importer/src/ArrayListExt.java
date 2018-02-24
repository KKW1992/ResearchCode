import java.util.ArrayList;

public class ArrayListExt{
    public ArrayList replace(ArrayList<String> List, String old, String change){
        int x =0;
        ArrayList<String> replacedList = List;

        for(String t:replacedList){
            if (t.equals(old)) {
                replacedList.set(x,change);
            }
        }
        return replacedList;
    }
}
