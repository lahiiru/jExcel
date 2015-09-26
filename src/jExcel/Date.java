package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
public class Date{
    int year=1000;
    int month=1;
    int day=1;
    public void echo(){
        System.out.println(String.format(String.valueOf(year)+"/"+String.valueOf(month)+"/"+ String.valueOf(day)));
    }
}
