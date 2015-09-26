package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
public class Time extends Angle{
    public Time(int v2,int v1,double v0){
        super(v2,v1,v0);
        super.setSymbol(":");
        super.limit=60;
    }
}
