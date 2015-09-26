package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
public class dType {
    public static Object getType(String input){
        
        if(input.matches("[-+]?([0-9]*\\.[0-9]+|[0-9]+)")){
            Number n=new Number(0.0);
            if(input.contains(".")){
                n.setDecimalPlaces(input.length()-1-input.indexOf("."));
            }
            else{
                n.setDecimalPlaces(0);
            }
            return n;
        }
        if(input.matches("([1-2][0-9][0-9]{2})/(([0][0-9])|([1][0-2]))/(([0-2][0-9])|([3][0-1]))")){
            Date n=new Date();
            return n;
        }
        if(input.matches("([^0-9+-]{1,3})([0-9]*\\.[0-9]+|[0-9]+)")){
            Currency n=new Currency(0.0);
            if(input.matches("([^0-9+-]{1,3})([0-9]*\\.[0-9]+)")){
                n.setDecimalPlaces(input.length()-1-input.indexOf(".", input.indexOf("([^0-9+-]{1,3})")));
            }
            else{
                n.setDecimalPlaces(0);
            }
            n.setSymbol(input.split("([0-9]*\\.[0-9]+|[0-9]+)")[0]);
            return n;
        }
        if(input.matches("[0-9]+:[0-5][0-9]:[0-5][0-9]")){
            Time n=new Time(0,0,0.0);
            return n;        
        }
        if(input.matches("[0-9]+.[0-5][0-9].[0-5][0-9]")){
            Angle n=new Angle(0,0,0.0);
            return n;          
        }
        Text n=new Text ("");
        return n;  
    }
}
