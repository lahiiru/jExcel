package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
public class Currency extends Number {
    private String currencySymbol="$";
    
    public Currency(double value){
        super(value);
    }
    public void echo(){
        System.out.println(String.format( currencySymbol+" %."+String.valueOf(decimalPlaces)+"f",this.getValue() ));
    }
    public void setSymbol(String symbol){
        this.currencySymbol=symbol;
    }
    public String getSymbol(){
        return this.currencySymbol;
    }
}
