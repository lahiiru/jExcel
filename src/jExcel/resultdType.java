package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
public class resultdType {
    public static Object resultdType(String input){
        String firstElem="";
        int k=0;
        while(k<input.length() && firstElem.length()<1){
            firstElem=input.split("[\\+\\*\\-\\)\\(/]")[k];
            k++;
        }
        
        Object lastDataTypeObj=dType.getType(firstElem);
        String lastDataType=objToStr(lastDataTypeObj);
        for(String x:input.split("[\\+\\*\\-\\)\\(/]")){
           if(x.length()==0)continue;
            Object DataTypeObj=dType.getType(x);
            String DataType=objToStr(DataTypeObj); 
           if(!lastDataType.equals(DataType))
            {
                
                return "#ERROR";
            }
            lastDataType=DataType;

        }
        return lastDataTypeObj;
    }
    public static String resultdTypeStr(String input){
        return objToStr(resultdType(input));
    }
    public static String objToStr(Object o){
        String r="Text";
        if(o.getClass().equals(Number.class))r="Number";
        if(o.getClass().equals(Currency.class))r="Currency";
        if(o.getClass().equals(Date.class))r="Date";
        if(o.getClass().equals(Time.class))r="Time";
        if(o.getClass().equals(Angle.class))r="Angle";
        return r;
    }
}
