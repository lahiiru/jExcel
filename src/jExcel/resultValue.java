package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
public class resultValue {

    private String val = "";
    private Object valObj;
    private String dType = "Text";

    public resultValue(String input) {
        Object dataTypeObj = resultdType.resultdType(input);
        dType = resultdType.objToStr(dataTypeObj);
        if (dType.equals("Number")) {
            InfixExpression myExpression = new InfixExpression(input);
            val = myExpression.getValue();
            Number valObj = new Number(Double.valueOf(val));
            this.valObj = valObj;
        } else if (dType.equals("Currency")) {
            InfixExpression myExpression = new InfixExpression(input);
            val = myExpression.getValue();
            Currency valObj = new Currency(Double.valueOf(val));
            valObj.setSymbol(val);
            this.valObj = valObj;
        } else if (dType.equals("Text")) {
            if (input.contains("+") | input.contains("*") | input.contains("-") | input.contains("/")) {
                val = "#ERROR";
                this.valObj = "#ERROR";
            }

        }

    }

    public Object getResultObj() {
        return valObj;
    }

    public Object getResultStr() {
        return val;
    }

    // simplify functions
    public static String getRawExpression(String function) {
        String f = function.substring(0, 3);
        Integer c = function.split(",").length;
        if (f.toUpperCase().equals("SUM")) {
            return function.replaceFirst("[A-Z]+|[a-z]+", "").replaceAll(",", "+");
        } else if (f.toUpperCase().equals("COU")) {

            return c.toString();
        } else if (f.toUpperCase().equals("AVE")) {
            return function.replaceFirst("[A-Z]+|[a-z]+", "").replaceAll(",", "+") + "/" + c.toString();
        } else if (f.toUpperCase().equals("MAX")) {
            String[] tokens = function.split("[\\(\\),]");
            Double max = 0.0;
            for (String x : tokens) {
                if (!resultdType.resultdTypeStr(x).equals("Number")) {
                    continue;
                }
                Double i = Double.parseDouble(x);
                if (max < i) {
                    max = i;
                }
            }
            return max.toString();
        } else if (f.toUpperCase().equals("MIN")) {
            String[] tokens = function.split("[\\(\\),]");
            Double min = Double.MAX_VALUE;
            for (String x : tokens) {
                if (!x.matches("[0-9\\.]")) {
                    continue;
                }
                Double i = Double.parseDouble(x);
                if (min > i) {
                    min = i;
                }
            }
            return min.toString();
        }
        System.out.println("Unkown function!");
        return "";
    }
}
