package jExcel;

import javax.swing.*;
import java.io.FileNotFoundException;
import javax.xml.stream.XMLStreamException;
/**
 */
public class Spreadsheet{
    public static void main(String[] args) throws XMLStreamException, FileNotFoundException {
        // Testing classes
        // Spreadsheet is as follows
        /**
        *   Name        EPF No      Work duration       Salary
        *   Amal        1250        1:30:00             Rs.200.00
        *   Kamal       1251        3:45:00             Rs.550.00
        *   Lahiru      1252        0:30:00             Rs.55.50
        *   Chamara     1253        2:30:00             Rs.400.00
        *   Total       -           Total time          Total cost
        */  
        
        //formatting cells and datatypes  *******************************
        Text[] Names= new Text[5];          //Names column
        Number[] EPF= new Number[5];        //EPF No column
        Time[] WorkDuration= new Time[5];   //Work duration column
        Currency[] Salary= new Currency[5]; //Salary column
        
        for(int n = 0; n < 5; n ++){
            Salary[n] = new Currency(0.0);
            Salary[n].setSymbol("Rs.");
            Salary[n].setDecimalPlaces(2);
        }
        
        //storing data *************************************************
        Names[0]=new Text("Amal");
        Names[1]=new Text("Kamal");
        Names[2]=new Text("Lahiru");
        Names[3]=new Text("Chamara");
        
        for(int n = 0; n < 4; n ++){
            EPF[n] = new Number(1250+n);
        }
        
        WorkDuration[0]=new Time(1,30,0.0);
        WorkDuration[1]=new Time(3,45,0.0);
        WorkDuration[2]=new Time(0,30,0.0);
        WorkDuration[3]=new Time(2,30,0.0);

        Salary[0].setValue(200.00);
        Salary[1].setValue(550.00);
        Salary[2].setValue(55.50);
        Salary[3].setValue(400.00);
        
        //generating total row ***************************************
        Names[4]=new Text("Total");
        EPF[4] = new Number(0);

        WorkDuration[4]=new Time(0,0,0.0);
        WorkDuration[4].add(WorkDuration[0], WorkDuration[1]);
        WorkDuration[4].add(WorkDuration[4], WorkDuration[2]);
        WorkDuration[4].add(WorkDuration[4], WorkDuration[3]);
        
        Salary[4].add(Salary[0].getValue(), Salary[1].getValue(), Salary[2].getValue(), Salary[3].getValue());

        System.out.print("Total work duration H:M:S\t:");WorkDuration[4].echo();
        System.out.print("Total cost\t\t\t:");Salary[4].echo();

        System.out.println("data type of the result of 10+15.0 is a "+resultdType.resultdTypeStr("10+15.1"));
        System.out.println("data type of the result of lahiru+2014/99/99 is a "+resultdType.resultdTypeStr("lahiru+2014/99/99"));
        System.out.println("data type of the result of LKR10.00-$1 is a "+resultdType.resultdTypeStr("LKR10.00+$1"));
        System.out.println("data type of the result of 2014/01/01+2014/12/01 is a "+resultdType.resultdTypeStr("2014/01/01+2014/12/01"));

        String expression1="(5+2*3)";
        String expression2="Lahiru+Jayakody";
        InfixExpression a=new InfixExpression(expression1);
        System.out.println(""+a.convertToPostfix());
        resultValue answer1 = new resultValue(expression1);
        System.out.println("Answer for "+expression1+" is "+answer1.getResultStr());
        
        resultValue answer2 = new resultValue(expression2);
        System.out.println("Answer for "+expression2+" is "+answer2.getResultStr()); 
        
        System.out.println(Cell.colonRemove("=sum(F0:E0)*sum(A0:E0)"));
       
       //Start jExcel GUI

       Window instance=new Window(null);
       instance.getBook().addNewSheet();            //adds a sheet as default
       SwingUtilities.invokeLater(instance);
    }

}
