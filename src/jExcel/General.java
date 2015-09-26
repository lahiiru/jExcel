/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package jExcel;

/**
 *
 * @author adm.trine@gmail.com
 */
abstract public class General{
    private double value=0.0;
    public General(double value){
        this.value=value;
    }
    abstract public void echo();

    public void setValue(double value){
        this.value=value;
    }
    public double getValue(){
        return this.value;
    }
        public void add(double... values){
         double temp=0.0;
         for (Double arg : values) {temp+=arg;}
         setValue(temp);
    }
    public void substract(double value1,double value2){
         setValue(value1-value2);
    }
    public void divide(double value1,double value2){
         setValue(value1/value2);
    }
    public void multiply(double value1,double value2){
         setValue(value1*value2);
    }
    public void max(double value1,double value2){
        if(value1>=value2){
            setValue(value1);
        }
        else{
            setValue(value2);;
        }
    }
    public void min(double value1,double value2){
        if(value1<value2){
            setValue(value1);
        }
        else{
            setValue(value2);;
        }
    }
}
