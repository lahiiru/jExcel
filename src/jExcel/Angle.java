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
public class Angle {
    private String delemeter=".";
    protected int limit=360;
    private double level0=0.0;
    private int level1=0;
    private int level2=0;
    public Angle(int v2,int v1,double v0){
        this.level0=v0;
        this.level1=v1;
        this.level2=v2;        
    }
    public void setValue(int v2,int v1,double v0){
        this.level0=v0;
        this.level1=v1;
        this.level2=v2;
    }
    public void setSymbol(String symbol){
        this.delemeter=symbol;
    }

    public void echo(){
        System.out.println(level2+delemeter+level1+delemeter+level0);
    }
    public double[] getValue(){
        double[] value={this.level0,this.level1,this.level2};
        return value;
    }
    public void add(Angle angle1,Angle angle2){
        double[] value1=angle1.getValue();
        double[] value2=angle2.getValue();
        double temp1=0.0;
        double temp2=0.0;
        for(int i=0;i<3;i++){
            temp1+=value1[i]*Math.pow(limit, i);
        }
        for(int i=0;i<3;i++){
            temp2+=value2[i]*Math.pow(limit, i);
        }
        temp1+=temp2;
        
        this.level0=temp1%limit;temp1=temp1/limit;
        this.level1=(int)(temp1%limit);temp1=temp1/limit;
        this.level2=(int)(temp1);
    }
    public void substract(Angle angle1,Angle angle2){
        double[] value1=angle1.getValue();
        double[] value2=angle2.getValue();
        this.level0=value1[0]+value1[0];
        this.level1=(int)(value1[0]+value1[0]);
        this.level2=(int)(value1[0]+value1[0]);
    }
}
