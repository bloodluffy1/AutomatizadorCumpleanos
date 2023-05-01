package org.example;
import java.util.Date;
public class Empleado {
    private String nombre;
    private Date  Cumpleanos;
    private String Area;
    private  String Correo;
    public Empleado(String nombre, Date Cumpleanos, String Area, String Correo) {
        this.nombre = nombre;
        this.Cumpleanos = Cumpleanos;
        this.Area = Area;
        this.Correo = Correo;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public Date getCumpleanos() {
        return Cumpleanos;
    }

    public void setCumpleanos(Date Cumpleanos) {
        this.Cumpleanos = Cumpleanos;
    }

    public String getArea() {
        return Area;
    }

    public void setArea(String Area) {
        this.Area = Area;
    }

    public String getCorreo() {
        return Correo;
    }

    public void setCorreo(String Correo) {
        this.Correo = Correo;
    }


}
