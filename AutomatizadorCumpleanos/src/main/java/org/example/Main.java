package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
public class Main {
    //Proceso INICIAL
    public static void main(String[] args) {
        Timer timer = new Timer();
        timer.schedule(new TareaProgramada(), calculaDelayInicial(), 24 * 60 * 60 * 1000);
    }
    // 1. Metodo Para obtener la hora configurada
    /*private static long calculaDelayInicial() {
        Calendar calendario = Calendar.getInstance();
        int horaActual = calendario.get(Calendar.HOUR_OF_DAY);
        int minutosActuales = calendario.get(Calendar.MINUTE);
        int segundosActuales = calendario.get(Calendar.SECOND);
        long tiempoHastaLas8AM = (8 - horaActual) * 60 * 60 * 1000 - minutosActuales * 60 * 1000 - segundosActuales * 1000;
        if (tiempoHastaLas8AM < 0) {
            tiempoHastaLas8AM += 24 * 60 * 60 * 1000;
        }
        System.out.println(tiempoHastaLas8AM);
        return tiempoHastaLas8AM;
    }*/
    // 1. Metodo Para obtener la hora configurada
    private static long calculaDelayInicial() {
        Calendar calendario = Calendar.getInstance();
        int horaActual = calendario.get(Calendar.HOUR_OF_DAY);
        int minutosActuales = calendario.get(Calendar.MINUTE);
        int segundosActuales = calendario.get(Calendar.SECOND);
        long tiempoHastaLas930PM = (21 - horaActual) * 60 * 60 * 1000 + (30 - minutosActuales) * 60 * 1000 - segundosActuales * 1000;
        if (tiempoHastaLas930PM < 0) {
            tiempoHastaLas930PM += 24 * 60 * 60 * 1000;
        }
        System.out.println(tiempoHastaLas930PM);
        return tiempoHastaLas930PM;
    }

    //2. Si cumple con la ejecucion hora programa se ejecuta el proceso
    static class TareaProgramada extends TimerTask {
        public void run() {
            List<Empleado> listaEmpleados = leerArchivoExcel("C:/Clientes/Enero.xlsx");
            ValidaProcesoEnvioSaludo(listaEmpleados);

        }
    }

    //3. Lee el archivo excel y devuelve una lista con la informacion de los empleados
    public static List<Empleado> leerArchivoExcel(String archivo) {
        // Declara el objeto en una lista
        List<Empleado> listaEmpleados = new ArrayList<Empleado>();
        //Obtiene el archivo excel.
        try (FileInputStream archivoEntrada = new FileInputStream(archivo);
             Workbook libroExcel = WorkbookFactory.create(archivoEntrada)) {
            Sheet hoja = libroExcel.getSheetAt(0);
            int filaInicial = 2; // la fila B3 es la fila 2
            int filaFinal = hoja.getLastRowNum(); // última fila de la hoja
            int columnaInicial = 1; // columna B
            int columnaFinal = 4; // columna E

            boolean salida = false;
            //Obtiene la informacion desde un rango de celdas configurado.
            CellRangeAddress rangoCeldas = new CellRangeAddress(filaInicial, filaFinal, columnaInicial, columnaFinal);
            // Iteración de las filas.
            for (int i = rangoCeldas.getFirstRow(); i <= rangoCeldas.getLastRow(); i++) {
                Row fila = hoja.getRow(i);
                if (fila != null) {
                    String nombre = "";
                    Date fecha = null;
                    String area = "";
                    String correo = "";
                    // Iteración de las columnas de cada fila : Para esta caso son 4 iteraciones por fila.
                    for (int j = rangoCeldas.getFirstColumn(); j <= rangoCeldas.getLastColumn(); j++) {
                        Cell celda = fila.getCell(j);
                        if (celda != null) {
                            switch (j) {
                                case 1:
                                    nombre = celda.getStringCellValue();
                                    break;
                                case 2:
                                    fecha = celda.getDateCellValue();
                                    break;
                                case 3:
                                    area = celda.getStringCellValue();
                                case 4:
                                    correo = celda.getStringCellValue();
                                    break;
                            }
                        }else{
                            salida = true;
                            break;
                        }
                    }

                    if(!salida)
                    {
                        //Alimenta el objeto a las lista con la informacion de los clientes
                        listaEmpleados.add(new Empleado(nombre, fecha, area, correo));
                    }else{
                        break;
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return listaEmpleados;
    }

    //4. Valida la informacion de los empleados su fecha de cumpleaños
    /*
       Criterios son los siguiente:
       Fecha de hoy es igual a la fecha de cumpleaños
       Llama al proceso de envio de correo saludo (enviarCorreoDeFelicitacion)
    * */
    private static void ValidaProcesoEnvioSaludo(List<Empleado> listaEmpleados)
    {
        for (Empleado emp : listaEmpleados) {
            Date fechaActual = new Date();
            Calendar FechaActual = Calendar.getInstance();
            FechaActual.setTime(fechaActual);
            Calendar cumpleanos = Calendar.getInstance();
            cumpleanos.setTime(emp.getCumpleanos());
            int mesCumpleanos = cumpleanos.get(Calendar.MONTH) + 1; // sumamos 1 porque los meses en Calendar comienzan en 0
            int diaCumpleanos = cumpleanos.get(Calendar.DAY_OF_MONTH);

            int mesActual = FechaActual.get(Calendar.MONTH) + 1; // sumamos 1 porque los meses en Calendar comienzan en 0
            int diaActual = FechaActual.get(Calendar.DAY_OF_MONTH);

            if (mesCumpleanos == mesActual && diaCumpleanos == diaActual)
            {
                enviarCorreoDeFelicitacion(emp);
            }

        }

    }

    // 5. Envia el correo a los empleados que cumplen años en la fecha actual
    private static void enviarCorreoDeFelicitacion(Empleado empleado) {
        // Configura las propiedades de la sesión de correo.
        Properties propiedades = new Properties();
        propiedades.put("mail.smtp.auth", "true");
        propiedades.put("mail.smtp.starttls.enable", "true");
        propiedades.put("mail.smtp.host", "smtp.gmail.com");
        propiedades.put("mail.smtp.port", "587");
        // Ingresa tus credenciales de correo electrónico de Gmail aquí.
        final String correoUsuario = "bloodluffy@gmail.com";
        final String contrasenia = "rlcrwdkxvjdevnfo";

        // Inicia sesión en el servidor de correo.
        Session sesion = Session.getInstance(propiedades, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(correoUsuario, contrasenia);
            }
        });
        try {
            // Construye el mensaje de correo electrónico.
            Message mensaje = new MimeMessage(sesion);
            mensaje.setFrom(new InternetAddress(correoUsuario));
            mensaje.setRecipients(Message.RecipientType.TO, InternetAddress.parse(empleado.getCorreo()));
            mensaje.setSubject("¡Feliz cumpleaños, " + empleado.getNombre() + "!");
            // Adjunta una imagen de felicitación al mensaje de correo electrónico.
            BodyPart mensajeParte0 = new MimeBodyPart();
            mensajeParte0.setText("¡Feliz cumpleaños!");
            BodyPart mensajeParte1 = new MimeBodyPart();
            mensajeParte1.setText("\n Nuestros mejores anhelos de salud y éxitos para " + empleado.getNombre());
            BodyPart mensajeParte2 = new MimeBodyPart();
            mensajeParte2.setText("\n del área " + empleado.getArea() + " en este día tan especial. ");
            BodyPart mensajeParte3 = new MimeBodyPart();
            mensajeParte3.setText("\n con aprecio y admiración,");
            BodyPart mensajeParte4 = new MimeBodyPart();
            mensajeParte4.setText("\n le desea,");
            BodyPart mensajeParte5 = new MimeBodyPart();
            mensajeParte5.setText("\n El equipo de Grupo Imaco.");
            BodyPart mensajeParte6 = new MimeBodyPart();
            DataSource fuente = new FileDataSource("C:/Clientes/Saludo/Saludo.jpg");
            mensajeParte6.setDataHandler(new DataHandler(fuente));
            mensajeParte6.setFileName("Saludo.jpg");
            // Crea un mensaje compuesto que incluye el texto y la imagen de felicitación.
            Multipart mensajeCompuesto = new MimeMultipart();
            mensajeCompuesto.addBodyPart(mensajeParte0);
            mensajeCompuesto.addBodyPart(mensajeParte1);
            mensajeCompuesto.addBodyPart(mensajeParte2);
            mensajeCompuesto.addBodyPart(mensajeParte3);
            mensajeCompuesto.addBodyPart(mensajeParte4);
            mensajeCompuesto.addBodyPart(mensajeParte5);
            mensajeCompuesto.addBodyPart(mensajeParte6);
            mensaje.setContent(mensajeCompuesto);
            // Envía el mensaje de correo electrónico.
            Transport.send(mensaje);
            System.out.println("Mensaje enviado a " + empleado.getCorreo());
        } catch (MessagingException e) {
            e.printStackTrace();
        }
    }
}