import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;

public class TraceToExcel {
    public static void main(String[] args){
        if (args.length == 0){
            System.out.println("Утилита для конвертации файла трассировки из AFDPRO в Excel");
            System.out.println("Автор: Александр Махоткин");
            System.out.println();
            System.out.println("Arguments:");
            System.out.println("\t\t[FILE]");
            return;
        }

        File inputFile = new File(args[0]);
        String fileDir = inputFile.getPath().substring(0, inputFile.getPath().indexOf(inputFile.getName()));

        int dotPos = inputFile.getName().lastIndexOf('.');
        String nameWithoutExt = inputFile.getName();
        if (dotPos != -1)
            nameWithoutExt = inputFile.getName().substring(0, dotPos);

        try {
            ArrayList<CommandState> array = parseTrace(inputFile);
            generateExcel(fileDir + nameWithoutExt + ".xlsx", array);
        }catch(Exception ex){
            System.out.println(ex);
        }
    }

    public static ArrayList<CommandState> parseTrace(File inputFile) throws IOException{
        BufferedReader br = new BufferedReader(new FileReader(inputFile));
        br.readLine();
        br.readLine();
        br.readLine();

        String line = br.readLine();

        ArrayList<CommandState> states = new ArrayList<CommandState>();
        CommandState comState = null;

        while(!line.equals("*** End of TRACE buffer ***")){

            comState = new CommandState();
            comState.commandAddr = line.substring(0, 4);
            comState.ip = comState.commandAddr;

            comState.command = line.substring(5, 5 + 7).trim().toLowerCase();

            String arguments = line.substring(12, 12 + 28).trim();
            comState.arguments = arguments;
            int memAddrStart = arguments.indexOf('[');
            int memAddrEnd = arguments.indexOf(']');
            if (memAddrStart != -1){
                comState.memAddr = arguments.substring(memAddrStart + 1, memAddrEnd);
            }

            comState.ax = line.substring(44, 44 + 4);

            comState.stackState = line.substring(76, 76 + 4);


            line = br.readLine();

            comState.bx = line.substring(44, 44 + 4);

            line = br.readLine();

            comState.cx = line.substring(44, 44 + 4);

            line = br.readLine();

            String flagsLine = line.substring(17, 17 + 23).trim();
            String[] flags = flagsLine.split(" {2}");
            comState.Fof = Integer.parseInt(flags[0]);
            comState.Fdf = Integer.parseInt(flags[1]);
            comState.Fif = Integer.parseInt(flags[2]);
            comState.Fsf = Integer.parseInt(flags[3]);
            comState.Fzf = Integer.parseInt(flags[4]);
            comState.Faf = Integer.parseInt(flags[5]);
            comState.Fpf = Integer.parseInt(flags[6]);
            comState.Fcf = Integer.parseInt(flags[7]);

            comState.dx = line.substring(44, 44 + 4);

            states.add(comState);
            line = br.readLine();
        }

        if (comState != null) {
            comState.ip = "0000";
        }

        //Сдвиг

        CommandState lastState = new CommandState();
        states.add(lastState);
        for (int i = states.size() - 2; i >= 0; i--){
            CommandState st = states.get(i);
            CommandState stNext = states.get(i+1);
            stNext.commandAddr = st.commandAddr;
            stNext.command = st.command;
            stNext.memAddr = st.memAddr;
            stNext.arguments = st.arguments;
        }

        states.remove(0);

        br.close();
        return states;
    }

    public static void generateExcel(String file, ArrayList<CommandState> commandStates) throws IOException {
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Trace");

        Row row = sheet.createRow(0);
        row.setHeight((short) 600);

        CellStyle headStyle = book.createCellStyle();
        Font font = book.createFont();

        headStyle.setAlignment(HorizontalAlignment.CENTER);
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setBorderBottom(BorderStyle.THIN);

        font.setFontName("Times New Roman");
        font.setFontHeightInPoints((short) 10);

        headStyle.setFont(font);

        String[] headTitles = {"Адрес", "Команда", "AX", "BX", "CX", "DX", "IP", "OF", "DF", "IF", "SF", "ZF", "AF", "PF", "CF", "Память", "Стек", "Аргументы команды"};

        for (int i = 0; i < headTitles.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headTitles[i]);
            cell.setCellStyle(headStyle);
        }

        sheet.setColumnWidth(2, 1700);
        sheet.setColumnWidth(3, 1700);
        sheet.setColumnWidth(4, 1700);
        sheet.setColumnWidth(5, 1700);
        sheet.setColumnWidth(6, 1700);

        sheet.setColumnWidth(7, 800);
        sheet.setColumnWidth(8, 800);
        sheet.setColumnWidth(9, 800);
        sheet.setColumnWidth(10, 800);
        sheet.setColumnWidth(11, 800);
        sheet.setColumnWidth(12, 800);
        sheet.setColumnWidth(13, 800);
        sheet.setColumnWidth(14, 800);

        sheet.setColumnWidth(17, 5000);

        CellStyle dataStyle = book.createCellStyle();

        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);

        dataStyle.setFont(font);

        for (int i = 0; i < commandStates.size(); i++) {
            row = sheet.createRow(1 + i);
            CommandState comState = commandStates.get(i);

            Cell dataCell = row.createCell(0);
            dataCell.setCellValue(comState.commandAddr);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(1);
            dataCell.setCellValue(comState.command);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(2);
            dataCell.setCellValue(comState.ax);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(3);
            dataCell.setCellValue(comState.bx);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(4);
            dataCell.setCellValue(comState.cx);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(5);
            dataCell.setCellValue(comState.dx);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(6);
            dataCell.setCellValue(comState.ip);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(7);
            dataCell.setCellValue(comState.Fof);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(8);
            dataCell.setCellValue(comState.Fdf);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(9);
            dataCell.setCellValue(comState.Fif);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(10);
            dataCell.setCellValue(comState.Fsf);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(11);
            dataCell.setCellValue(comState.Fzf);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(12);
            dataCell.setCellValue(comState.Faf);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(13);
            dataCell.setCellValue(comState.Fpf);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(14);
            dataCell.setCellValue(comState.Fcf);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(15);
            dataCell.setCellValue(comState.memAddr);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(16);
            dataCell.setCellValue(comState.stackState);
            dataCell.setCellStyle(dataStyle);

            dataCell = row.createCell(17);
            dataCell.setCellValue(comState.arguments);
            dataCell.setCellStyle(dataStyle);
        }

        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }
}

class CommandState{
    String commandAddr, command, memAddr, arguments;
    String ax, bx, cx, dx, ip;
    int Fof, Fdf, Fif, Fsf, Fzf, Faf, Fpf, Fcf;
    String stackState;
}
