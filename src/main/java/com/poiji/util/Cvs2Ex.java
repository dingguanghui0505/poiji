package com.poiji.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.*;
import java.util.ArrayList;
import java.util.UUID;

public final class Cvs2Ex {
    private static final Files files = Files.getInstance();

    /*
     * 判断是否是cvs文件
     * */
    public static Boolean isCvs(File file) {
        String fileName = file.getName();
        String filesExtension = files.getExtension(fileName);
        return ".csv".equalsIgnoreCase(filesExtension);
    }


    /**
     * 读取csv 文件内容
     */
    private static void readFile(ArrayList arList, BufferedReader myInput) throws IOException {
        ArrayList al;
        String thisLine ;
        while ((thisLine = myInput.readLine()) != null) {
            al = new ArrayList();
            String strar[] = thisLine.split(",");
            for (int j = 0; j < strar.length; j++) {
                al.add(strar[j]);
            }
            arList.add(al);
        }
    }


    public static File transfromToEx(File file) {

        FileInputStream fis = null;
        ArrayList arList = new ArrayList();
        HSSFWorkbook hwb = null;
        BufferedReader myInput = null;
        try {
            fis = new FileInputStream(file);
            myInput = new BufferedReader(new InputStreamReader(fis,"gbk"));
            arList = new ArrayList();
            readFile(arList, myInput);
            UUID.randomUUID();
            File filOut = new File(UUID.randomUUID().toString() + ".xls");

            hwb = new HSSFWorkbook();
            HSSFSheet sheet = hwb.createSheet("new sheet");
            for (int k = 0; k < arList.size(); k++) {
                ArrayList ardata = (ArrayList) arList.get(k);
                HSSFRow row = sheet.createRow((short) 0 + k);
                transForm(ardata, row);
            }
            hwb.write(filOut);
            return filOut;
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        } finally {
            try {
                hwb.close();
                myInput.close();
                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static FileInputStream transfromToEx(InputStream fis) {
        ArrayList arList = null;
        ArrayList al = null;
        BufferedReader myInput = null;
        FileInputStream inputStream = null;
        HSSFWorkbook hwb = null;
        try {
            myInput = new BufferedReader(new InputStreamReader(fis,"gbk"));
            arList = new ArrayList();
            readFile(arList, myInput);
            File filOut = new File(UUID.randomUUID().toString() + ".xls");


            hwb = new HSSFWorkbook();
            HSSFSheet sheet = hwb.createSheet("new sheet");
            for (int k = 0; k < arList.size(); k++) {
                ArrayList ardata = (ArrayList) arList.get(k);
                HSSFRow row = sheet.createRow((short) 0 + k);
                transForm(ardata, row);
            }
            hwb.write(filOut);
            inputStream = new FileInputStream(filOut);
            filOut.delete();
            return inputStream;
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }



    /**
     * 转换
     */
    private static void transForm(ArrayList ardata, HSSFRow row) {
        for (int p = 0; p < ardata.size(); p++) {
            HSSFCell cell = row.createCell((short) p);
            String data = ardata.get(p).toString();
            if (data.startsWith("=")) {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                data = data.replaceAll("\"", "");
                data = data.replaceAll("=", "");
                cell.setCellValue(data);
            } else if (data.startsWith("\"")) {
                data = data.replaceAll("\"", "");
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(data);
            } else {
                data = data.replaceAll("\"", "");
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(data);
            }
        }
    }


}
