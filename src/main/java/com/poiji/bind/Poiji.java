package com.poiji.bind;

import com.poiji.bind.mapping.UnmarshallerHelper;
import com.poiji.exception.IllegalCastException;
import com.poiji.exception.InvalidExcelFileExtension;
import com.poiji.exception.PoijiExcelType;
import com.poiji.exception.PoijiException;
import com.poiji.option.PoijiOptions;
import com.poiji.option.PoijiOptions.PoijiOptionsBuilder;
import com.poiji.util.Cvs2Ex;
import com.poiji.util.Files;

import java.io.*;
import java.util.List;
import java.util.Objects;

import static com.poiji.util.PoijiConstants.XLSX_EXTENSION;
import static com.poiji.util.PoijiConstants.XLS_EXTENSION;

/**
 * The entry point of the mapping process.
 * <p>
 * Example:
 * <pre>
 * List<Employee> employees = Poiji.fromExcel(new File("employees.xls"), Employee.class);
 * employees.size();
 * // 3
 * Employee firstEmployee = employees.get(0);
 * // Employee{employeeId=123923, name='Joe', surname='Doe', age=30, single=true, birthday='4/9/1987'}
 * </pre>
 * <p>
 * Created by hakan on 16/01/2017.
 */
public final class Poiji {

    private static final Files files = Files.getInstance();

    private Poiji() {
    }

    /**
     * converts excel rows into a list of objects
     *
     * @param file excel file ending with .xls or .xlsx.
     * @param type type of the root object.
     * @param <T>  type of the root object.
     * @return the newly created a list of objects
     * @throws PoijiException            if an internal exception occurs during the mapping process.
     * @throws InvalidExcelFileExtension if the specified excel file extension is invalid.
     * @throws IllegalCastException      if this Field object is enforcing Java language access control and the underlying field is either inaccessible or final.
     * @see Poiji#fromExcel(File, Class, PoijiOptions)
     */
    public static synchronized <T> List<T> fromExcel(final File file, final Class<T> type) {

        //--------------<<
        Boolean cvs = Cvs2Ex.isCvs(file);
        if (cvs) {
            File file1Ex = Cvs2Ex.transfromToEx(file);
            final Unmarshaller unmarshaller = deserializer(file1Ex, PoijiOptionsBuilder.settings().build());
            return unmarshaller.unmarshal(type);
        }
        //--------------<<
        final Unmarshaller unmarshaller = deserializer(file, PoijiOptionsBuilder.settings().build());
        return unmarshaller.unmarshal(type);
    }

    /**
     * converts excel rows into a list of objects
     *
     * @param inputStream excel file stream
     * @param excelType   type of the excel file, xls or xlsx
     * @param type        type of the root object.
     * @param <T>         type of the root object.
     * @return the newly created a list of objects
     * @throws PoijiException            if an internal exception occurs during the mapping process.
     * @throws InvalidExcelFileExtension if the specified excel file extension is invalid.
     * @throws IllegalCastException      if this Field object is enforcing Java language access control and the underlying field is either inaccessible or final.
     * @see Poiji#fromExcel(File, Class, PoijiOptions)
     */
    public static synchronized <T> List<T> fromExcel(final InputStream inputStream,
                                                     PoijiExcelType excelType,
                                                     final Class<T> type) {
        Objects.requireNonNull(excelType);
        //--------------------------------------------<<
        if(excelType == PoijiExcelType.CSV) {

            PoijiInputStream poijiInputStream = new PoijiInputStream<>(Cvs2Ex.transfromToEx(inputStream));
            Unmarshaller unmarshaller = UnmarshallerHelper.HSSFInstance(poijiInputStream, PoijiOptionsBuilder.settings().build());
            return unmarshaller.unmarshal(type);
        }
        //-------------------------------------------<<

        final Unmarshaller unmarshaller = deserializer(inputStream, excelType, PoijiOptionsBuilder.settings().build());
        return unmarshaller.unmarshal(type);
    }

    /**
     * converts excel rows into a list of objects
     *
     * @param file    excel file ending with .xls or .xlsx.
     * @param type    type of the root object.
     * @param <T>     type of the root object.
     * @param options specifies to change the default behaviour of the poiji.
     * @return the newly created a list of objects
     * @throws PoijiException            if an internal exception occurs during the mapping process.
     * @throws InvalidExcelFileExtension if the specified excel file extension is invalid.
     * @throws IllegalCastException      if this Field object is enforcing Java language access control and the underlying field is either inaccessible or final.
     * @see Poiji#fromExcel(File, Class)
     */
    public static synchronized <T> List<T> fromExcel(final File file, final Class<T> type, final PoijiOptions options) {
        //--------------<<
        Boolean cvs = Cvs2Ex.isCvs(file);
        if(cvs){
            File file1Ex = Cvs2Ex.transfromToEx(file);
            final Unmarshaller unmarshaller = deserializer(file1Ex, options);
            return unmarshaller.unmarshal(type);
        }
        //------------<<
        final Unmarshaller unmarshaller = deserializer(file, options);
        return unmarshaller.unmarshal(type);
    }

    /**
     * converts excel rows into a list of objects
     *
     * @param inputStream excel file stream
     * @param excelType   type of the excel file, xls or xlsx
     * @param type        type of the root object.
     * @param <T>         type of the root object.
     * @param options     specifies to change the default behaviour of the poiji.
     * @return the newly created a list of objects
     * @throws PoijiException            if an internal exception occurs during the mapping process.
     * @throws InvalidExcelFileExtension if the specified excel file extension is invalid.
     * @throws IllegalCastException      if this Field object is enforcing Java language access control and the underlying field is either inaccessible or final.
     * @see Poiji#fromExcel(File, Class)
     */
    public static synchronized <T> List<T> fromExcel(final InputStream inputStream,
                                                     final PoijiExcelType excelType,
                                                     final Class<T> type,
                                                     final PoijiOptions options) {
        Objects.requireNonNull(excelType);

        if(excelType == PoijiExcelType.CSV) {

            PoijiInputStream poijiInputStream = new PoijiInputStream<>(Cvs2Ex.transfromToEx(inputStream));
            Unmarshaller unmarshaller = UnmarshallerHelper.HSSFInstance(poijiInputStream, options);
            return unmarshaller.unmarshal(type);
        }

        final Unmarshaller unmarshaller = deserializer(inputStream, excelType, options);
        return unmarshaller.unmarshal(type);
    }

    @SuppressWarnings("unchecked")
    private static Unmarshaller deserializer(final File file, final PoijiOptions options) {
        final PoijiFile poijiFile = new PoijiFile(file);

        String extension = files.getExtension(file.getName());

        if (XLS_EXTENSION.equals(extension)) {
            return UnmarshallerHelper.HSSFInstance(poijiFile, options);
        } else if (XLSX_EXTENSION.equals(extension)) {
            return UnmarshallerHelper.XSSFInstance(poijiFile, options);
        }
        else {
            throw new InvalidExcelFileExtension("Invalid file extension (" + extension + "), excepted .xls or .xlsx");
        }
    }

    @SuppressWarnings("unchecked")
    private static Unmarshaller deserializer(final InputStream inputStream, PoijiExcelType excelType, final PoijiOptions options) {
        final PoijiInputStream poijiInputStream = new PoijiInputStream<>(inputStream);

        if (excelType == PoijiExcelType.XLS) {
            return UnmarshallerHelper.HSSFInstance(poijiInputStream, options);
        } else if (excelType == PoijiExcelType.XLSX) {
            return UnmarshallerHelper.XSSFInstance(poijiInputStream, options);
        } else {
            throw new InvalidExcelFileExtension("Invalid file extension (" + excelType + "), excepted .xls or .xlsx");
        }
    }






}
