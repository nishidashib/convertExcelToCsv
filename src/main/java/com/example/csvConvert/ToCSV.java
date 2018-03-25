package com.example.csvConvert;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ToCSV {

    private Workbook workbook;
    private ArrayList<ArrayList<String>> csvData;
    private int maxRowWidth;
    private int formattingConvention;
    private DataFormatter formatter;
    private FormulaEvaluator evaluator;
    private String separator;

    private static final String CSV_FILE_EXTENSION = ".csv";
    private static final String DEFAULT_SEPARATOR = ",";

    public static final int EXCEL_STYLE_ESCAPING = 0;

    public static final int UNIX_STYLE_ESCAPING = 1;

    /**
     * @param strSource エクセルが格納されているディレクトリ。またはエクセルファイル自身
     * @param strDestination csvファイルが格納されるディレクトリ
     * @throws FileNotFoundException .
     * @throws IOException
     * @throws IllegalArgumentException (存在しないファイルやディレクトリを指定したとき)
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException (xlsxがパースできなかった時)
     * @throws EncryptedDocumentException
     */
    public void convertExcelToCSV(String strSource, String strDestination)
            throws FileNotFoundException, IOException, IllegalArgumentException, InvalidFormatException, EncryptedDocumentException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        this.convertExcelToCSV(strSource, strDestination, ToCSV.DEFAULT_SEPARATOR, ToCSV.EXCEL_STYLE_ESCAPING);
    }

    /**
     *
     * @param strSource エクセルが格納されているディレクトリ。またはエクセルファイル自身
     * @param strDestination csvファイルが格納されるディレクトリ
     * @param separator csvの区切り文字
     * @throws FileNotFoundException .
     * @throws IOException
     * @throws IllegalArgumentException (存在しないファイルやディレクトリを指定したとき)
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException  (xlsxがパースできなかった時)
     * @throws EncryptedDocumentException
     */
    public void convertExcelToCSV(String strSource, String strDestination, String separator)
                       throws FileNotFoundException, IOException, IllegalArgumentException, InvalidFormatException, EncryptedDocumentException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        this.convertExcelToCSV(strSource, strDestination, separator, ToCSV.EXCEL_STYLE_ESCAPING);
    }

    /**
     * @param strSource エクセルが格納されているディレクトリ。またはエクセルファイル自身
     * @param strDestination csvファイルが格納されるディレクトリ
     * @param separator csvの区切り文字
     * @param formattingConvention エスケープが必要な文字を、エクセルの書式かunixの書式でエスケープするかを選択
     * @throws FileNotFoundException
     * @throws IOException
     * @throws IllegalArgumentException (存在しないファイルやディレクトリを指定したとき)
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException (xlsxがパースできなかった時)
     * @throws EncryptedDocumentException
     */
    public void convertExcelToCSV(String strSource, String strDestination, String separator, int formattingConvention)
                       throws FileNotFoundException, IOException, IllegalArgumentException, InvalidFormatException, EncryptedDocumentException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        File source = new File(strSource);
        File destination = new File(strDestination);
        File[] filesList;
        String destinationFilename;

        if(!source.exists()) {
            throw new IllegalArgumentException("Excelファイル、もしくは格納されているディレクトリが見つかりません。");
        }

        if(!destination.exists()) {
            throw new IllegalArgumentException("CSVを格納するディレクトリが見つかりません。");
        }
        if(!destination.isDirectory()) {
                throw new IllegalArgumentException("ファイルではなく、ディレクトリを指定して下さい。");
        }

        if(formattingConvention != ToCSV.EXCEL_STYLE_ESCAPING && formattingConvention != ToCSV.UNIX_STYLE_ESCAPING) {
                throw new IllegalArgumentException("渡されたパラメータが不正です。");
        }

        this.separator = separator;
        this.formattingConvention = formattingConvention;

        if(source.isDirectory()) {
            filesList = source.listFiles(new ExcelFilenameFilter());
        } else {
                filesList = new File[]{source};
        }

        // 現状の欠点として、ファイル名が同じxlsとxlsxがある場合は,(ex aaa.xlsとaaa.xlsx)
        // 変換されたcsvのファイル名が同じなので、一方のcsvがもう一方のcsvファイルを上書きしてしまう。
        // この問題はとりあえず放置。
        if (filesList != null) {
            for(File excelFile : filesList) {
                this.openWorkbook(excelFile);
                this.convertToCSV();
                destinationFilename = excelFile.getName();
                destinationFilename = destinationFilename.substring(0,destinationFilename.lastIndexOf(".")) + ToCSV.CSV_FILE_EXTENSION;
                    this.saveCSVFile(new File(destination, destinationFilename)
                );
            }
        }
    }

    /**
     * @param file エクセルファイルオブジェクト(拡張子はxls,またはxlsx)
     * @throws FileNotFoundException
     * @throws IOException
     * @throws InvalidFormatException (xlsxがパースできなかった時)
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException
     * @throws EncryptedDocumentException
     */
    private void openWorkbook(File file) throws FileNotFoundException, IOException, InvalidFormatException, EncryptedDocumentException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        FileInputStream fileInputStream = null;
        try {
            System.out.println("Opening workbook [" + file.getName() + "]");

            fileInputStream = new FileInputStream(file);

            this.workbook = WorkbookFactory.create(fileInputStream);
            // セルに計算式が仕込まれてても、値を取得できるようにする処理
            this.evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
            // セルに書式が設定されていても、値を表示できるようにする処理
            this.formatter = new DataFormatter(true);
        }
        finally {
            if(fileInputStream != null) {
                fileInputStream.close();
            }
        }
    }

    private void convertToCSV() {
        Sheet sheet;
        Row row;
        int lastRowNum;
        this.csvData = new ArrayList<>();

        System.out.println("Converting files contents to CSV format.");

        int numSheets = this.workbook.getNumberOfSheets();

        for(int i = 0; i < numSheets; i++) {
            sheet = this.workbook.getSheetAt(i);
            if(sheet.getPhysicalNumberOfRows() > 0) {
                lastRowNum = sheet.getLastRowNum();
                for(int j = 0; j <= lastRowNum; j++) {
                    row = sheet.getRow(j);
                    this.rowToCSV(row);
                }
            }
        }
    }

    /**
     * @param file csvファイルをハンドルするファイルオブジェクト
     * @throws FileNotFoundException(ファイルが存在しない時)
     * @throws IOException(ファイルシステムでの入出力エラー)
     */
    private void saveCSVFile(File file) throws FileNotFoundException, IOException {
        BufferedWriter bw = null;
        ArrayList<String> line;
        StringBuffer buffer;
        String csvLineElement;
        try {
            System.out.println("Saving the CSV file [" + file.getName() + "]");
            OutputStreamWriter osw  = new OutputStreamWriter(new FileOutputStream(file), "Shift-JIS");
            bw = new BufferedWriter(osw);

            for(int i = 0; i < this.csvData.size(); i++) {
                buffer = new StringBuffer();
                line = this.csvData.get(i);
                for(int j = 0; j < this.maxRowWidth; j++) {
                    if(line.size() > j) {
                        csvLineElement = line.get(j);
                        if(csvLineElement != null) {
                            buffer.append(this.escapeEmbeddedCharacters(csvLineElement));
                        }
                    }
                    if(j < (this.maxRowWidth - 1)) {
                        buffer.append(this.separator);
                    }
                }

                bw.write(buffer.toString().trim());

                if(i < (this.csvData.size() - 1)) {
                    bw.newLine();
                }
            }
        }
        finally {
            if(bw != null) {
                bw.flush();
                bw.close();
            }
        }
    }

    /**
     * @param row Apache Poiでエクセルの行を扱うようにするための行オブジェクト
     */
    private void rowToCSV(Row row) {
        Cell cell;
        int lastCellNum;
        ArrayList<String> csvLine = new ArrayList<>();

        if(row != null) {
            lastCellNum = row.getLastCellNum();
            for(int i = 0; i <= lastCellNum; i++) {
                cell = row.getCell(i);
                if(cell == null) {
                    csvLine.add("");
                }
                else {
                    // if(cell.getCellType() != CellType.FORMULA) {
                    if(cell.getCellTypeEnum() != CellType.FORMULA) {
                        csvLine.add(this.formatter.formatCellValue(cell));
                    }
                    else {
                        csvLine.add(this.formatter.formatCellValue(cell, this.evaluator));
                    }
                }
            }
            if(lastCellNum > this.maxRowWidth) {
                this.maxRowWidth = lastCellNum;
            }
        }
        this.csvData.add(csvLine);
    }

    /**
     * @param field セルの値に対応したStringオブジェクト
     * @return セルの中の値に"などがあったら、正しくエスケープされた後のStringオブジェクト
     */
    private String escapeEmbeddedCharacters(String field) {
        StringBuffer buffer;

        if(this.formattingConvention == ToCSV.EXCEL_STYLE_ESCAPING) {
            // セルの値に"やセパレータ(デフォルトは",") 行末文字(EOL)が存在した場合は、excelの規則に従ってエスケープする
            // セル内の「"」エスケープ。ex) 彼は"了解"と伝えた → "彼は""了解""と伝えた"
            if(field.contains("\"")) {
                buffer = new StringBuffer(field.replaceAll("\"", "\"\""));
                buffer.insert(0, "\"");
                buffer.append("\"");
            }
            // セルの値にセパレータ(デフォルトは",")がある場合はエスケープ ex) 2,434,886,692 → "2,434,886,692"
            // セルの値に改行がある場合は1行で表示。ex) 航空機燃料(ここで改行)譲与税  → "航空機燃料譲与税"
            else {
                buffer = new StringBuffer(field);
                if((buffer.indexOf(this.separator)) > -1 || (buffer.indexOf("\n")) > -1) {
                    buffer = new StringBuffer(field.replaceAll("\\n", ""));
                    buffer.insert(0, "\"");
                    buffer.append("\"");
                }
            }
            return(buffer.toString().trim());
        } else if (this.formattingConvention == ToCSV.UNIX_STYLE_ESCAPING) {
            // セルの値に改行コードや区切り文字を含む場合は、unixの規則に従ってエスケープする
            if(field.contains(this.separator)) {
                field = field.replaceAll(this.separator, ("\\\\" + this.separator));
            }
            if(field.contains("\n")) {
                field = field.replaceAll("\n", "\\\\\n");
            }
            return field;
        }
        return field;
    }

    class ExcelFilenameFilter implements FilenameFilter {

        /**
         * @param file Fileオブジェクト
         * @param name ファイル名
         * @return 拡張子がxls,またはxlsxであるファイル名
         */
        @Override
        public boolean accept(File file, String name) {
            return(name.endsWith(".xls") || name.endsWith(".xlsx"));
        }
    }
}
