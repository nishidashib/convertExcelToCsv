package com.example.csvConvert;

import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class App {

	public static void main(String[] args) {
        ToCSV converter;
        boolean converted = true;
        long startTime = System.currentTimeMillis();
        try {
            converter = new ToCSV();
            if(args.length == 2) {
                converter.convertExcelToCSV(args[0], args[1]);
            }
            else if(args.length == 3){
                converter.convertExcelToCSV(args[0], args[1], args[2]);
            }
            else if(args.length == 4) {
                converter.convertExcelToCSV(args[0], args[1], args[2], Integer.parseInt(args[3]));
            }
            else {
                System.out.println("Usage: java ToCSV [Source File/Folder] " +
                    "[Destination Folder] [Separator] [Formatting Convention]\n" +
                    "\tSource File/Folder\tこの引数には、単一のExcel、または1つ以上のExcelを含むディレクトリ名を\n" +
                    "\t\t\t\tフルパスで指定する必要があります。\n" +
                    "\tDestination Folder\tCSVファイルを書き込むディレクトリ名をフルパスで指定する必要があります。\n" +
                    "\t\t\t\tこのディレクトリはあらかじめ作成しておいて下さい。\n" +
                    "\tSeparator\tオブション。セルの区切り文字。デフォルトはカンマ。\n" +
                    "\tFormatting Convention\tオプション。0の時は エスケープが必要な文字を、Excelの書式でエスケープ。\n" +
                    "\t\t\t\t1の時は エスケープが必要な文字を、Unixの書式でエスケープ。\n" +
                    "\t\t\t\tデフォルトはExcel。");
                converted = false;
            }
        }
        catch(Exception ex) {
            System.out.println("Caught an: " + ex.getClass().getName());
            System.out.println("Message: " + ex.getMessage());
            System.out.println("Stacktrace follows:.....");
            ex.printStackTrace(System.out);
            converted = false;
        }

        if (converted) {
            System.out.println("Conversion took " +
                  (int)((System.currentTimeMillis() - startTime)/1000) + " seconds");
        }
	}
}