## Excelからcsvに変換

### ビルド
```
docker build -t convert-csv .
```

### 実行

#### ディレクトリ指定(中のエクセル全てcsvに変換)
```
docker run -v "$PWD":/convertCSV \
-w /convertCSV  convert-csv \
./gradlew run -Pargs="./convertCsvExamples/excel ./convertCsvExamples/csv"
```

#### ファイル指定
```
docker run -v "$PWD":/convertCSV \
-w /convertCSV  convert-csv \
./gradlew run -Pargs="./convertCsvExamples/excel/変換したいエクセルファイル ./convertCsvExamples/csv"
```

ホストOSの./convertCsvExamples/csvにcsvファイルが出力される。
