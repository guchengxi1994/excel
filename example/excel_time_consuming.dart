// ignore_for_file: avoid_print

import 'dart:io';

import 'package:excel/excel.dart';
// ignore: depend_on_referenced_packages
import 'package:path/path.dart';

void main(List<String> args) {
  final Stopwatch stopwatch = Stopwatch()..start();

  final Excel excel = Excel.createExcel();
  final Sheet sh = excel['Sheet1'];
  for (int i = 0; i < 8; i++) {
    sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).value =
        'Col $i';
    //sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).cellStyle =CellStyle(bold: true);
  }
  for (int row = 1; row < 9000; row++) {
    for (int col = 0; col < 80; col++) {
      sh
          .cell(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col))
          .value = '$row$col value';
    }
  }
  print('Generating executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  final fileBytes = excel.encode();

  print('Encoding executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  if (fileBytes != null) {
    File(join("/Users/kawal/Desktop/r2.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
  print('Downloaded executed in ${stopwatch.elapsed}');
}
