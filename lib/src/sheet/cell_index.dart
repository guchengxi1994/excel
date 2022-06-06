// ignore_for_file: prefer_constructors_over_static_methods

part of excel;

// ignore: must_be_immutable
class CellIndex extends Equatable {
  CellIndex._({int? col, int? row}) {
    // ignore: prefer_asserts_in_initializer_lists
    assert(col != null && row != null);
    _columnIndex = col!;
    _rowIndex = row!;
  }

  ///
  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///```
  static CellIndex indexByColumnRow({int? columnIndex, int? rowIndex}) {
    assert(columnIndex != null && rowIndex != null);
    return CellIndex._(col: columnIndex, row: rowIndex);
  }

  ///
  ///```
  /// CellIndex.indexByColumnRow('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByColumnRow('A2'); // columnIndex: 0, rowIndex: 1
  ///```
  static CellIndex indexByString(String cellIndex) {
    final List<int> li = _cellCoordsFromCellId(cellIndex);
    return CellIndex._(row: li[0], col: li[1]);
  }

  /// Avoid using it as it is very process expensive function.
  ///
  /// ```
  /// var cellIndex = CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 );
  /// var cell = cellIndex.cellId; // A1
  String get cellId {
    return getCellId(columnIndex, rowIndex);
  }

  late int _rowIndex;

  int get rowIndex {
    return _rowIndex;
  }

  late int _columnIndex;

  int get columnIndex {
    return _columnIndex;
  }

  @override
  List<Object?> get props => [_rowIndex, _columnIndex];
}
