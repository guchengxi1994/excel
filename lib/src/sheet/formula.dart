// ignore_for_file: prefer_constructors_over_static_methods, always_declare_return_types, type_annotate_public_apis

part of excel;

class Formula {
  late String _formula;

  Formula._(String formula) {
    _formula = formula;
  }

  /// Helps to initiate a custom formula
  ///```
  ///var my_custom_formula = Formula.custom('=SUM(1,2)');
  ///```
  static Formula custom(String formula) {
    return Formula._(formula);
  }

  /// get Formula
  get formula {
    return _formula;
  }

  @override
  String toString() {
    return _formula;
  }
}
