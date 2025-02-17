// ignore_for_file: unsafe_html

import 'dart:html' as html;

// A wrapper to save the
// ignore: avoid_classes_with_only_static_members
class SavingHelper {
  static List<int>? saveFile(List<int>? val, String fileName) {
    final blob = html.Blob([val]);
    final url = html.Url.createObjectUrlFromBlob(blob);
    final anchor = html.document.createElement('a') as html.AnchorElement
      ..href = url
      ..style.display = 'none'
      ..download = fileName;
    html.document.body?.children.add(anchor);

    // download the file
    anchor.click();
    // cleanup
    html.document.body?.children.remove(anchor);
    html.Url.revokeObjectUrl(url);
    return val;
  }
}
